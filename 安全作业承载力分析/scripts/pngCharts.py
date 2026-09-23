# -*- coding: utf-8 -*-
"""PNG 图表渲染 + 插入成品（原 renderReportCharts + applyCharts 合并）。

子命令：
  python pngCharts.py render  ...原 renderReportCharts 参数...
  python pngCharts.py apply   ...原 applyCharts 参数...
无子命令时：若含 --doc 走 apply，否则走 render（兼容旧调用）。
"""
from __future__ import annotations

# ----- render (原 renderReportCharts) -----

import argparse
import json
import logging
import os
import sys
from typing import Dict, List, Optional, Tuple

from capacity.config import (CONFIG_SECTION_INPUT, DEFAULT_INPUT_CONFIG,
                             apply_formula_config, load_config)
from capacity.constants import (InputDataError, FileAccessError, FileContentError)
from capacity.excelLoader import read_xls_dataset
from capacity.calculator import (period_manage_caps_from_days,
                                 period_unit_caps_from_daily)

# ---- 依赖检查 ----
try:
    import matplotlib
    matplotlib.use("Agg")  # 无界面
    import matplotlib.pyplot as plt
    from matplotlib import font_manager
except ImportError as exc:  # pragma: no cover
    print(json.dumps({"ok": False, "errorCode": "DEP_MISSING",
                      "message": f"缺少绘图依赖 matplotlib：{exc}"},
                     ensure_ascii=False))
    sys.exit(1)

LOGGER = logging.getLogger("pngCharts.render")

EXIT_OK, EXIT_UNEXPECTED, EXIT_USAGE, EXIT_INPUT, EXIT_IO = 0, 1, 2, 3, 4

# 图表尺寸（对应 .doc 模板图框宽 441pt≈155.6mm≈6.1in；A4 正文宽）
# 注意：模板图框是横向长条（441×160~188pt，宽高比≈2.5），图表必须用横向长条版式，
# 否则插入 Word 时按宽等比缩放会严重超高、撑破页面。
FIG_W_INCH = 6.15
FIG_H_INCH = 2.00
HBAR_MAX_H_INCH = 3.20  # 紧凑封顶，避免半页空白
DPI = 300

# 配色：深蓝主色 + 承载力预警分级（v1.6.4 起重载不用黄色系，全图 深蓝→浅蓝→红→深红，
# 去除黄色，柱状清晰不刺眼）
C_MAIN = "#1457A8"      # 深蓝（轻载/适度 统一主色）
C_HEAVY = "#5A8FD4"     # 浅蓝 · 重载 75~90%（深蓝同谱，替代原橙色/黄色）
C_FULL = "#E3483F"      # 红 · 满载 >90%
C_OVER = "#B02418"      # 深红 · 超满载 >100%
C_GRID = "#D8DEE8"      # 网格浅灰
C_TXT = "#2B2B2B"       # 正文深灰
C_NOTE = "#8A8A8A"      # 脚注浅灰

FONT_NAME = "SimSun"
CHART_FS = 12.0  # 小四


def _setup_font() -> None:
    """注册中文字体（宋体优先，图内小四）。

    Windows 有 SimSun 时用宋体。Linux 无宋体时回退 Noto Serif CJK SC
    （fonts-noto-cjk）或文泉驿正黑，避免中文变方框。
    """
    cands = ("SimSun", "NSimSun", "宋体", "Microsoft YaHei",
             "Noto Serif CJK SC", "Noto Sans CJK SC",
             "WenQuanYi Zen Hei", "WenQuanYi Micro Hei")
    global FONT_NAME
    for path in (
            "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc",
            "/usr/share/fonts/opentype/noto/NotoSerifCJK-Bold.ttc",
            "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"):
        if os.path.isfile(path):
            try:
                font_manager.fontManager.addfont(path)
            except Exception:
                pass
    names = {f.name for f in font_manager.fontManager.ttflist}
    chosen = None
    for cand in cands:
        if cand in names:
            chosen = cand
            break
    if chosen is None:
        for path in (
                "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc",
                "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"):
            if os.path.isfile(path):
                font_manager.fontManager.addfont(path)
                chosen = font_manager.FontProperties(fname=path).get_name()
                break
    if chosen:
        FONT_NAME = chosen
        plt.rcParams["font.family"] = chosen
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["font.size"] = CHART_FS


def bar_color(v: float) -> str:
    if v > 100:
        return C_OVER
    if v > 90:
        return C_FULL
    if v > 75:
        return C_HEAVY
    return C_MAIN


def load_matrix(data_path: str) -> Dict[str, Dict[str, Dict]]:
    """读取 calcCapacity 输出 JSON 的 matrix：{date: {unit: {承载力, 预警}}}。"""
    with open(data_path, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    matrix = data.get("matrix")
    if not matrix:
        raise ValueError("data 中无 matrix（需组织架构映射.是否启用=true 后重跑）")
    return matrix


def unit_period_mean(matrix: Dict, lo: str, hi: str,
                     units: Optional[List[str]] = None) -> List[Tuple[str, float]]:
    """区间单位承载力：Σ自然日日承载力 / 工作日天数，返回 [(单位, 周期值)]。"""
    import datetime
    lo_d = datetime.date.fromisoformat(lo)
    hi_d = datetime.date.fromisoformat(hi)
    cap_by_day: Dict[str, Dict[str, float]] = {}
    for dkey, umap in matrix.items():
        cap_by_day[dkey] = {u: float((info or {}).get("承载力") or 0)
                            for u, info in umap.items()}
    caps = period_unit_caps_from_daily(cap_by_day, lo_d, hi_d)
    keys = units if units is not None else list(caps.keys())
    out = [(u, caps[u]) for u in keys if u in caps]
    out.sort(key=lambda kv: kv[1], reverse=True)
    return out


_PLAN_TYPE_CACHE: Dict[Tuple[str, str], object] = {}


def plan_type_count(xls_path: str, config_path: str, lo: str,
                    hi: str) -> List[Tuple[str, int]]:
    """按作业类型统计 lo..hi 区间计划项数（.xls/.xlsm/.xlsx 双引擎）。

    同进程内按源文件+配置缓存数据集，避免月报多张图重复读 Excel。
    """
    import datetime
    from collections import Counter

    key = (os.path.abspath(xls_path), os.path.abspath(config_path))
    dataset = _PLAN_TYPE_CACHE.get(key)
    if dataset is None:
        cfg = json.load(open(config_path, "r", encoding="utf-8"))
        input_cfg = cfg.get(CONFIG_SECTION_INPUT) or dict(DEFAULT_INPUT_CONFIG)
        input_cfg = dict(input_cfg)
        manage = dict(input_cfg.get("管理承载力") or {})
        manage["是否启用"] = False
        input_cfg["管理承载力"] = manage
        pa = dict(input_cfg.get("人员档案") or {})
        pa["是否加载"] = False
        input_cfg["人员档案"] = pa
        dataset = read_xls_dataset(xls_path, input_cfg)
        _PLAN_TYPE_CACHE[key] = dataset
    d_lo = datetime.date.fromisoformat(lo)
    d_hi = datetime.date.fromisoformat(hi)
    cnt: Counter = Counter((p.plan_type or "其他") for p in dataset.plans
                           if d_lo <= p.start <= d_hi)
    return sorted(cnt.items(), key=lambda kv: -kv[1])


def render_capacity_chart(data: List[Tuple[str, float]], out_path: str,
                          label: str, note: Optional[str] = None) -> None:
    """各单位作业承载力竖柱（横轴=单位、纵轴=%；标签整数不带单位）。"""
    _render_vbar_chart(
        data, out_path, label, "%",
        value_fmt=lambda v: f"{int(round(v))}", ref_lines=(75, 90),
        pct_axis=True,
        note=note or "注：承载力=Σ自然日日承载力÷工作日天数；>100 表示超满负荷。")


def _ptype_colors(n: int) -> List[str]:
    """作业类型图配色：类少统一深蓝主色，类多（>6）按排序由深到浅的蓝渐变。"""
    if n <= 6:
        return [C_MAIN] * n
    cmap = plt.get_cmap("Blues")
    return ["#%02x%02x%02x" % tuple(int(round(255 * c)) for c in cmap(0.90 - 0.25 * (i / (n - 1)))[:3])
            for i in range(n)]


def render_ptype_chart(data: List[Tuple[str, int]], out_path: str,
                       label: str) -> None:
    """各作业类型日计划数量竖柱（横轴=类型、纵轴=项；标签整数）。"""
    n = len(data)
    colors = _ptype_colors(n)
    _render_vbar_chart(
        data, out_path, label, "项",
        value_fmt=lambda v: f"{int(v)}",
        colors=colors, over_threshold=None,
        note="注：按作业类型统计区间内日计划项数。")


def _wrap_tick(s: str, n_cats: int) -> str:
    """分类轴标签：类多按 2 字换行，禁止倾斜。"""
    text = str(s)
    width = 2 if n_cats >= 10 else (3 if n_cats >= 6 else 8)
    if len(text) <= width:
        return text
    return "\n".join(text[i:i + width] for i in range(0, len(text), width))


def _render_vbar_chart(data: List[Tuple[str, float]], out_path: str,
                       label: str, ylabel: str, value_fmt,
                       ref_lines: Tuple[float, ...] = (), note: str = "",
                       colors: Optional[List[str]] = None,
                       over_threshold: Optional[float] = 90.0,
                       pct_axis: bool = False) -> None:
    """国网竖柱：横轴类别、纵轴数值；无外框；标签整数；单位只在纵轴标题。"""
    n = len(data)
    names = [_wrap_tick(t, n) for t, _ in data]
    vals = [float(v) for _, v in data]
    if colors is None:
        colors = [bar_color(v) for v in vals]
    fig_h = min(HBAR_MAX_H_INCH, max(FIG_H_INCH, 2.05))
    fig, ax = plt.subplots(figsize=(FIG_W_INCH, fig_h), dpi=DPI)
    x = list(range(n))
    ax.bar(x, vals, color=colors, width=0.62, edgecolor="white",
           linewidth=0.4, zorder=3)
    peak = max(vals) if vals else 0.0
    if pct_axis:
        ymax = 125.0 if peak <= 125.0 else (int((peak - 1e-9) // 25) + 1) * 25.0
    else:
        ymax = max(peak * 1.18, 1.0) if vals else 1.0
    for xi, v in zip(x, vals):
        over = over_threshold is not None and v > over_threshold
        ax.text(xi, v + ymax * 0.012, value_fmt(v),
                ha="center", va="bottom", fontsize=CHART_FS - 1,
                color=C_OVER if over else C_TXT,
                fontweight="bold" if over else "normal",
                fontname=FONT_NAME)
    for rv in ref_lines:
        rc = C_HEAVY if rv < 80 else C_FULL
        ax.axhline(rv, color=rc, lw=0.8, ls="--", alpha=0.7, zorder=2)
        ax.text(n - 0.5, rv, f"{int(rv)}", ha="left", va="bottom",
                fontsize=10, color=rc, clip_on=False, fontname=FONT_NAME)
    ax.set_xticks(x)
    tick_fs = 9.0 if n >= 10 else CHART_FS
    ax.set_xticklabels(names, fontsize=tick_fs, color=C_TXT, fontname=FONT_NAME)
    ax.set_xlim(-0.55, n - 0.45)
    ax.set_ylim(0, ymax)
    ax.grid(axis="y", color=C_GRID, lw=0.6, zorder=0)
    ax.set_axisbelow(True)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    ax.spines["left"].set_color(C_GRID)
    ax.spines["bottom"].set_color(C_GRID)
    ax.tick_params(axis="both", length=0, labelsize=tick_fs)
    ax.set_ylabel(ylabel, fontsize=CHART_FS, color=C_TXT, fontname=FONT_NAME)
    fig.suptitle(label, fontsize=CHART_FS, fontweight="bold", color=C_MAIN,
                 y=0.99, x=0.5, fontname=FONT_NAME)
    if note:
        fig.text(0.5, 0.01, note, fontsize=9, color=C_NOTE, ha="center",
                 fontname=FONT_NAME)
    fig.tight_layout(rect=[0, 0.04, 1, 0.96])
    fig.savefig(out_path, dpi=DPI, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def render_manage_chart(data_path: str, lo: str, hi: str, label: str,
                        out_path: str) -> None:
    """管理承载力竖柱（横轴=单位、纵轴=%）。"""
    import datetime
    with open(data_path, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    days = (data.get("management") or {}).get("days") or {}
    lo_d = datetime.date.fromisoformat(lo)
    hi_d = datetime.date.fromisoformat(hi)
    caps = period_manage_caps_from_days(days, lo_d, hi_d)
    if not caps:
        raise ValueError(f"区间 {lo}~{hi} 无可用的管理承载力数据（management.days 为空）")
    rows: List[Tuple[str, float]] = sorted(caps.items(), key=lambda kv: kv[1], reverse=True)
    title = label.replace("各单位承载力", "").strip(" -　") + "管理承载力"
    _render_vbar_chart(
        rows, out_path, title, "%",
        value_fmt=lambda v: f"{int(round(v))}", ref_lines=(75, 90),
        pct_axis=True,
        note="注：F_管理=Σ自然日实需人天÷(在册人数×工作日天数)。")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="renderReportCharts.py",
        description="承载力量化报告图表生成器（各单位承载力条形图 + 作业类型分布图）")
    p.add_argument("--data", required=True, help="calcCapacity 输出 JSON（含 matrix）")
    p.add_argument("--xls", required=True, help="作业计划源 .xls/.xlsx")
    p.add_argument("--config", required=True, help="capacity_config.json 路径")
    p.add_argument("--lo", default=None, help="区间起始日期 YYYY-MM-DD")
    p.add_argument("--hi", default=None, help="区间结束日期 YYYY-MM-DD")
    p.add_argument("--label", required=True, help="图表标题（周期描述，如『2026年9月』）")
    p.add_argument("--prefix", required=True, help="输出文件名前缀")
    p.add_argument("--out-dir", required=True, help="输出目录")
    p.add_argument("--no-ptype", action="store_true",
                   help="跳过作业类型分布图（仅生成承载力图）")
    p.add_argument("--mgmt", action="store_true",
                   help="额外生成管理承载力各单位横向条形图（{prefix}_管理承载力.png，读 calc 的 management.days）")
    return p


def _fail_exit(error_code: str, message: str, code: int) -> int:
    """统一失败响应：错误码+信息写 stderr 与 stdout(JSON)，返回退出码。"""
    LOGGER.error(message)
    print(json.dumps({"ok": False, "errorCode": error_code, "message": message},
                     ensure_ascii=False))
    return code


def render_main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    _setup_font()
    try:
        apply_formula_config(load_config(args.config))
        matrix = load_matrix(args.data)
        if not args.lo or not args.hi:
            return _fail_exit("USAGE", "--lo/--hi 必填", EXIT_USAGE)
        cap = unit_period_mean(matrix, args.lo, args.hi)
        if not cap:
            raise ValueError(f"区间 {args.lo}~{args.hi} 无单位承载力数据")
        os.makedirs(args.out_dir, exist_ok=True)
        cap_path = os.path.join(args.out_dir, f"{args.prefix}_单位承载力.png")
        render_capacity_chart(cap, cap_path, args.label)
        outputs = [cap_path]
        if not args.no_ptype:
            ptypes = plan_type_count(args.xls, args.config, args.lo, args.hi)
            pt_path = os.path.join(args.out_dir, f"{args.prefix}_作业类型.png")
            render_ptype_chart(ptypes, pt_path, args.label + " · 各作业类型日计划数量")
            outputs.append(pt_path)
        if args.mgmt:
            mgmt_path = os.path.join(args.out_dir, f"{args.prefix}_管理承载力.png")
            try:
                render_manage_chart(args.data, args.lo, args.hi, args.label, mgmt_path)
                outputs.append(mgmt_path)
            except ValueError as exc:
                # 区间内无管理承载力数据（如某周无计划）：跳过不产图、不中断流水线；
                # applyCharts 见 png 不存在会自动跳过对应评述段的插图。
                LOGGER.warning("跳过管理承载力图（区间无数据）：%s", exc)
        print(json.dumps({"ok": True, "range": [args.lo, args.hi], "label": args.label,
                          "outputs": outputs}, ensure_ascii=False))
        return EXIT_OK
    except ValueError as exc:
        return _fail_exit("INPUT_INVALID", f"输入数据错误：{exc}", EXIT_INPUT)
    except InputDataError as exc:
        return _fail_exit("INPUT_INVALID", exc.to_payload()["message"], EXIT_INPUT)
    except (FileAccessError, FileContentError) as exc:
        return _fail_exit(exc.code, exc.to_payload()["message"], EXIT_IO)
    except Exception as exc:  # noqa: BLE001
        return _fail_exit("UNEXPECTED", f"未预期异常：{exc}", EXIT_UNEXPECTED)



# ----- apply (原 applyCharts) -----

import argparse
import json
import os
import sys
import time

from typing import Optional

try:
    import win32com.client as win32
except ImportError:  # Linux / 无 Word 时仍可 render PNG；apply 再要求 win32
    win32 = None

try:  # PIL 用于读取 PNG 原始宽高（等比缩放依据）
    from PIL import Image
    _HAS_PIL = True
except ImportError:  # pragma: no cover
    Image = None
    _HAS_PIL = False

CHART_W = 435  # pt，对应模板图框宽 441pt 略收
MAX_H_PT = 240  # pt，横向条紧凑封顶；正文 29pt 不改

P_TYPE_ANCHOR = "各作业类型日计划数量如下图所示"

# 单实例 WPS：自动化挂到与用户会话共用的同一后端上，迭代多次后偶发「文档打开/
# 保存失败」(3010/3011)。给出短暂恢复间隙重试，属外部 GUI 自动化的常规稳健兜底。
MAX_TRIES = 3
RETRY_GAP = 1.5  # 秒

# 日/周 报告单张承载力图前缀（月报按图题段匹配 cap_m / cap_w1..w5）
SINGLE_PREFIX = {"day": "cap_d1", "week": "cap_w1"}

# 月报 13 图位：周图题关键词 → 图形文件前缀（作业与管理共用 w1..w5）
_WEEK_PREFIX = (("第1周", "w1"), ("第2周", "w2"), ("第3周", "w3"),
                ("第4周", "w4"), ("第5周", "w5"))


def para_texts(doc):
    return [doc.Paragraphs(i).Range.Text.rstrip("\r\n\x07") for i in range(1, doc.Paragraphs.Count + 1)]


def caption_after(doc, texts, pindex):
    """月报：图段后紧跟图题段，返回该段文本（周报无图题段则 None）。"""
    for j in range(pindex + 1, min(pindex + 4, len(texts) + 1)):
        t = texts[j - 1]
        if "图 " in t or t.strip().startswith("图"):
            return t
    return None


def choose_cap_png(caption, index, report, cap_dir):
    """按图题/序号选择本次应插入的 PNG 路径（月报 13 图位按键匹配）。

    月报图题形如「图 3公司X月份第1周作业承载力」：作业类型→cap_m 作业类型；
    管理→cap_m/cap_wN 管理承载力；其余按 全月/第N周 → cap_m/cap_wN 单位承载力。
    日/周无月图题，仅首张承载力图（cap_d1/cap_w1）。
    """
    if report != "month":
        if index == 1:
            prefix = SINGLE_PREFIX.get(report, "cap_w1")
            return os.path.join(cap_dir, f"{prefix}_单位承载力.png")
        return None
    if not caption:
        return None
    if "作业类型" in caption:
        return os.path.join(cap_dir, "cap_m_作业类型.png")
    week = next((w for k, w in _WEEK_PREFIX if k in caption), None)
    head = "管理承载力" if "管理承载力" in caption else "单位承载力"
    return os.path.join(cap_dir, f"cap_{week if week else 'm'}_{head}.png")


def _png_ratio(png: str) -> Optional[float]:
    """读取 PNG 原始高/宽比（h/w）。无 PIL 时返回 None（回退旧等宽逻辑）。"""
    if not _HAS_PIL:
        return None
    with Image.open(png) as im:
        w, h = im.size
        return (h / w) if w else None


def add_pic(doc, png, rng):
    """在空 range 处插入图片：按 PNG 原始比例精确缩放，段落居中。

    先以 PIL 读 PNG 原始宽高计算目标宽高（宽 435pt，高随比例；过高时以高
    为准反算宽），再**关闭锁定纵横比**后显式写入宽高，避免部分 Word/WPS
    后端在锁定比例下把宽高分别独立设置导致图形被拉伸变形。
    图段统一「单倍行距 + KeepTogether + KeepWithNext」：防模板固定 29pt 行距把高图
    顶部压扁裁切；与图题同页。不改 Font，不改正文 29pt。
    """
    ratio = _png_ratio(png)
    if ratio is None:
        # 无 PIL 兜底：沿用旧的等宽锁比逻辑
        pic = doc.InlineShapes.AddPicture(png, False, True, rng)
        pic.LockAspectRatio = -1
        pic.Width = CHART_W
        if pic.Height > MAX_H_PT:
            pic.Height = MAX_H_PT
    else:
        pic = doc.InlineShapes.AddPicture(png, False, True, rng)
        width = CHART_W
        height = width * ratio
        if height > MAX_H_PT:
            height = MAX_H_PT
            width = height / ratio  # 以高为准反算宽，保持等比
        pic.LockAspectRatio = 0  # 关锁，显式宽高精确生效
        pic.Width = width
        pic.Height = height
    p = pic.Range.Paragraphs(1)
    _format_figure_para(p)
    return pic


# 图段分页统一交给 main() 保存后的「重开-按页码扫描」修复循环决定（编辑期
# Information(6) 几何在 WPS 下不可靠，不在此处做空间检测）。


def _fallback_anchor(doc):
    """兜底锚点：找含「承载力…如图所示」的段落索引（日/周单图报告无图框时）。"""
    for idx in range(1, doc.Paragraphs.Count + 1):
        t = doc.Paragraphs(idx).Range.Text.rstrip("\r\n\x07").strip()
        if "承载力" in t and "如图所示" in t:
            return idx
    return None


def replace_capacity_shapes(app, doc, report, cap_dir):
    """把渲染好的 PNG 填进模板图位，返回替换数。

    - 日/周：遍历文档中占位 InlineShape，替换为单张承载力图 PNG（cap_d1/cap_w1）；
      找不到图框时兜底插到「承载力…如图所示」段首。
    - 月报（13 图位 PNG 回退）：遍历「图 N公司…」图题段，其前一段为图位段，按图题
      关键词匹配对应 PNG 插到图位段首——图位段有无占位图均可（空段直接插）。
    """
    texts = para_texts(doc)
    specs = []
    if report == "month":
        for idx in range(1, doc.Paragraphs.Count + 1):
            t = texts[idx - 1]
            if t.startswith("图 ") and len(t) <= 40:
                png = choose_cap_png(t, idx, report, cap_dir)
                if png and idx > 1:
                    specs.append((png, idx - 1, t))
    else:
        count = doc.InlineShapes.Count
        for i in range(1, count + 1):
            rng0 = doc.InlineShapes(i).Range
            pi = None
            for idx in range(1, doc.Paragraphs.Count + 1):
                if abs(doc.Paragraphs(idx).Range.Start - rng0.Start) < 2:
                    pi = idx
                    break
            cap_txt = caption_after(doc, texts, pi)
            png = choose_cap_png(cap_txt, i, report, cap_dir)
            if png:
                specs.append((png, pi, cap_txt))
        if not specs:
            pi = _fallback_anchor(doc)
            if pi is not None:
                prefix = SINGLE_PREFIX.get(report, "cap_w1")
                specs = [(os.path.join(cap_dir, f"{prefix}_单位承载力.png"), pi, None)]

    print("待替换承载力图:", len(specs))
    for png, pi, cap_txt in specs:
        if pi is None:
            continue
        para = doc.Paragraphs(pi)
        prng = para.Range.Duplicate
        # 删除原占位图对象（保留段落）
        for k in range(1, doc.InlineShapes.Count + 1):
            if abs(doc.InlineShapes(k).Range.Start - prng.Start) < 2:
                doc.InlineShapes(k).Delete()
                break
        r = prng.Duplicate
        r.Collapse(1)  # 段首
        add_pic(doc, png, r)
        print(f"  替换 P{pi} 图题={cap_txt!r} -> {os.path.basename(png)}")
    return len(specs)


def insert_ptype(app, doc, s_report, cap_dir):
    """定位作业类型引导句段，其后新增一段插入作业类型分布图并居中。"""
    prefix = SINGLE_PREFIX.get(s_report, "cap_m")
    png = os.path.join(cap_dir, f"{prefix}_作业类型.png")
    if not os.path.exists(png):
        print("  跳过作业类型图（不存在）：", png)
        return False
    anchor = None
    for idx in range(1, doc.Paragraphs.Count + 1):
        t = doc.Paragraphs(idx).Range.Text.rstrip("\r\n\x07").strip()
        if t.startswith(P_TYPE_ANCHOR):
            anchor = doc.Paragraphs(idx)
            break
    if anchor is None:
        print("  未找到作业类型引导句段，跳过作业类型图")
        return False
    anchor.Range.InsertParagraphAfter()
    new_para = anchor.Next()
    nr = new_para.Range
    nr.Collapse(1)
    pic = add_pic(doc, png, nr)
    print(f"  作业类型图插入 anchor 段后 -> {os.path.basename(png)}")
    return True


# 管理承载力评述段锚点子串（create by genReportParams 的真实 F_管理 句）
MANAGE_ANCHOR = {"day": "本日管理承载力", "week": "本周管理承载力",
                 "month": "全月管理承载力"}

# 月报逐（全月+周）管理承载力：锚点关键词 → 图表前缀。正文各周评述段已由
# genReportParams 生成「全月/第N周管理承载力为X%」，含显式周号不歧义；
# 与 runReports 全周期 --mgmt 渲染的 {prefix}_管理承载力.png 一一对应。
_MONTH_MGMT_PLAN = [
    ("全月管理承载力", "cap_m"),
    ("第1周管理承载力", "cap_w1"),
    ("第2周管理承载力", "cap_w2"),
    ("第3周管理承载力", "cap_w3"),
    ("第4周管理承载力", "cap_w4"),
    ("第5周管理承载力", "cap_w5"),
]


def _insert_one_manage(doc, kw: str, png: str) -> bool:
    """在含 kw 的评述段后新增一段插入管理承载力横向条形图并居中。

    图文件不存在或评述段未找到时安全跳过（不影响报告其余部分）。
    评述段（含「管理承载力为」）由 _is_bind_para 判为绑定段，随图同页。
    """
    if not os.path.exists(png):
        print("  跳过管理承载力图（不存在）：", os.path.basename(png))
        return False
    anchor = None
    for idx in range(1, doc.Paragraphs.Count + 1):
        t = doc.Paragraphs(idx).Range.Text.rstrip("\r\n\x07").strip()
        if kw in t:
            anchor = doc.Paragraphs(idx)
            break
    if anchor is None:
        print(f"  未找到管理承载力评述段（含『{kw}』），跳过")
        return False
    anchor.Range.InsertParagraphAfter()
    nr = anchor.Next().Range
    nr.Collapse(1)
    # add_pic 已把图段统一为「单倍行距 + KeepTogether（整段不分页）」：
    # 前者防正文评述段的「固定行距（Rule=4 29pt）」把高图顶部标题/刻度压扁
    # 裁切，后者防超高图跨页中段被裁。图段分页由 main() 保存后的修复循环决定。
    add_pic(doc, png, nr)
    print(f"  管理承载力图插入『…{kw}…』段后 -> {os.path.basename(png)}")
    return True


def insert_manage_chart(app, doc, s_report, cap_dir):
    """在「…管理承载力为 X%」评述段后插入横向管理承载力条形图并居中。

    日/周各 1 张（{prefix}_管理承载力.png）；月报 6 张（全月 + 第1~5周各 1 张，
    「生成全」：凡有数据的周期都出图）。图由 renderReportCharts --mgmt 生成。
    未找到评述段或图文件不存在时安全跳过（不影响报告其余部分）。
    """
    if s_report == "month":
        plan = _MONTH_MGMT_PLAN
    else:
        kw = MANAGE_ANCHOR.get(s_report)
        prefix = SINGLE_PREFIX.get(s_report, "cap_m")
        plan = [(kw, prefix)]
    if not plan or not plan[0][0]:
        print(f"  不支持的报告形态 {s_report}，跳过管理承载力图")
        return False
    ok = False
    for kw, prefix in plan:
        if not kw:
            continue
        if _insert_one_manage(doc, kw,
                              os.path.join(cap_dir, f"{prefix}_管理承载力.png")):
            ok = True
    return ok


import re as _re  # 正则模块级一次性编译，供标题/引导句判定与组头回溯共用

_TITLE_RE = _re.compile(
    r"^(?:[一二三四五六七八九十]+、"
    r"|（[^）]*[0-9一二三四五六七八九十][^）]*）"
    r"|[0-9]+[.、．])")
_GUIDE_RE = _re.compile(r"(?:如图|如下图|如下表|如表)所示")

def _is_bind_para(p):
    """标题/引导句/评述段判定：章节编号开头（一、/（一）/1./（1））、「…如图/如表…所示」
    引导句、或「…管理承载力为X%…」管理评述段。纯文本判据，不涉字号——页位判定
    统一走「重开-按页码扫描」修复循环，不再依赖编辑期几何。

    管理评述段（含「管理承载力为」）须与其后紧随的管理承载力趋势图同页：否则评述
    落在页尾、趋势图被挤到次页，二者分离导致「图不见了」的观感。纳入绑定后由
    repair_layout 按可靠页码补段前分页，使评述与图整体迁新页。"""
    t = _para_plain(p)
    return bool(t) and (_TITLE_RE.match(t) or _GUIDE_RE.search(t)
                        or "管理承载力为" in t)


def _next_nonempty(doc, idx):
    """自 idx 之后找第一个「有效内容」段索引：非空文本，或含 InlineShape 的空图段
    （图段文本常为空，但必须以它为紧随内容才能锚定图题/引导句与图的同页关系）。"""
    n = doc.Paragraphs.Count
    j = idx + 1
    while j <= n:
        p = doc.Paragraphs(j)
        if p.Range.Text.rstrip("\r\n\x07").strip():
            return j
        try:
            if p.Range.InlineShapes.Count > 0:
                return j
        except Exception:
            pass
        j += 1
    return None


def _para_plain(p) -> str:
    try:
        return p.Range.Text.rstrip("\r\n\x07").strip()
    except Exception:
        return ""


def _format_figure_para(p) -> None:
    """仅图段：单倍行距、去首行缩进、居中、与下段同页。不改 Font。"""
    try:
        p.Alignment = 1
    except Exception:
        pass
    try:
        p.LineSpacingRule = 0
    except Exception:
        pass
    try:
        p.KeepTogether = True
        p.KeepWithNext = True
        p.Format.PageBreakBefore = False
    except Exception:
        pass
    try:
        p.FirstLineIndent = 0
    except Exception:
        pass
    try:
        p.Format.CharacterUnitFirstLineIndent = 0
    except Exception:
        pass
    try:
        p.CharacterUnitFirstLineIndent = 0
    except Exception:
        pass


def _format_caption_para(p) -> None:
    """图题：禁止单独段前分页。不改行距、不改 Font。"""
    p.KeepTogether = True
    p.KeepWithNext = False
    p.Format.PageBreakBefore = False


def _should_keep_with_next(doc, idx: int) -> bool:
    """标题/引导句只在后面紧跟图、图题、表，或短标题带下一段时绑定。"""
    p = doc.Paragraphs(idx)
    t = _para_plain(p)
    if not t:
        return False
    ni = _next_nonempty(doc, idx)
    if ni is None:
        return False
    nxt = doc.Paragraphs(ni)
    nt = _para_plain(nxt)
    if _has_img(nxt) or CAP_RE.match(nt):
        return True
    if _GUIDE_RE.search(t) or "管理承载力为" in t:
        return _has_img(nxt) or CAP_RE.match(nt)
    if _TITLE_RE.match(t) and _GUIDE_RE.search(nt):
        n2 = _next_nonempty(doc, ni)
        return bool(n2) and _has_img(doc.Paragraphs(n2))
    if _TITLE_RE.match(t) and len(t) <= _BIND_MAX_LEN:
        return True
    return False


def _figure_group_start(doc, fig_idx: int) -> int:
    """图组起点：短标题/引导句 + 中间空段 + 图段。"""
    start = fig_idx
    taken = 0
    j = fig_idx - 1
    while j >= 1 and taken < 2:
        q = doc.Paragraphs(j)
        t = _para_plain(q)
        if not t and not _has_img(q):
            start = j
            j -= 1
            continue
        if _is_bind_para(q) and (
                _GUIDE_RE.search(t) or "管理承载力为" in t
                or (_TITLE_RE.match(t) and len(t) <= _BIND_MAX_LEN)):
            start = j
            taken += 1
            j -= 1
            continue
        break
    return start


def apply_dynamic_layout(app, doc):
    """保存前整稿排版：图组绑定 + 图段单倍行距。

    正文 29pt 固定行距与模板字体一律不改。
    仅含图段落改为单倍行距并去掉首行缩进；引导句/短标题 KeepWithNext；
    图题禁止单独段前分页。
    """
    anchored, spelt = 0, 0
    n = doc.Paragraphs.Count
    for idx in range(1, n + 1):
        try:
            p = doc.Paragraphs(idx)
            if _should_keep_with_next(doc, idx):
                p.Format.KeepWithNext = True
                anchored += 1
            t = _para_plain(p)
            if CAP_RE.match(t):
                _format_caption_para(p)
            if not t and idx < n:
                nxt = doc.Paragraphs(idx + 1)
                if _has_img(nxt) or CAP_RE.match(_para_plain(nxt)):
                    p.Format.KeepWithNext = True
                    p.Format.PageBreakBefore = False
        except Exception as exc:
            print(f"  跳过段 {idx}: {exc}")
    try:
        nshp = int(doc.InlineShapes.Count)
    except Exception:
        nshp = 0
    for i in range(1, nshp + 1):
        try:
            shp = doc.InlineShapes(i)
            h = shp.Height
            if h <= 80:
                continue
            p = shp.Range.Paragraphs(1)
            _format_figure_para(p)
            spelt += 1
        except Exception as exc:
            print(f"  跳过图 {i}: {exc}")
    print(f"  动态排版: 标题/引导绑定 {anchored} 段，"
          f"图段单倍行距 {spelt} 段（正文 29pt/字体未改）")
    return anchored, spelt


def _close_prev_by_name(app, doc_path: str) -> None:
    """关闭共享实例里与目标稿同名的残留文档（只按自家中间稿文件名精确匹配，
    绝不触碰用户会话里的其它文档），避免上一流程未完全释放导致 Open 被 3010 拒绝。"""
    target = os.path.basename(doc_path).lower()
    for i in range(1, app.Documents.Count + 1):
        try:
            nm = app.Documents(i).Name
        except Exception:
            continue
        if nm and nm.lower() == target:
            app.Documents(i).Close(SaveChanges=False)
            return


CAP_RE = _re.compile(r"^图\s*\d")  # 月报图题段「图 1公司9月份承载力」
_BIND_MAX_LEN = 24  # 紧随文本的绑定段超此长度视为正文条目，不做孤行分页（防过度翻页）
MAX_CYCLES = 8  # 「重开-扫描-修复」收敛轮数上限

# ---- 宽表自动收窄（v2.6）-----------------------------------------------------
# 表格列宽总和超版心（如风险明细表 763pt、分块前矩阵表 933pt）时，按「语义
# 下限 + 余量按比例分摊」收窄到版心宽，保证所有列进版心、数据完整可见。
# 幂等：总宽已 ≤ 版心的表不再处理（repair_layout 收敛循环重开时为 no-op）。
_FLOOR_NO = 40      # 「序号」列下限（与 fillReport 序号列固定 40pt 一致）
_FLOOR_RISK = 36    # 「风险等级」列下限
_FLOOR_DATE = 56    # 日期列下限（矩阵表首列「日期」/ 模板日期列头 X月X日）
_FLOOR_MIN = 30     # 其余列下限


def _cell_head(tbl, c) -> str:
    """读第 c 列第 1 行表头文本（剥离单元格结束符/段落标记/并入前缀文本）。"""
    try:
        txt = tbl.Cell(1, c).Range.Text
    except Exception:
        return ""
    x07 = txt.rfind("\x07")
    if x07 >= 0:
        txt = txt[:x07]
    while txt.endswith("\r"):
        txt = txt[:-1]
    if "\r" in txt:
        txt = txt.split("\r")[-1]
    return txt.strip()


def fit_tables_to_text(app, doc) -> int:
    """列宽总和超版心的表格自动收窄到版心宽，返回收窄的表数。

    逐列下限按表头语义：序号 40 / 风险等级 36 / 日期 56 / 其余 30（pt）；
    当前宽已比下限窄的列不膨胀（下限取当前宽）；全部列按下限仍超版心时
    整表等比缩放。收窄后 AllowAutoFit=False 锁宽，防保存时被重置。
    """
    ps = doc.PageSetup
    text_w = ps.PageWidth - ps.LeftMargin - ps.RightMargin
    fitted = 0
    for ti in range(1, doc.Tables.Count + 1):
        tbl = doc.Tables(ti)
        try:
            ncols = tbl.Columns.Count
            widths = [tbl.Columns(c).Width for c in range(1, ncols + 1)]
        except Exception:
            continue
        total = sum(widths)
        if total <= text_w + 0.5:
            continue
        floors = []
        for c in range(1, ncols + 1):
            h = _cell_head(tbl, c)
            if "序号" in h:
                f = _FLOOR_NO
            elif "风险等级" in h:
                f = _FLOOR_RISK
            elif "日期" in h or _re.match(r"^\d{1,2}月\d{1,2}日?$", h):
                f = _FLOOR_DATE
            else:
                f = _FLOOR_MIN
            floors.append(min(widths[c - 1], f))
        if sum(floors) >= text_w:
            scale = text_w / total
            new_w = [w * scale for w in widths]
        else:
            surplus = text_w - sum(floors)
            room = [w - f for w, f in zip(widths, floors)]
            tot_room = sum(room)
            if tot_room <= 0:
                new_w = [f + surplus / ncols for f in floors]
            else:
                new_w = [f + r * surplus / tot_room for f, r in zip(floors, room)]
        for c in range(1, ncols + 1):
            try:
                tbl.Columns(c).Width = new_w[c - 1]
            except Exception:
                pass
        try:
            tbl.AllowAutoFit = False
        except Exception:
            pass
        fitted += 1
        print(f"  宽表收窄 T{ti}: {total:.0f}pt -> {text_w:.0f}pt（{ncols} 列）")
    return fitted


def _para_page(doc, p):
    """段起点的页码（Information(3)=wdActiveEndPageNumber，折叠到段首）。
    仅对「新鲜打开」的已存稿件可靠（重开后与最终渲染一致）；编辑期数值与渲染
    不一致，故修复循环一律先保存再重开测量。"""
    try:
        r = p.Range.Duplicate
        r.Collapse(1)  # wdCollapseStart=1
        v = r.Information(3)
        return int(v) if isinstance(v, (int, float)) else None
    except Exception:
        return None


def _has_img(p):
    try:
        return p.Range.InlineShapes.Count > 0
    except Exception:
        return False


def _prev_figure_para(doc, idx):
    """自 idx 向上找最近的含图段（跳过空段；遇首个非空无图段即停）。"""
    for j in range(idx - 1, 0, -1):
        q = doc.Paragraphs(j)
        if q.Range.Text.rstrip("\r\n\x07").strip():
            return j if _has_img(q) else None
    return None


def repair_layout(app, doc_path, report):
    """保存后的「重开-按页码扫描-修复」收敛循环：WPS 编辑期 Information(3)/(6)
    均不可靠（i(6) 常错误折叠回页顶、i(3) 编辑期会级联误分页），而新鲜打开已存
    稿件时 i(3) 页码与最终渲染一致。每轮：重开 → 用可靠页码找「孤行标题/图题分离」
    → 统一加 PageBreakBefore → 重存；分页让页位变化，下轮重扫直至无命中。

    命中规则：
    - 标题/引导句/评述段页码 < 紧随内容页码，且（紧随为图段，
      或自身较短 ≤ _BIND_MAX_LEN）→ 给组头段前分页，整组迁新页。
    - 图题与上方图段异页 → 给图组起点段前分页，禁止图题单独 PageBreakBefore。
    不改正文 29pt、不改 Font。
    返回收敛轮数。
    """
    target = os.path.abspath(doc_path)
    used = 0
    for cycle in range(1, MAX_CYCLES + 1):
        used = cycle
        _close_prev_by_name(app, target)
        doc = app.Documents.Open(target, ReadOnly=False, AddToRecentFiles=False)
        n = doc.Paragraphs.Count
        to_break = {}
        for idx in range(1, n + 1):
            try:
                p = doc.Paragraphs(idx)
                t = _para_plain(p)
                if not t:
                    continue
                pp = _para_page(doc, p)
                if pp is None:
                    continue
                ni = _next_nonempty(doc, idx)
                if ni is None:
                    continue
                npp = _para_page(doc, doc.Paragraphs(ni))
                if npp is None or npp <= pp:
                    continue
                if CAP_RE.match(t):
                    f_ = _prev_figure_para(doc, idx)
                    if f_ is not None:
                        fp = _para_page(doc, doc.Paragraphs(f_))
                        if fp is not None and fp < pp:
                            g = _figure_group_start(doc, f_)
                            to_break[g] = f"图题P{idx}分离（整组随迁）"
                elif _is_bind_para(p):
                    if _has_img(doc.Paragraphs(ni)) or len(t) <= _BIND_MAX_LEN:
                        to_break[idx] = f"孤行标题（后随P{ni}在次页）"
            except Exception as exc:
                print(f"    跳过段 {idx}: {exc}")
        if not to_break:
            doc.Close(SaveChanges=False)
            break
        for idx in sorted(to_break):
            doc.Paragraphs(idx).Format.PageBreakBefore = True
            print(f"    [{report}] P{idx} 段前分页（{to_break[idx]}）: "
                  f"{doc.Paragraphs(idx).Range.Text.split(chr(13))[0][:24]}")
        doc.Save()
        doc.Close(SaveChanges=False)
    return used


def polish_native_layout(doc_path: str, report: str) -> int:
    """docx 注入可编辑图后：图段单倍行距 + 图组防拆页。不改正文 29pt/字体。

    使用独立 Word 实例，不复用流水线共用 COM：重开收敛若把实例打挂，
    不得连累 fillReport 的共享 Word。
    """
    target = os.path.abspath(doc_path)
    app = win32.DispatchEx("Word.Application")
    app.Visible = False
    try:
        app.DisplayAlerts = False
    except Exception:
        pass
    try:
        _close_prev_by_name(app, target)
        doc = app.Documents.Open(target, ReadOnly=False, AddToRecentFiles=False)
        try:
            apply_dynamic_layout(app, doc)
        except Exception as exc:
            print(f"  动态排版未完成：{exc}")
        try:
            fit_tables_to_text(app, doc)
        except Exception as exc:
            print(f"  宽表收窄未完成：{exc}")
        doc.Save()
        doc.Close(SaveChanges=False)
        try:
            return repair_layout(app, target, report)
        except Exception as exc:
            print(f"  分页收敛未完成：{exc}")
            return 0
    finally:
        try:
            app.Quit()
        except Exception:
            pass


def apply_main(argv=None):
    if win32 is None:
        print(json.dumps({"ok": False, "errorCode": "DEP_MISSING",
                          "message": "未安装 pywin32 / 无法驱动 Word，无法 apply PNG"},
                         ensure_ascii=False))
        return 1
    ap = argparse.ArgumentParser(description="承载力报告动态图表插入")
    ap.add_argument("--doc", required=True)
    ap.add_argument("--report", choices=("day", "week", "month"), required=True)
    ap.add_argument("--charts", required=True)
    ap.add_argument("--out")
    args = ap.parse_args(argv)

    # 统一绝对路径：Word/WPS 的 COM Open/AddPicture 以进程自身 CWD 解析相对路径，
    # 与调用方 CWD 不一致时会 3010「文档打开失败」。入口即转绝对路径根治。
    args.doc = os.path.abspath(args.doc)
    args.charts = os.path.abspath(args.charts)
    if args.out:
        args.out = os.path.abspath(args.out)

    if not os.path.exists(args.doc):
        print(json.dumps({"ok": False, "errorCode": "IO", "message": "文档不存在"},
                         ensure_ascii=False))
        return 1

    def _rpc_dead(exc: BaseException) -> bool:
        text = str(exc)
        return ("-2147023174" in text) or ("RPC" in text) or ("远程过程调用" in text)

    last_err = None
    force_owned = False
    for attempt in range(1, MAX_TRIES + 1):
        owned = False
        app = None
        try:
            from fillReport import current_word_app, recycle_shared_word
            shared_app = None if force_owned else current_word_app()
        except Exception:
            shared_app = None
            recycle_shared_word = None  # type: ignore
        if shared_app is not None:
            app = shared_app
        else:
            app = win32.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = False
            owned = True
        try:
            _close_prev_by_name(app, args.doc)
            doc = app.Documents.Open(args.doc, ReadOnly=False, AddToRecentFiles=False)
            n1 = replace_capacity_shapes(app, doc, args.report, args.charts)
            if args.report != "month":
                insert_ptype(app, doc, args.report, args.charts)
                insert_manage_chart(app, doc, args.report, args.charts)
            apply_dynamic_layout(app, doc)
            fit_tables_to_text(app, doc)
            if args.out:
                doc.SaveAs(args.out)
            else:
                doc.Save()
            n_after = -1
            sizes = []
            try:
                n_after = int(doc.InlineShapes.Count)
                sizes = [
                    f"{doc.InlineShapes(i).Width:.0f}x{doc.InlineShapes(i).Height:.0f}"
                    for i in range(1, n_after + 1)
                ]
            except Exception:
                pass
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
            target = args.out if args.out else args.doc
            cycles = 0
            try:
                cycles = repair_layout(app, target, args.report)
            except Exception as exc:  # noqa: BLE001 收敛失败不推翻已保存成品
                print(f"  分页收敛跳过：{exc}")
            print(json.dumps({"ok": True, "report": args.report,
                              "replaced": n1, "inline_shapes_after": n_after,
                              "sizes": sizes, "layout_cycles": cycles},
                             ensure_ascii=False))
            return 0
        except Exception as exc:  # noqa: BLE001
            last_err = exc
            try:
                _close_prev_by_name(app, args.doc)
            except Exception:
                pass
            if _rpc_dead(exc):
                force_owned = True
                if recycle_shared_word is not None:
                    try:
                        recycle_shared_word()
                    except Exception:
                        pass
            if attempt < MAX_TRIES:
                time.sleep(RETRY_GAP * attempt)
        finally:
            if owned and app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
    print(json.dumps({"ok": False, "errorCode": "UNEXPECTED",
                      "message": str(last_err)}, ensure_ascii=False))
    return 1




def main(argv=None):
    """入口：pngCharts.py render|apply ... 或按参数推断。"""
    import sys
    args = list(sys.argv[1:] if argv is None else argv)
    if args and args[0] in ("render", "apply"):
        mode = args.pop(0)
        return render_main(args) if mode == "render" else apply_main(args)
    # 兼容：带 --doc 视为 apply
    if any(a == "--doc" or a.startswith("--doc=") for a in args):
        return apply_main(args)
    return render_main(args)


if __name__ == "__main__":
    raise SystemExit(main())
