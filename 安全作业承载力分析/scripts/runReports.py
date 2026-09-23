# -*- coding: utf-8 -*-
"""
报告一键编排生成器（技能「安全作业承载力分析」内置工具）
======================================================

把 承载力量化计算 → 报告正文参数 → 图表渲染 → 模板填充 → 图表插入 五步串成
一条流水线，一键产出 日 / 周 / 月 正式报告（成品 .doc，含动态承载力图与作业类型图）。

调用（示例）：
  python runReports.py \
      --source ../../测试数据/工作计划-测试数据.xls \
      --template 报告-月 \
      --report-date 2026-09-01 \
      --config ../config/capacity_config.json \
      --out-dir ../承载力报告

流水线（同进程调用各脚本 main()；中间产物统一放在 <out-dir>/_work/）：
  1) calcCapacity.py --period 日            → 结果 JSON（days / management + matrix）
  2) genReportParams.py --template 报告-X    → fillReport 参数 JSON
  3) pngCharts.py render                 → 承载力图 + 作业类型图 PNG（--chart png）
  4) fillReport.py --template 报告-X         → 正文草稿
  5) pngCharts.py apply / nativeCharts.py → 插图或原生图成品
  月报 native：calc 只算一次，fillReport 后走 nativeCharts（office_chart.native_main）。
  --all：三份报告共用同一次 calc；日→周→月之间 recycle_shared_word 清 COM 脏状态。

说明：
  - 承载力图依赖 calc 输出的 matrix，而 matrix 需配置「输入解析.组织架构映射.是否启用=true」。
    若原配置未启用，本脚本自动生成一份临时配置副本（仅打开该开关）供本次流水线使用，
    不修改原配置；如需按真实组织架构归列，请直接在该配置中打开开关。
  - 报告成品按模板命名规则落盘到 <out-dir>/；--keep 可保留 _work/ 中间产物便于核对。

约束（勿违反，否则自撞难排查）：
  - 同一 out-dir 不能并发跑两个 runReports：中间文件（calc_result/letters/params.json）
    不经并发保护，并行时互相覆写，产物张冠李戴。
  - 不同 out-dir 也不要并行：fillReport/applyCharts 挂在用户会话共用的 WPS 单实例后端，
    两路并行会争用导致 3011「文档保存失败」（3 次重试也未必救回）。多份报告/函件请串行。
  - 全局纪律：同一时刻仅一路 WPS 自动化 + 每 runReports 独占 out-dir。

退出码：0=成功  2=用法错误  3=输入/中间数据错误  4=文件/IO 错误  1=未知异常
"""
from __future__ import annotations

import argparse
import contextlib
import copy
import io
import json
import logging
import os
import shutil
import sys
from datetime import date, timedelta
from typing import Any, Dict, List, Optional, Tuple

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOGGER = logging.getLogger("runReports")

EXIT_OK, EXIT_UNEXPECTED, EXIT_USAGE, EXIT_INPUT, EXIT_IO = 0, 1, 2, 3, 4

# ---- 复用 genReportParams 的周/月区间划分与模板表 ---------------------------
# 保证：图表区间 = 正文周期 = 报告标题月份，三处口径完全一致，不重复实现。
from genReportParams import TEMPLATES as _TEMPLATES  # noqa: E402
from genReportParams import (_fmt_d, _iso_week, _month_bounds,  # noqa: E402
                             _month_weeks)

# 模板 → applyCharts 报告形态（day/week/month 三选一）
_REPORT_KIND = {"报告-日": "day", "报告-周": "week", "报告-月": "month"}

# 函件模板（无图表，仅 calc → genLetterParams → 逐封 fillReport）
_IS_LETTER = {"警示函-满载", "提示函-重载", "提示函-长期作业"}


def _month_ordinal(d: date) -> int:
    """报告日所在月内周序号（与正文标题『X月份第N周』一致）。"""
    return next((o for o, a, b in _month_weeks(d) if a <= d <= b), 1)


def _is_full_month(lo: date, hi: date) -> bool:
    """范围是否为整自然月（lo 为该月 1 日、hi 为该月末日、同月同年）。"""
    if lo.year != hi.year or lo.month != hi.month or lo.day != 1:
        return False
    nxt = (date(hi.year + 1, 1, 1) if hi.month == 12
           else date(hi.year, hi.month + 1, 1)) - timedelta(days=1)
    return hi == nxt


def _pick_range_template(lo: date, hi: date) -> str:
    """v2.8 自定义时间范围 → 报告类型自动判定。

    1 天→日报；2~7 天→周报；>7 天且整自然月→月报；>7 天不足整月→周报
    （月报仅支持整自然月，用户确认口径「不到一个月的不生成」）。
    """
    span = (hi - lo).days + 1
    if span == 1:
        return "报告-日"
    if span > 7 and _is_full_month(lo, hi):
        return "报告-月"
    return "报告-周"


def _schedules(template: str, d: date,
               lo: Optional[date] = None, hi: Optional[date] = None
               ) -> List[Dict[str, Any]]:
    """按模板返回图表生成计划：[{prefix, lo, hi, label, ptype}]。

    v2.8：lo/hi（用户自定义时间范围）给定时主周期图表按范围出图，
    周次/月次命名仍按锚定日 d 推导（保留原命名）。
    """
    m = d.month
    if template == "报告-日":
        return [{"prefix": "cap_d1", "lo": d, "hi": d,
                 "label": f"2026年{m}月{d.day}日各单位承载力", "ptype": True}]
    if template == "报告-周":
        w_lo, w_hi = (lo, hi) if lo is not None and hi is not None else _iso_week(d)
        return [{"prefix": "cap_w1", "lo": w_lo, "hi": w_hi,
                 "label": f"2026年{m}月份第{_month_ordinal(d)}周"
                          f"（{_fmt_d(w_lo)}至{_fmt_d(w_hi)}）各单位承载力",
                 "ptype": True}]
    # 报告-月：全月 1 张（含作业类型图）+ 每周 1 张（渲染时每周另带管理承力图）
    m_lo, m_hi = (lo, hi) if lo is not None and hi is not None else _month_bounds(d)
    specs = [{"prefix": "cap_m", "lo": m_lo, "hi": m_hi,
              "label": f"2026年{m_lo.month}月份各单位承载力", "ptype": True}]
    # v2.8：有范围时按范围所在月枚举月内周（整月范围与锚定日所在月一致）
    for ordinal, w_lo, w_hi in _month_weeks(m_hi)[:5]:
        specs.append({"prefix": f"cap_w{ordinal}", "lo": w_lo, "hi": w_hi,
                      "label": f"2026年{m}月份第{ordinal}周"
                               f"（{_fmt_d(w_lo)}至{_fmt_d(w_hi)}）各单位承载力",
                      "ptype": False})
    return specs


def _report_basename(template: str, d: date) -> str:
    """报告成品文件名（与模板『国网青岛供电公司2026年X月X日…』命名规则一致）。"""
    m, dd = d.month, d.day
    if template == "报告-日":
        return f"国网青岛供电公司2026年{m}月{dd}日安全承载力分析报告"
    if template == "报告-周":
        return f"国网青岛供电公司2026年{m}月份第{_month_ordinal(d)}周安全承载力分析报告"
    return f"国网青岛供电公司2026年{m}月份安全承载力分析报告"


def _ensure_org_cfg(cfg_path: str, work_dir: str) -> str:
    """返回本次流水线使用的配置路径。

    组织架构映射未启用时，生成一份临时副本（仅把该开关打开）供 calc / params /
    charts 共同使用 —— 承载力图依赖 calc 输出的 matrix。原配置保持不变。
    """
    with open(cfg_path, "r", encoding="utf-8") as fh:
        cfg = json.load(fh)
    inp = cfg.get("输入解析") or {}
    org = inp.get("组织架构映射") or {}
    if org.get("是否启用"):
        return cfg_path
    LOGGER.warning("组织架构映射未启用：承载力图需 calc 输出 matrix。"
                   "已为本次流水线临时生成启用副本，原配置未改动。")
    org["是否启用"] = True
    inp["组织架构映射"] = org
    cfg["输入解析"] = inp
    out = os.path.join(work_dir, "_config_org.json")
    tmp = out + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, out)
    return out


# 内置脚本 → 模块名（同目录，由 runReports 所在路径进入 sys.path）
_SCRIPT_MAINS = {
    "calcCapacity.py": "calcCapacity",
    "genReportParams.py": "genReportParams",
    "genLetterParams.py": "genLetterParams",
    "fillReport.py": "fillReport",
    "pngCharts.py": "pngCharts",
    "nativeCharts.py": "nativeCharts",
}
_COM_SCRIPTS = frozenset({"fillReport.py", "pngCharts.py", "nativeCharts.py"})


def _run(cmd: List[str], desc: str) -> None:
    """同进程调用内置脚本 main(argv)；非零退出直接抛错中断流水线。

    内置脚本一律相对 SCRIPT_DIR 定位（cmd 首个元素为脚本名），
    使流水线与调用方当前目录无关——从任何目录引用本技能都能跑通。
    stdout 捕获（避免 calc JSON 刷屏）；日志仍走各脚本 stderr。
    """
    if cmd and cmd[0].endswith(".py") and not os.path.isabs(cmd[0]):
        cmd = [os.path.join(SCRIPT_DIR, cmd[0])] + list(cmd[1:])
    LOGGER.info("- %s", desc)
    script = os.path.basename(cmd[0])
    mod_name = _SCRIPT_MAINS.get(script)
    if mod_name is None:
        raise RuntimeError(f"{desc} 失败：未知内置脚本 {script}")
    # pngCharts render 不需要 Word；仅 fillReport / nativeCharts / pngCharts apply 起 COM
    needs_word = script in ("fillReport.py", "nativeCharts.py") or (
        script == "pngCharts.py" and len(cmd) > 1 and cmd[1] == "apply")
    if needs_word:
        import fillReport as _fr
        if _fr.word_com_available() and _fr.current_word_app() is None:
            _fr.begin_shared_word()
    mod = __import__(mod_name)
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            code = mod.main(list(cmd[1:]))
    except SystemExit as exc:
        code = exc.code
        if code is None:
            code = 0
        elif not isinstance(code, int):
            code = 1
    except Exception as exc:  # noqa: BLE001 把脚本异常收成流水线错误
        tail = buf.getvalue().strip()[-400:]
        raise RuntimeError(f"{desc} 失败：{exc}\nstdout: {tail}") from exc
    if not code:
        return
    tail = buf.getvalue().strip()[-400:]
    raise RuntimeError(f"{desc} 失败（退出码 {code}）\nstdout: {tail}")


def _letter_pipeline(args: argparse.Namespace, cfg_path: str,
                     work_dir: str, report_date: date) -> List[str]:
    """函件编排：承载力量化 → 函件参数 → 逐封 fillReport（无图表）。

    返回成品 .doc 路径列表（周期内无命中单位/负责人时为空列表，非错误）。
    v2.8：带 --start/--end 时函件按自定义时间范围判定（--lo/--hi 透传）。
    """
    source = os.path.abspath(args.source_path)
    calc_json = os.path.join(work_dir, "calc_result.json")
    letters_json = os.path.join(work_dir, "letters.json")
    common = [f"--config", cfg_path]
    _run(["calcCapacity.py", "--xls", source, "--period", "日",
          "--out", calc_json, *common], "承载力量化计算（函件）")
    lcmd = ["genLetterParams.py", "--result", calc_json, "--source", source,
            "--template", args.template, "--report-date", report_date.isoformat(),
            "--config", cfg_path, "--out", letters_json]
    if getattr(args, "office", None):
        lcmd += ["--office", args.office]
    if args.start and args.end:
        lcmd += ["--lo", args.start, "--hi", args.end]
    _run(lcmd, "生成函件参数（genLetterParams）")
    with open(letters_json, "r", encoding="utf-8") as fh:
        payload = json.load(fh)
    letters = payload.get("letters") or []
    outputs: List[str] = []
    for i, letter in enumerate(letters, start=1):
        params_file = os.path.join(work_dir, f"letter_{i}.params.json")
        with open(params_file, "w", encoding="utf-8") as fh:
            json.dump({"replace": letter.get("replace") or {},
                       "paragraph_replace": letter.get("paragraph_replace") or [],
                       "tables": letter.get("tables") or []},
                      fh, ensure_ascii=False)
        final = os.path.join(args.out_dir, letter["basename"] + ".doc")
        _run(["fillReport.py", "--template", letter.get("template") or args.template,
              "--params", params_file, "--out", final],
             f"填充函件（{i}/{len(letters)}：{letter['basename']}）")
        outputs.append(os.path.abspath(final))
    return outputs


def _pipeline(args: argparse.Namespace, cfg_path: str, work_dir: str,
              report_date: date, specs: List[Dict[str, Any]],
              chart_mode: str = "png", skip_calc: bool = False) -> str:
    """顺序执行，返回最终报告路径。

    v2.19：日/周/月默认 chart_mode=native → 走 _native_pipeline（Word 原生可编辑
    竖柱，横轴类型、纵轴百分比/项数）；`--chart png` 回退走原 PNG 五步。
    skip_calc=True 时复用 work_dir 已有 calc_result.json（--all / 上游已算过）。
    """
    source = os.path.abspath(args.source_path)
    calc_json = os.path.join(work_dir, "calc_result.json")
    params_json = os.path.join(work_dir, "params.json")

    common = [f"--config", cfg_path]
    # 1) 承载力量化计算（日粒度：逐日出 days + matrix + management）
    if skip_calc:
        if not os.path.exists(calc_json):
            raise RuntimeError(f"跳过计算但缺少结果文件：{calc_json}")
    else:
        _run(["calcCapacity.py", "--xls", source, "--period", "日",
              "--out", calc_json, *common], "承载力量化计算")
    # 2) 报告正文参数
    cmd = ["genReportParams.py", "--result", calc_json, "--source", source,
           "--template", args.template, "--report-date", report_date.isoformat(),
           "--config", cfg_path, "--out", params_json]
    if args.advice_path:
        cmd += ["--advice", os.path.abspath(args.advice_path)]
    # v2.8：自定义时间范围透传——正文各段按范围判定，范围外数据不进汇总
    if args.start and args.end:
        cmd += ["--lo", args.start, "--hi", args.end]
    _run(cmd, "生成报告填充参数")

    # v2.19：日/周/月 + native → Word 原生可编辑竖柱（calc/params 已在上方完成）
    if args.template in ("报告-日", "报告-周", "报告-月") and chart_mode == "native":
        return _native_pipeline(args, cfg_path, work_dir, report_date, specs)

    # 3) 图表渲染（PNG 位图：每个周期一张承载力图+一张管理承力图；仅主周期附带
    #    作业类型图）。月报 13 图位所需 PNG 全集：全月 cap_m（单位/管理/作业类型）
    #    + 每周 cap_w1..w5（单位/管理），文件名带唯一 prefix，不冲突。
    charts_dir = os.path.join(work_dir, "charts")
    os.makedirs(charts_dir, exist_ok=True)
    for i, sp in enumerate(specs, start=1):
        cmd = ["pngCharts.py", "render", "--data", calc_json, "--xls", source,
               "--config", cfg_path,
               "--lo", sp["lo"].isoformat(), "--hi", sp["hi"].isoformat(),
               "--label", sp["label"], "--prefix", sp["prefix"],
               "--out-dir", charts_dir]
        if not sp["ptype"]:
            cmd.append("--no-ptype")
        cmd.append("--mgmt")
        _run(cmd, f"渲染图表（{i}/{len(specs)}：{sp['prefix']}）")
    # 4) 模板填充 → 正文草稿
    draft = os.path.join(work_dir, args.basename + ".doc")
    _run(["fillReport.py", "--template", args.template,
          "--params", params_json, "--out", draft],
         "填充报告模板")
    # 5) 插入图表 → 最终成品（先回收共用 Word，避免上游 COM 脏状态拖垮 apply）
    try:
        import fillReport as _fr
        if _fr.current_word_app() is not None:
            _fr.recycle_shared_word()
    except Exception:
        pass
    final = os.path.join(args.out_dir, args.basename + ".doc")
    _run(["pngCharts.py", "apply", "--doc", draft, "--report", _REPORT_KIND[args.template],
          "--charts", charts_dir, "--out", final],
         "插入承载力图与作业类型图")
    return final


def _native_pipeline(args: argparse.Namespace, cfg_path: str, work_dir: str,
                     report_date: date, specs: List[Dict[str, Any]]) -> str:
    """日/周/月原生可编辑图表编排：填充 → DrawingML 竖柱 → 排版收敛 → 成品 .docx。

    计算与正文参数须已由 _pipeline 写入 work_dir（calc_result.json / params.json）。
    """
    calc_json = os.path.join(work_dir, "calc_result.json")
    params_json = os.path.join(work_dir, "params.json")
    if not os.path.exists(calc_json) or not os.path.exists(params_json):
        raise RuntimeError("原生图表编排缺少 calc_result.json 或 params.json")
    draft = os.path.join(work_dir, args.basename + ".docx")
    _run(["fillReport.py", "--template", args.template,
          "--params", params_json, "--out", draft],
         "填充报告模板")
    week_specs = json.dumps(
        [{"prefix": sp["prefix"], "lo": sp["lo"].isoformat(),
          "hi": sp["hi"].isoformat()} for sp in specs[1:]], ensure_ascii=False)
    kind = _REPORT_KIND[args.template]
    final = os.path.join(args.out_dir, args.basename + ".docx")
    _run(["nativeCharts.py", "--doc", draft, "--result", calc_json,
          "--source", os.path.abspath(args.source_path), "--config", cfg_path,
          "--lo", specs[0]["lo"].isoformat(), "--hi", specs[0]["hi"].isoformat(),
          "--kind", kind, "--specs", week_specs, "--out", final],
         "写入 Word 可编辑图表")
    return final


def build_main_parser() -> argparse.ArgumentParser:
    from genReportParams import DEFAULT_OFFICE
    p = argparse.ArgumentParser(
        prog="runReports.py",
        description="承载力报告一键编排：计算→正文→图表→成稿（技能「安全作业承载力分析」）")
    p.add_argument("--source", dest="source_path", required=True,
                   metavar="PATH", help="作业计划源 .xls/.xlsm/.xlsx 或 .json")
    p.add_argument("--template", dest="template", default=None,
                   choices=list(_TEMPLATES.keys()),
                   help="报告/函件模板（指定单份）；与 --all / --templates 互斥")
    p.add_argument("--templates", dest="templates", default=None,
                   metavar="LIST",
                   help="多份报告逗号分隔，如 报告-日,报告-周（两份同批串行）；"
                        "与 --all / --template 互斥")
    p.add_argument("--all", dest="all", action="store_true",
                   help="一句话生成全部主报告（报告-日 + 报告-周 + 报告-月，同锚定日期，"
                        "WPS 单实例串行）；与 --template / --templates 互斥")
    p.add_argument("--report-date", dest="report_date", default=None,
                   metavar="YYYY-MM-DD",
                   help="报告锚定日期（日报=当天，周报=所在周，月报=所在月）；"
                        "指定 --start/--end 时间范围时默认取范围结束日")
    p.add_argument("--start", dest="start", default=None, metavar="YYYY-MM-DD",
                   help="自定义时间范围起始日（须与 --end 成对）：报告/函件按该范围判定，"
                        "范围外的数据不进报告汇总")
    p.add_argument("--end", dest="end", default=None, metavar="YYYY-MM-DD",
                   help="自定义时间范围结束日（须与 --start 成对）")
    p.add_argument("--config", dest="config_path", default=None,
                   metavar="PATH", help="capacity_config.json 路径（缺省自动找技能目录 config/）")
    p.add_argument("--advice", dest="advice_path", default=None,
                   metavar="PATH",
                   help="AI 辅助建议 JSON（出任一报告必填）：须含本次每份报告模板的非空正文，"
                        "注入文末「四、辅助建议与风险管控提示」章；禁止缺省回退 DEFAULT_ADVICE")
    p.add_argument("--office", dest="office", default=None,
                   metavar="NAME",
                   help="所属机构确认（出报告或函件均必填）：用户确认后传入；"
                        f"默认名「{DEFAULT_OFFICE}」确认后可原样传入；"
                        "县公司安监部改为对应安委办全名（函件落款替换用）")
    p.add_argument("--out-dir", dest="out_dir", required=True,
                   metavar="PATH",
                   help="产物根目录（工作区建议为「承载力报告」）：其下自动建"
                        "「承载力报告{YYYYMMDD}」日期子文件夹收纳全部成品，不散放文件")
    p.add_argument("--chart", dest="chart_mode", default="native",
                   choices=("native", "png"),
                   help="图表渲染方式：native=Word 可编辑竖柱（默认，日/周/月，双击改数据）；"
                        "png=位图回退")
    p.add_argument("--keep", action="store_true",
                   help="保留 _work/ 中间产物（calc.json / params.json / charts / 草稿.doc）")
    return p


def _parse_templates_arg(raw: str) -> List[str]:
    """解析 --templates 逗号分隔列表，去空白、保序去重。"""
    seen = set()
    out: List[str] = []
    for part in raw.split(","):
        name = part.strip()
        if not name or name in seen:
            continue
        seen.add(name)
        out.append(name)
    return out


def _run_report_batch(
        args: argparse.Namespace,
        templates: List[str],
        cfg_path: str,
        work_dir: str,
        report_date: date,
        range_lo: Optional[date],
        range_hi: Optional[date],
        precomputed_calc: bool) -> Tuple[List[str], int]:
    """串行生成多份主报告；precomputed_calc=True 时全部 skip_calc，否则首份算、其后复用。"""
    files: List[str] = []
    total_charts = 0
    for idx, tpl in enumerate(templates):
        args.template = tpl
        args.basename = _report_basename(tpl, report_date)
        specs = _schedules(tpl, report_date, range_lo, range_hi)
        LOGGER.info("报告：%s（锚定 %s，图表 %d 张）", args.basename,
                    report_date.isoformat(), len(specs))
        skip = precomputed_calc or idx > 0
        files.append(os.path.abspath(
            _pipeline(args, cfg_path, work_dir, report_date, specs,
                      args.chart_mode, skip_calc=skip)))
        total_charts += len(specs)
        try:
            import fillReport as _fr
            if _fr.current_word_app() is not None:
                _fr.recycle_shared_word()
        except Exception:
            pass
    return files, total_charts


def _setup_logging() -> None:
    logging.basicConfig(level=logging.INFO, stream=sys.stderr,
                        format="[%(levelname)s] %(message)s")


def _begin_template_prefetch(templates: List[str]):
    """无 Word 时并行把日/周/月 .doc 转成 docx，和后面的计算叠在一起跑。

    返回 (pool, futures)。调用方必须 _finish_template_prefetch，避免线程残留。
    """
    from concurrent.futures import ThreadPoolExecutor

    try:
        import fillReport as fr
        if fr.word_com_available():
            return None, []
        from fillReportLinux import doc_to_docx_cached
    except Exception as exc:  # noqa: BLE001
        LOGGER.warning("模板预转换未启动：%s", exc)
        return None, []
    paths: List[str] = []
    for tpl in templates:
        name = fr.TEMPLATES.get(tpl)
        if not name:
            continue
        path = os.path.join(fr.TEMPLATES_DIR, name)
        if os.path.isfile(path) and not path.lower().endswith(".docx"):
            paths.append(path)
    if not paths:
        return None, []
    pool = ThreadPoolExecutor(max_workers=min(3, len(paths)))
    futures = [pool.submit(doc_to_docx_cached, path) for path in paths]
    return pool, futures


def _finish_template_prefetch(pool, futures) -> None:
    if pool is None and not futures:
        return
    for fut in futures or []:
        try:
            fut.result()
        except Exception as exc:  # noqa: BLE001
            LOGGER.warning("模板预转换未完成，填充时会再试：%s", exc)
    if pool is not None:
        pool.shutdown(wait=True)


def main(argv: Optional[List[str]] = None) -> int:
    args = build_main_parser().parse_args(argv)
    _setup_logging()
    def _fail(key: str, msg: str, code: int = EXIT_UNEXPECTED) -> int:
        """统一失败响应：错误码+信息写 stderr 与 stdout(JSON)，返回退出码。"""
        LOGGER.error("%s", msg)
        print(json.dumps({"ok": False, "errorCode": key, "message": msg},
                         ensure_ascii=False))
        return code

    # v2.8：自定义时间范围（--start/--end，成对）
    range_lo: Optional[date] = None
    range_hi: Optional[date] = None
    if args.start or args.end:
        if not (args.start and args.end):
            return _fail("USAGE", "--start 与 --end 必须成对指定（YYYY-MM-DD）", EXIT_USAGE)
        try:
            range_lo = date.fromisoformat(args.start)
            range_hi = date.fromisoformat(args.end)
        except ValueError as exc:
            return _fail("USAGE", f"--start/--end 应形如 YYYY-MM-DD：{exc}", EXIT_USAGE)
        if range_lo > range_hi:
            return _fail("USAGE", f"起始日 {range_lo} 晚于结束日 {range_hi}", EXIT_USAGE)

    # 锚定日：显式 --report-date > 范围结束日（v2.8 默认）> 报错
    report_date: Optional[date] = None
    if args.report_date:
        try:
            report_date = date.fromisoformat(args.report_date)
        except ValueError as exc:
            return _fail("USAGE", f"--report-date 应形如 YYYY-MM-DD：{exc}", EXIT_USAGE)
    elif range_lo is not None:
        report_date = range_hi
    if report_date is None:
        return _fail("USAGE", "请指定 --report-date，或用 --start/--end 指定时间范围", EXIT_USAGE)
    # v2.8.1：显式锚定日必须落在时间范围内（否则标题周次/月份命名与范围数据脱节）
    if range_lo is not None and args.report_date and not (range_lo <= report_date <= range_hi):
        return _fail("USAGE",
                     f"锚定日 {report_date.isoformat()} 在时间范围"
                     f"（{range_lo.isoformat()}~{range_hi.isoformat()}）之外：锚定日必须落在范围内"
                     "（不指定时默认取范围结束日）", EXIT_USAGE)

    if not os.path.exists(args.source_path):
        return _fail("FILE_NOT_FOUND", f"计划源不存在：{args.source_path}", EXIT_IO)

    mode_count = sum(bool(x) for x in (args.all, args.template, args.templates))
    if mode_count > 1:
        return _fail("USAGE", "--all / --template / --templates 三选一，不能同时指定",
                     EXIT_USAGE)

    # 解析本批目标模板列表（函件仍走单 --template）
    selected: List[str] = []
    if args.all:
        selected = ["报告-日", "报告-周", "报告-月"]
    elif args.templates:
        selected = _parse_templates_arg(args.templates)
        if not selected:
            return _fail("USAGE", "--templates 不能为空", EXIT_USAGE)
        for tpl in selected:
            if tpl not in _TEMPLATES:
                return _fail("USAGE",
                             f"未知模板：{tpl}，可用：{' / '.join(_TEMPLATES)}",
                             EXIT_USAGE)
            if tpl in _IS_LETTER:
                return _fail("USAGE",
                             "--templates 仅支持主报告（报告-日/周/月），函件请用 --template",
                             EXIT_USAGE)
    elif args.template:
        selected = [args.template]

    if range_lo is not None:
        if args.all or args.templates:
            return _fail("USAGE",
                         "指定时间范围时不能用 --all / --templates："
                         "按范围长度自动判定单份报告，或用 --template 明确指定",
                         EXIT_USAGE)
        if args.template is None:
            args.template = _pick_range_template(range_lo, range_hi)
            selected = [args.template]
            LOGGER.info("自定义范围 %s~%s（%d 天）自动判定为：%s",
                        range_lo.isoformat(), range_hi.isoformat(),
                        (range_hi - range_lo).days + 1, args.template)
        if args.template == "报告-月" and not _is_full_month(range_lo, range_hi):
            return _fail("USAGE",
                         f"月报需覆盖整自然月：本范围（{range_lo.isoformat()}~"
                         f"{range_hi.isoformat()}）未达整月，不生成月报；"
                         "可改 --template 报告-周，或将范围调整为整月", EXIT_USAGE)
        if args.template == "报告-日" and range_lo != range_hi:
            return _fail("USAGE", "日报仅支持单日范围：本范围跨多日，"
                                  "请改 --template 报告-周/月，或把范围缩到单日", EXIT_USAGE)
    if not selected:
        return _fail("USAGE",
                     "请用 --all、--templates 报告-日,报告-周 或 --template <模板>",
                     EXIT_USAGE)
    for tpl in selected:
        if tpl not in _TEMPLATES:
            return _fail("USAGE", f"未知模板：{tpl}，可用：{' / '.join(_TEMPLATES)}",
                         EXIT_USAGE)

    # 强制：出报告（单份/两份/--all）均须 AI 建议覆盖本批每一份
    from genReportParams import DEFAULT_OFFICE, validate_advice_for_templates
    advice_err = validate_advice_for_templates(selected, args.advice_path)
    if advice_err:
        return _fail("USAGE", advice_err, EXIT_USAGE)

    # 强制：报告与函件均须确认所属机构后传 --office（默认名可原样传入表示已确认）
    if not (isinstance(args.office, str) and args.office.strip()):
        return _fail(
            "USAGE",
            "须确认所属机构后传入 --office（默认「"
            + DEFAULT_OFFICE
            + "」，确认后可原样传入；县公司安监部改为对应安委办全名）",
            EXIT_USAGE)
    args.office = args.office.strip()

    cfg_path = args.config_path
    if cfg_path and not os.path.exists(cfg_path):
        return _fail("FILE_NOT_FOUND", f"配置文件不存在：{cfg_path}", EXIT_IO)

    # v2.10 输出目录规范：成品一律收纳进「承载力报告{YYYYMMDD}」日期子文件夹
    # （按生成当日归档，同一批生成的日/周/月报告合并入同一日期夹），工作区根不散放文件。
    out_root = os.path.abspath(args.out_dir)
    args.out_dir = os.path.join(out_root, f"承载力报告{date.today():%Y%m%d}")
    work_dir = os.path.join(args.out_dir, "_work")
    try:
        os.makedirs(args.out_dir, exist_ok=True)
        os.makedirs(work_dir, exist_ok=True)
    except OSError as exc:
        return _fail("OUTPUT_WRITE_FAILED", f"创建输出目录失败：{exc}", EXIT_IO)

    prefetch_pool = None
    prefetch_futs: list = []
    try:
        if cfg_path is None:
            candidate = os.path.join(SCRIPT_DIR, "..", "config", "capacity_config.json")
            cfg_path = os.path.normpath(candidate)
            if not os.path.exists(cfg_path):
                return _fail("FILE_NOT_FOUND",
                             f"自动查找配置失败：{cfg_path}（请用 --config 指定）",
                             EXIT_IO)
        cfg_path = _ensure_org_cfg(cfg_path, work_dir)
        report_batch = [t for t in selected if t not in _IS_LETTER]
        is_letter = len(selected) == 1 and selected[0] in _IS_LETTER
        if report_batch and not is_letter:
            prefetch_pool, prefetch_futs = _begin_template_prefetch(report_batch)
        if is_letter:
            args.template = selected[0]
            files = _letter_pipeline(args, cfg_path, work_dir, report_date)
            charts = 0
            if not files:
                LOGGER.warning("周期内无命中单位/负责人，未生成函件（非错误）")
        elif len(report_batch) > 1:
            # 多份（--all 或 --templates）：先 calc 一次，再串行 skip_calc
            source = os.path.abspath(args.source_path)
            calc_json = os.path.join(work_dir, "calc_result.json")
            _run(["calcCapacity.py", "--xls", source, "--period", "日",
                  "--out", calc_json, "--config", cfg_path], "承载力量化计算")
            _finish_template_prefetch(prefetch_pool, prefetch_futs)
            prefetch_pool, prefetch_futs = None, []
            files, charts = _run_report_batch(
                args, report_batch, cfg_path, work_dir, report_date,
                range_lo, range_hi, precomputed_calc=True)
        else:
            args.template = report_batch[0]
            args.basename = _report_basename(args.template, report_date)
            specs = _schedules(args.template, report_date, range_lo, range_hi)
            if range_lo is not None:
                LOGGER.info("报告：%s（范围 %s~%s，锚定 %s，图表 %d 张）",
                            args.basename, range_lo.isoformat(),
                            range_hi.isoformat(), report_date.isoformat(), len(specs))
            else:
                LOGGER.info("报告：%s（锚定 %s，图表 %d 张）", args.basename,
                            report_date.isoformat(), len(specs))
            _finish_template_prefetch(prefetch_pool, prefetch_futs)
            prefetch_pool, prefetch_futs = None, []
            final = _pipeline(args, cfg_path, work_dir, report_date, specs,
                              args.chart_mode)
            files = [os.path.abspath(final)]
            charts = len(specs)
        if not args.keep:
            shutil.rmtree(work_dir, ignore_errors=True)
    except RuntimeError as exc:
        return _fail("PIPELINE_FAILED", str(exc), EXIT_INPUT)
    except Exception as exc:  # noqa: BLE001
        return _fail("UNEXPECTED", f"未预期异常：{exc}", EXIT_UNEXPECTED)
    finally:
        _finish_template_prefetch(prefetch_pool, prefetch_futs)
        try:
            import fillReport as _fr
            _fr.end_shared_word()
        except Exception:
            pass

    if args.all:
        tpl_label = "全部(日/周/月)"
    elif len(selected) > 1:
        tpl_label = ",".join(selected)
    else:
        tpl_label = selected[0]
    print(json.dumps({"ok": True,
                      "template": tpl_label,
                      "templates": selected,
                      "office": args.office,
                      "report_date": report_date.isoformat(),
                      "range": [args.start, args.end] if args.start else None,
                      "charts": charts,
                      "file": (files[0] if files else None),
                      "files": files,
                      "work": os.path.abspath(work_dir) if args.keep else None},
                     ensure_ascii=False))
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
