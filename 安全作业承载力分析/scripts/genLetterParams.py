# -*- coding: utf-8 -*-
"""
函件成套参数生成器（技能「安全作业承载力分析」内置工具）
========================================================

把 calcCapacity 计算结果（matrix 日期×单位承载力）+ 作业计划源，按触发规则
筛选命中单位 / 工作负责人，逐封生成 fillReport 可直接消费的填充参数
（replace + tables），覆盖三份官方函件模板的自动链路。

支持模板（--template 取值）：
  - 警示函-满载       周期内单位承载力最大值 >90（含 >100） -> 每命中单位一封
  - 提示函-重载       周期内单位承载力最大值在 (75, 90]      -> 每命中单位一封（两档不重复发）
  - 提示函-长期作业   周期内某班组连续在岗作业天数 ≥ 阈值 -> 每班组一封（负责人真名按班组匹配）

------------------------------------------------------------
一、输入 / 输出
------------------------------------------------------------
  python genLetterParams.py
      --result calc_result.json         # calcCapacity 输出（matrix）
      --source 工作计划.xls             # 计划源（作业清单附表需要原始电压等级等列）
      --template 警示函-满载             # 三函件之一
      --report-date 2026-09-01          # 函件周期锚定日期（按所在 ISO 周取区间）
      --config ../config/capacity_config.json  # 可选：组织架构/长期作业阈值配置
      --out letters.json                # 输出函件参数 JSON

  退出码：0=成功（letters 为空也算成功大门类下无命中） 2=用法错误 3=输入数据错误
         4=文件/IO 错误 1=未知异常

------------------------------------------------------------
二、输出结构
------------------------------------------------------------
  {
    "ok": true, "template": "提示函-重载",
    "period": {"lo": "2026-09-01", "hi": "2026-09-06", "week": 36,
               "文号": "2026年第36号"},
    "letters": [
      {
        "basename": "安全风险提示函（市北中心2026年第36周承载力重载）",
        "template": "提示函-重载",
        "replace": { ... } ,            // fillReport 直接消费
        "tables": [ {"columns": [10列表头], "rows": [作业清单行] } ]
      }
    ]
  }

  无命中单位/负责人时 letters 为空数组（ok=true，退出码 0），由下游自行提示。

------------------------------------------------------------
三、填充要点（占位符与官方模板逐字核对）
------------------------------------------------------------
  - 文号占位（P2）      （20××年第××号） -> （2026年第36号）
  - 统计区间（P8 长键） 20××年××月××日至××月××日/20××年第×周（20××年××月××日至××月××日）
                    -> 2026年第36周（2026年9月1日至9月5日）  [长键先替换清掉斜杠二选一]
  - 单位短语（P8）      ××公司××工区（若……“××班组”）  /  ××公司××中心（若……“××班组”）
                    -> 国网青岛供电公司{单位名}（同时移除模板自带编辑说明括注）
  - 抬头单位（P7）      ××单位： -> {单位名}：
  - 成文/落款日期       20××年××月××日 与 ××××年××月××日 -> 2026年9月5日（周期末日）
  - 长期作业句（P9）    ××公司××工区工作负责人×××连续工作超××天（20××年××月××日至××月××日）
                    -> 国网青岛供电公司{单位}工作负责人{真名们}连续工作超{N}天（9月1日至9月5日）
                       [连续在岗落到「班组」粒度：计划「工作负责人」列是数量非姓名，具体个人
                        无法定位；真名取该班组在人员档案（角色含「工作负责人」者）中全部姓名]
  - 作业清单附表 10 列：序号/工作日期/工作内容/电压等级/风险等级/工作负责人/作业班组/
    自有人员数量/外包施工单位/外包人数；外包两列按「—」（外包专项暂不开展）；
    工作负责人一律输出真名（按班组从人员档案匹配，名字不脱敏）；行 = 命中单位（或该班组）周期内计划。
  - 函件用 resolve_team：多班组串里任一班组匹配不到，整条计划不进入函件。承载力计算另走按班组切片，见维护手册。
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from collections import defaultdict
from datetime import date, timedelta
from typing import Any, Dict, List, Optional, Tuple

from capacity import constants as C
from capacity.config import (CONFIG_SECTION_FORMULA, CONFIG_SECTION_INPUT,
                             DEFAULT_INPUT_CONFIG, _cfg_get, apply_formula_config,
                             load_config)
from capacity.excelLoader import read_sheet_rows, read_xls_persons
from capacity.models import PersonRecord
from capacity.org import UnitNormalizer, load_org_sheet_rows
from capacity.output import atomic_write_text
from capacity.parsing import coerce_date, parse_date, read_dataset_file
from capacity.roster import build_roster_index, resolve_team
from genReportParams import _fmt_d, _iso_week, _unit_cap_by_day

LOGGER = logging.getLogger("genLetterParams")

# ---- 退出码 / 错误码（与全链路一致） -----------------------------------
EXIT_OK, EXIT_UNEXPECTED, EXIT_USAGE, EXIT_INPUT, EXIT_IO = 0, 1, 2, 3, 4
_ERROR_EXIT = {"USAGE": EXIT_USAGE, "INPUT_INVALID": EXIT_INPUT,
               "FILE_NOT_FOUND": EXIT_IO, "OUTPUT_WRITE_FAILED": EXIT_IO,
               "UNEXPECTED": EXIT_UNEXPECTED}

# 支持的三份函件模板（与 genReportParams.TEMPLATES 注册清单一致）
LETTER_TEMPLATES: Tuple[str, ...] = ("警示函-满载", "提示函-重载", "提示函-长期作业")

# 单位短语占位：键 = 模板 P8 原文（含模板自带“编辑说明”括注，整体替换以清除括注），
# 值 = 落到成品的公司+单位名。满载函用“工区”，重载函用“中心”。
_UNIT_PHRASE = {
    "警示函-满载": (
        "××公司××工区（若工区满载，就只描述到工区，若工区未满载，但班组满载，"
        "可以加上“××班组”）",
        "国网青岛供电公司{unit}",
    ),
    "提示函-重载": (
        "××公司××中心（若工区满载，就只描述到中心或县公司，若中心/县公司未满载，"
        "但班组满载，可以加上“××班组”）",
        "国网青岛供电公司{unit}",
    ),
}

# 函件文件名前缀（对应三份官方模板命名规则）
_LETTER_NAME = {
    "警示函-满载": "安全风险警示函-",
    "提示函-重载": "安全风险提示函",
    "提示函-长期作业": "安全风险提示函",
}

# 封面句占位（长期作业模板 P9 原文，含日期括注整体替换）

_STR_KEY = "××公司××工区工作负责人×××连续工作超××天（20××年××月××日至××月××日）"
_RANGE_KEY = "20××年××月××日至××月××日/20××年第×周（20××年××月××日至××月××日）"

# 模板自带编辑说明括注（发函前须整体清除，不进入成品正文）
_NOTE_KEY = "注：若县公司安监部账号，需要将公司名称改为对应县公司名称。"
# 落款单位（模板正文 P12/P13 原文）；--office 指定时动态替换为登陆人所在单位
_OFFICE_KEY = "国网青岛供电公司安全生产委员会办公室"
_OFFICE_DEFAULT = "国网青岛供电公司安全生产委员会办公室"

# 作业清单附表：10 列，与模板作业清单表头逐字一致
LETTER_TABLE_COLUMNS: List[str] = [
    "序号", "工作日期", "工作内容", "电压等级", "风险等级", "工作负责人",
    "作业班组", "自有人员数量", "外包施工单位", "外包人数",
]

_LETTER_FULL_THRESHOLD = 90.0   # 满载 >90（含 >100）
_LETTER_HEAVY_LO = 75.0         # 重载 (75, 90]


# =============================================================================
# 数据准备
# =============================================================================


def _to_int(value: Any, default: int = 0) -> int:
    """数值字段转 int（兼容 float 4.0 / 字符串）；无法解析给默认值。"""
    try:
        return int(round(float(value)))
    except (TypeError, ValueError):
        return default


def _fmt_work_date(start: date, end: date) -> str:
    """作业清单“工作日期”列：单日“9月1日”，多日“9月1日-3日（3天）”。"""
    if start == end:
        return _fmt_d(start)
    days = (end - start).days + 1
    if start.month == end.month:
        return f"{start.month}月{start.day}日-{end.day}日（{days}天）"
    return f"{_fmt_d(start)}-{end.month}月{end.day}日（{days}天）"


# Windows 非法文件名字符 -> 全角等价（`*` 等 ASCII 字符落在文件名里会致 WPS
# SaveAs 报 3011 保存失败，须变换为合法的全角字形；姓名真名下一般无 `*`，其余
# 非法字符仍需转换）
_FILENAME_SAFE = str.maketrans({"\\": "、", "/": "、", ":": "：", "*": "＊",
                                "?": "？", '"': "＂", "<": "〈", ">": "〉",
                                "|": "丨"})


def _safe_filename(name: str) -> str:
    """文件名安全化：把 Windows 非法字符替换为全角等价字符。"""
    return name.translate(_FILENAME_SAFE)


def _job_overlaps(job: Dict[str, Any], lo: date, hi: date) -> bool:
    """作业（dict 视图）是否与区间相交（对应 genReportParams._overlaps）。"""
    return job["start"] <= hi and job["end"] >= lo


def _load_jobs(source_path: str, input_cfg: Dict[str, Any],
               normalizer: UnitNormalizer) -> List[Dict[str, Any]]:
    """统一作业视图：[{start,end,name,content,voltage,risk,leader,team,members,unit}]。

    Excel：直接读原始行（含“电压等级”等展示列，与列映射无关）；
    JSON：从规范化 WorkDataset 读（电压等级缺省用“-”）。
    """
    if str(source_path).lower().endswith((".xls", ".xlsm", ".xlsx")):
        main_sheet = str(_cfg_get(input_cfg, "主工作表", "秋检计划"))
        hr = int(_cfg_get(input_cfg, "表头行号", 0) or 0)
        dr = int(_cfg_get(input_cfg, "数据起始行号", hr + 1) or (hr + 1))
        mode = str(_cfg_get(input_cfg, "时间列格式", "excel序列号"))
        raw_rows = read_sheet_rows(source_path, main_sheet, header_row=hr,
                                   data_start=dr, required=True)
        jobs: List[Dict[str, Any]] = []
        for row in raw_rows:
            start = parse_date(coerce_date(row.get("计划开始时间"), mode))
            end = parse_date(coerce_date(row.get("计划结束时间"), mode))
            if start is None or end is None or end < start:
                continue  # 与 _rows_to_plans 同口径：无排期/坏排期行跳过
            team = row.get("现场施工单位班组") or row.get("业主单位（工区）")
            if not team:
                continue
            jobs.append({
                "start": start, "end": end,
                "name": str(row.get("工程名称") or ""),
                "content": str(row.get("工作内容") or ""),
                "voltage": str(row.get("电压等级") or ""),
                "risk": str(row.get("作业风险等级") or "") or None,
                "leader": str(row.get("工作负责人") or "") or None,
                "team": str(team),
                "team_short": _team_short(normalizer,
                                          (row.get("业主单位（公司）"),
                                           row.get("业主单位（工区）"),
                                           row.get("现场施工单位班组")), team),
                "members": _to_int(row.get("所需自有人员人数（不含工作负责人）")),
                "unit": normalizer.unit_of((row.get("业主单位（公司）"),
                                            row.get("业主单位（工区）"),
                                            row.get("现场施工单位班组"))),
            })
        return jobs
    dataset = read_dataset_file(source_path)
    jobs = []
    for p in dataset.plans:
        jobs.append({
            "start": p.start, "end": p.end, "name": p.name,
            "content": p.content or "", "voltage": "-",
            "risk": p.risk, "leader": p.leader, "team": p.team,
            "team_short": _team_short(normalizer, (p.company, p.work_area, p.team), p.team),
            "members": p.members,
            "unit": normalizer.unit_of((p.company, p.work_area, p.team)),
        })
    return jobs


def _team_short(normalizer: UnitNormalizer, values: Tuple[Any, ...],
                fallback: str) -> str:
    """班组简称（组织架构 sheet「班组简称」列优先；无则源班组名）。"""
    _, _, tm = normalizer.short_name(values)
    return tm if tm and tm != "-" else str(fallback or "")


def _load_roster(source_path: str,
                 input_cfg: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    """按数据源加载人员档案并构建班组索引（与 calcCapacity 同口径）。

    Excel 用 read_xls_persons 读「人员信息」sheet；JSON 直接取 WorkDataset.persons。
    表缺失 / 非法时优雅降级为空索引（此时所有计划因班组匹配不到而被剔除，不计算）。
    """
    if str(source_path).lower().endswith((".xls", ".xlsm", ".xlsx")):
        try:
            persons = read_xls_persons(source_path, input_cfg)
        except Exception as exc:  # noqa: BLE001
            LOGGER.warning("人员档案不可用（%s）：负责人真名无法匹配，相关计划将被剔除",
                           exc)
            return {}
        recs = [PersonRecord(name=p["name"], team=p["team"], beta=float(p["beta"]),
                             is_leader=bool(p["is_leader"]),
                             county=p.get("county", ""),
                             work_area=p.get("work_area", ""),
                             role=p.get("role", "")) for p in persons]
        return build_roster_index(recs)
    try:
        dataset = read_dataset_file(source_path)
    except Exception as exc:  # noqa: BLE001
        LOGGER.warning("数据源人员档案不可用（%s）：负责人真名无法匹配，相关计划将被剔除",
                       exc)
        return {}
    return build_roster_index(dataset.persons)


def _attach_roster(jobs: List[Dict[str, Any]],
                   roster: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
    """给每条作业解析真实班组 + 工作负责人真名；班组匹配不到的整条剔除。

    计划的「工作负责人」列是数量（非姓名）：真名由该计划班组在人员档案（角色含
    「工作负责人」者）匹配取得。写回 ``j["leader"]``（真名顿号拼接，供作业清单列
    展示）与 ``j["teams"]``（清洗后班组 token 列表，供长期作业按班组重键）。
    """
    if not roster:
        if jobs:
            LOGGER.warning("人员档案为空：无法匹配任何班组，全部作业剔除（不计算）")
        return []
    kept: List[Dict[str, Any]] = []
    for j in jobs:
        res = resolve_team(roster, j["team"])
        if res is None:
            continue
        j["leader"] = "、".join(res["leaders"]) if res["leaders"] else ""
        j["teams"] = res["teams"]
        kept.append(j)
    dropped = len(jobs) - len(kept)
    if dropped:
        LOGGER.info("函件：剔除班组匹配不到的计划 %d 条（保留 %d 条）", dropped, len(kept))
    return kept


def _period_meta(report_date: date, lo: Optional[date] = None,
                 hi: Optional[date] = None) -> Dict[str, Any]:
    """函件周期（报告日所在 ISO 周）与文号。

    v2.8：lo/hi（用户自定义时间范围）给定时周期按范围判定（满载/重载峰值、
    长期作业连续在岗、作业清单、函件区间文案全部收敛到范围）；
    week/文号仍按锚定日推导（保留原命名）。
    """
    if lo is None or hi is None:
        lo = _iso_week(report_date)[0]
        hi = lo + timedelta(days=6)
    week = report_date.isocalendar()[1]
    return {"lo": lo, "hi": hi, "week": week, "文号": f"2026年第{week}号"}


def _unit_max_caps(cap_by_day: Dict[str, Dict[str, float]],
                   lo: date, hi: date) -> Dict[str, float]:
    """周期内各单位承载力的最大值（函件以“峰值是否超线”决定发不发）。"""
    max_cap: Dict[str, float] = {}
    cur = lo
    while cur <= hi:
        for u, v in (cap_by_day.get(cur.isoformat()) or {}).items():
            max_cap[u] = max(max_cap.get(u, 0.0), float(v))
        cur += timedelta(days=1)
    return max_cap


def _pick_full_or_heavy(max_cap: Dict[str, float], kind: str) -> List[str]:
    """按档位筛单位：满载 = 峰值 >90（含 >100）；重载 = (75, 90]，两档不重复。"""
    if kind == "满载":
        units = [u for u, v in max_cap.items() if v > _LETTER_FULL_THRESHOLD]
    else:
        units = [u for u, v in max_cap.items()
                 if _LETTER_HEAVY_LO < v <= _LETTER_FULL_THRESHOLD]
    return sorted(units, key=lambda u: (-max_cap.get(u, 0.0), u))


def _long_shift_teams(jobs: List[Dict[str, Any]], lo: date, hi: date,
                      threshold: int) -> List[Tuple[str, int, str]]:
    """周期内连续在岗作业天数 ≥ 阈值的班组 → [(班组, 连续天数, 单位)]，按天降序。

    计划的「工作负责人」列是数量（非姓名），无法定位到具体个人；连续在岗只能落到
    「班组」粒度——班组在某日有作业即视为该班组在岗。多班组作业（A-2、B-4）对每个
    涉及班组分别累计。班组的工作负责人真名在组函时按班组从人员档案匹配（见 _build_letters）。
    """
    by_team: Dict[str, set] = defaultdict(set)
    unit_of: Dict[str, str] = {}
    for j in jobs:
        if not _job_overlaps(j, lo, hi):
            continue
        d0 = max(j["start"], lo)
        d1 = min(j["end"], hi)
        cur = d0
        while cur <= d1:
            for t in (j.get("teams") or [j["team"]]):
                by_team[t].add(cur)
                unit_of.setdefault(t, j["unit"])
            cur += timedelta(days=1)
    result: List[Tuple[str, int, str]] = []
    for team, day_set in by_team.items():
        days = sorted(day_set)
        run = max_run = 1
        for prev, cur in zip(days, days[1:]):
            run = run + 1 if (cur - prev).days == 1 else 1
            max_run = max(max_run, run)
        if max_run >= threshold:
            result.append((team, max_run, unit_of[team]))
    result.sort(key=lambda kv: (-kv[1], kv[0]))
    return result


# =============================================================================
# 逐封函件参数组装
# =============================================================================


def _letter_rows(jobs: List[Dict[str, Any]], cap: int = 300,
                 leader_override: Optional[str] = None) -> List[List[str]]:
    """作业清单 10 列表格行（外包两列「—」，工作负责人输出真名）。

    工作负责人列：默认取作业已匹配到的真实负责人真名（j["leader"]，由 _attach_roster
    按班组从人员档案解析）；长期作业函传 leader_override，统一显示本函所涉班组的工作
    负责人真名。行数超过 cap 时仅列示前 cap 项，并在末行标注“其余略”。
    """
    rows: List[List[str]] = []
    total = len(jobs)
    shown = jobs[:cap]
    for i, j in enumerate(shown, start=1):
        rows.append([
            str(i),
            _fmt_work_date(j["start"], j["end"]),
            j["content"] or j["name"],
            j["voltage"] or "-",
            j["risk"] or "",
            leader_override if leader_override is not None else (j["leader"] or ""),
            j.get("team_short") or j["team"],
            str(j["members"]),
            "—",
            "—",
        ])
    if total > cap:
        LOGGER.warning("作业清单共 %d 行超过上限 %d，仅列示前 %d 项（其余略）",
                       total, cap, cap)
        rows.append(["…", "", f"…（共 {total} 项，仅列示前 {cap} 项，其余略）",
                     "", "", "", "", "", "", ""])
    return rows


def _common_replace(period: Dict[str, Any], unit: str,
                    kind: str, hi: date) -> Dict[str, str]:
    """函件通用占位（文号/区间/单位短语/抬头/成文日期）。kind ∈ 满载|重载。"""
    issued = f"2026年{hi.month}月{hi.day}日"
    unit_key, unit_fmt = _UNIT_PHRASE["警示函-满载" if kind == "满载" else "提示函-重载"]
    return {
        "（20××年第××号）": f"（{period['文号']}）",
        _RANGE_KEY: f"2026年第{period['week']}周（{_fmt_d(period['lo'])}至{_fmt_d(hi)}）",
        unit_key: unit_fmt.format(unit=unit),
        "××单位：": f"{unit}：",
        "20××年××月××日": issued,
        "××××年××月××日": issued,
    }


def _long_letter_replace(period: Dict[str, Any], unit: str,
                         leader: str, days: int, hi: date) -> Dict[str, str]:
    """长期作业函占位（主题句 + 通用项）。"""
    issued = f"2026年{hi.month}月{hi.day}日"
    return {
        "（20××年第××号）": f"（{period['文号']}）",
        _STR_KEY: (f"国网青岛供电公司{unit}工作负责人{leader}"
                   f"连续工作超{days}天（{_fmt_d(period['lo'])}至{_fmt_d(hi)}）"),
        "××单位：": f"{unit}：",
        "20××年××月××日": issued,
        "××××年××月××日": issued,
    }


def _finalize_replace(rep: Dict[str, str], office: Optional[str]) -> Dict[str, str]:
    """成品函件收尾：清除模板编辑说明括注；--office 指定时把落款单位替换为登陆人所在单位。"""
    rep[_NOTE_KEY] = ""
    if office and office != _OFFICE_DEFAULT:
        rep[_OFFICE_KEY] = office
    return rep


def _build_letters(template: str, jobs: List[Dict[str, Any]],
                   cap_by_day: Dict[str, Dict[str, float]],
                   period: Dict[str, Any], threshold: int,
                   roster: Dict[str, Dict[str, Any]],
                   row_cap: int = 300,
                   office: Optional[str] = None) -> List[Dict[str, Any]]:
    """按触发规则逐封组装（字母序稳定，空命中返回 []）。"""
    lo, hi = period["lo"], period["hi"]
    letters: List[Dict[str, Any]] = []
    if template in ("警示函-满载", "提示函-重载"):
        kind = "满载" if template == "警示函-满载" else "重载"
        max_cap = _unit_max_caps(cap_by_day, lo, hi)
        for unit in _pick_full_or_heavy(max_cap, kind):
            unit_jobs = sorted(
                (j for j in jobs if j["unit"] == unit and _job_overlaps(j, lo, hi)),
                key=lambda j: (j["start"], j["name"]))
            stage = "满载" if kind == "满载" else "重载"
            letters.append({
                "basename": _safe_filename(
                    f"{_LETTER_NAME[template]}（{unit}2026年第{period['week']}周"
                    f"承载力{stage}）"),
                "template": template,
                "replace": _common_replace(period, unit, kind, hi),
                "tables": [{
                    "columns": LETTER_TABLE_COLUMNS,
                    "rows": _letter_rows(unit_jobs, row_cap),
                }],
            })
    else:
        # 提示函-长期作业：按连续在岗的班组一封（工作负责人真名从人员档案按班组匹配）
        for team, days, unit in _long_shift_teams(jobs, lo, hi, threshold):
            res = resolve_team(roster, team)
            leaders = res["leaders"] if res else []
            leader_disp = "、".join(leaders) if leaders else team
            team_jobs = sorted(
                (j for j in jobs
                 if team in (j.get("teams") or [j["team"]]) and _job_overlaps(j, lo, hi)),
                key=lambda j: (j["start"], j["name"]))
            letters.append({
                "basename": _safe_filename(
                    f"{_LETTER_NAME[template]}（{unit}{team}长期作业）"),
                "template": template,
                "replace": _long_letter_replace(period, unit, leader_disp, days, hi),
                "tables": [{
                    "columns": LETTER_TABLE_COLUMNS,
                    "rows": _letter_rows(team_jobs, row_cap, leader_override=leader_disp),
                }],
            })
    for lt in letters:
        lt["replace"] = _finalize_replace(lt["replace"], office)
    return letters


# =============================================================================
# 控制器（CLI 入口）
# =============================================================================


def build_main_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="genLetterParams.py",
        description="函件成套参数生成器（技能「安全作业承载力分析」）—— 警示函/提示函",
        add_help=True,
    )
    parser.add_argument("--result", dest="result_path", required=True,
                        metavar="PATH", help="calcCapacity 输出 JSON（须含 matrix）")
    parser.add_argument("--source", dest="source_path", required=True,
                        metavar="PATH", help="计划源：.xls/.xlsm/.xlsx 或 .json")
    parser.add_argument("--template", dest="template", required=True,
                        choices=list(LETTER_TEMPLATES), help="函件模板")
    parser.add_argument("--report-date", dest="report_date", required=True,
                        metavar="YYYY-MM-DD", help="函件周期锚定日期")
    parser.add_argument("--lo", dest="lo", default=None, metavar="YYYY-MM-DD",
                        help="自定义时间范围起始日（须与 --hi 成对）：函件按该范围判定，"
                             "范围外的数据不参与触发")
    parser.add_argument("--hi", dest="hi", default=None, metavar="YYYY-MM-DD",
                        help="自定义时间范围结束日（须与 --lo 成对）")
    parser.add_argument("--office", dest="office", default=None,
                        metavar="NAME",
                        help="落款单位（默认模板原文「国网青岛供电公司安全生产委员会办公室」）；"
                             "登陆人属县公司安监部时传对应县公司安委办全名以动态替换落款")
    parser.add_argument("--config", dest="config_path", default=None,
                        metavar="PATH", help="capacity_config.json 路径")
    parser.add_argument("--out", dest="out_path", required=True, metavar="PATH",
                        help="输出函件参数 JSON")
    return parser


def _setup_logging() -> None:
    logging.basicConfig(level=logging.INFO, stream=sys.stderr,
                        format="[%(levelname)s] %(message)s")


def _fail(err_key: str, msg: str) -> int:
    LOGGER.error("%s", msg)
    print(json.dumps({"ok": False, "errorCode": err_key, "message": msg},
                     ensure_ascii=False))
    return _ERROR_EXIT.get(err_key, EXIT_UNEXPECTED)


def main(argv: Optional[List[str]] = None) -> int:
    args = build_main_parser().parse_args(argv)
    _setup_logging()

    # 1) 配置 / 数据源加载（与 calcCapacity 同路径，保证单位口径一致）
    try:
        cfg = load_config(args.config_path)
        apply_formula_config(cfg)
        input_cfg = cfg.get(CONFIG_SECTION_INPUT) or DEFAULT_INPUT_CONFIG
        org_cfg = input_cfg.get("组织架构映射") or {}
        normalizer = UnitNormalizer(org_cfg)
        sheet_rows = load_org_sheet_rows(args.source_path, input_cfg)
        if sheet_rows:
            normalizer.load_sheet(sheet_rows)
        jobs = _load_jobs(args.source_path, input_cfg, normalizer)
        # 与 calcCapacity 同口径：班组匹配不到人员档案的计划剔除（不计算、不进函件）；
        # 工作负责人真名按班组从人员档案匹配（计划该列是数量，非姓名）
        roster = _load_roster(args.source_path, input_cfg)
        jobs = _attach_roster(jobs, roster)
        # 与 calcCapacity 同口径：匹配不到组织架构的计划（归「其他」）不统计，
        # 不进入函件命中单位/负责人清单
        if normalizer.enabled:
            jobs = [j for j in jobs if j["unit"] != C.ORG_CATEGORY_OTHER]
        with open(args.result_path, "r", encoding="utf-8") as fh:
            result = json.load(fh)
    except Exception as exc:  # noqa: BLE001
        return _fail("FILE_NOT_FOUND" if "存在" in str(exc) else "INPUT_INVALID",
                     f"加载配置/数据源失败：{exc}")

    # 2) 周期与触发规则
    try:
        report_date = date.fromisoformat(args.report_date.strip())
        # v2.8：自定义时间范围（--lo/--hi 成对），函件判定按范围收敛
        lo = hi = None
        if args.lo or args.hi:
            if not (args.lo and args.hi):
                raise ValueError("--lo/--hi 必须成对指定（YYYY-MM-DD）")
            lo = date.fromisoformat(args.lo)
            hi = date.fromisoformat(args.hi)
            if lo > hi:
                raise ValueError(f"范围起始日 {lo} 晚于结束日 {hi}")
        period = _period_meta(report_date, lo=lo, hi=hi)
        formula = cfg.get(CONFIG_SECTION_FORMULA) or {}
        threshold = int(float(_cfg_get(formula, "长期作业连续天数阈值", 5)))
        row_cap = int(float(_cfg_get(formula, "函件作业清单行数上限", 300)))
        cap_by_day = _unit_cap_by_day(result) if isinstance(result, dict) else {}
        letters = _build_letters(args.template, jobs, cap_by_day, period,
                                 threshold, roster, row_cap, office=args.office)
    except Exception as exc:  # noqa: BLE001
        return _fail("INPUT_INVALID", f"生成函件参数失败：{exc}")

    payload = {"ok": True, "template": args.template,
               "period": {k: (v.isoformat() if isinstance(v, date) else v)
                          for k, v in period.items()},
               "letters": letters}
    # 3) 原子写盘
    try:
        atomic_write_text(os.path.abspath(args.out_path),
                          json.dumps(payload, ensure_ascii=False, indent=2))
    except Exception as exc:  # noqa: BLE001
        return _fail("OUTPUT_WRITE_FAILED", f"函件参数写盘失败：{exc}")

    print(json.dumps({"ok": True, "template": args.template,
                      "report_date": report_date.isoformat(),
                      "letters": len(letters),
                      "out": os.path.abspath(args.out_path)}, ensure_ascii=False))
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
