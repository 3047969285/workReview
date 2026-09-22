# -*- coding: utf-8 -*-
"""
报告指标 → fillReport 填充参数生成器（技能「安全作业承载力分析」内置工具）
====================================================================

把 calcCapacity 的计算结果 JSON（days / periods / management）+ 计划源
（JSON 规范化 或 Excel 作业计划）综合为 fillReport 可直接消费的填充参数
（replace + paragraph_replace + tables），逐模板生成正式报告正文与附表。

支持模板（--template 取值见 TEMPLATES）：
  - 报告-日    国网青岛供电公司安全承载力分析报告（日）
  - 报告-周    国网青岛供电公司安全承载力分析报告（周）
  - 报告-月    国网青岛供电公司安全承载力分析报告（月）

------------------------------------------------------------
一、输入 / 输出
------------------------------------------------------------
  python genReportParams.py
      --result calc_result.json        # calcCapacity 输出（days/periods/management）
      --source 工作计划.xls或数据.json   # 计划源：明细表/概况项数/评述“主要开展”用
      --template 报告-月                # 日 / 周 / 月
      --report-date 2026-09-30         # 可选：报告锚定日期，缺省取数据最后一天
      --config ../config/capacity_config.json  # 可选：组织架构/输入解析配置
      --advice advice.json          # 出报告必填：{模板名: 正文}，须含本次模板非空建议
      --office 落款机构名           # runReports 层报告/函件均必填（确认所属机构）
      --out params.json                # 输出 fillReport 参数 JSON

  退出码：0=成功  2=用法错误  3=输入数据错误  4=文件/IO 错误  1=未知异常

------------------------------------------------------------
二、填充要点
------------------------------------------------------------
  1) 概况句：日计划项数 / 二级 / 三级 风险项数（按计划与报告周期交叠统计）；
  2) 明细表：默认三级及以上风险（二级/三级）计划；若当期存在超100% 的计划则只写
     “超过的那些计划”（不限等级），数据量大时按 单位×风险等级 分组抽样保维度；
     日=工作负责人真名+班组成员人数，周=工作负责人真名+班组成员人数+高风险作业时间，
     月=仅高风险作业时间（不输出人员）；
     3) 作业承载力评述段：按单位-周期累加承载力（Σ自然日日值 / 工作日天数）分 满载(≥90)/重载(75~90)，
     附“主要开展工作”（取自该单位周期内计划名）；
  4) 管理承载力段落（真实数据）：F_管理 = 实需人天 ÷ 可供人天，
     日/周用 management.days 逐日聚合，月用 management 全区间 +
     逐 ISO 周聚合（paragraph_replace 逐周写入）；
  5) 日承载力矩阵表（周报附表）：日期 × 报告单位列（优先用结果 matrix，
     否则由 days.teams.unit 收敛）；
  6) 月报 2~5 周作业评述：模板在该处内嵌 5 月硬编码示例段落，生成器改为按
     周真实数据整段覆盖 —— paragraph_replace 用 next_paragraph 模式，锚点
     取每周稳定标题行'（1）作业承载力分析'，替换其下一段（保留段落结束符）。

  注意：管理承载力 / 评述句子一律不带句尾句号 —— 模板在占位后自带
  “。各单位作业承载力如图所示。”等尾部，避免叠出双句号。
  （整段覆盖的周评述例外：因性别替换整段，需自带句尾句号与“如图 N 所示”。）
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
from datetime import date, timedelta
from typing import Any, Dict, List, Optional, Tuple

from capacity import constants as C
from capacity.coefficients import classify_alert
from capacity.config import (CONFIG_SECTION_INPUT, DEFAULT_INPUT_CONFIG,
                             apply_formula_config, load_config)
from capacity.excelLoader import read_xls_dataset
from capacity.org import UnitNormalizer, load_org_sheet_rows, filter_matchable
from capacity.output import atomic_write_text
from capacity.parsing import read_dataset_file
from capacity.calculator import period_unit_caps_from_daily

LOGGER = logging.getLogger("genReportParams")

# =============================================================================
# 模板占位 key（必须与 templates/*.doc 中的原文逐字一致）
# =============================================================================

TEMPLATES: Dict[str, str] = {
    "警示函-满载": "安全风险警示函-（XX单位XX时间承载力满载）",
    "提示函-重载": "安全风险提示函（XX单位XX时间承载力重载）",
    "提示函-长期作业": "安全风险提示函（XX单位XXX长期作业）",
    "报告-日": "国网青岛供电公司2026年X月X日安全承载力分析报告",
    "报告-周": "国网青岛供电公司2026年X月份第X周安全承载力分析报告",
    "报告-月": "国网青岛供电公司2026年X月份安全承载力分析报告",
}

# 作业承载力评述占位块 A（日P39 / 周P36 / 月P38 共用同一句）
KEY_JOBBLOCK_A = ("X1、X2作业承载力满载（90%以上），X1主要开展XXXXX等工作，"
                  "X2主要开展XXXXX等工作。XX1、XX2重载（75%以上），"
                  "XX1主要开展XXXXX等工作，XX2主要开展XXXXX等工作。")
# 周报日分析引导段（P42，模板自带示例句，整段替换）
KEY_WEEK_DAILY_ANALYSIS = ("各单位日计划如表所示：电缆中心5月18、19日重载，"
                           "开展XXXXXXXX；莱西公司5月19-22日、电缆中心"
                           "5月21-23日超100%，请进一步优化调整工作计划安排，"
                           "确保现场安全。")
# 周报/月报第1周标题占位（周P5 / 月P46 公用）
KEY_WEEK_TITLE_GENERIC = "X月份第X周（X月X日至X日）"
# 月报 2~5 周标题（模板硬编码的示例月份，替换为报告月实际周次）
KEY_WEEK_TITLES_HARD = [
    "5月份第2周（5月4日至10日）",
    "5月份第3周（5月11日至17日）",
    "5月份第4周（5月18日至24日）",
    "5月份第5周（5月25日至31日）",
]
# 管理承载力段落锚点（日P40 / 周P37 / 月P42+各周共用同一占位）
ANCHOR_MANAGE = "管理承载力分析XXXXXXXXXX"
# 月度管理承载力总句（月P42，靠更长的 replace key 先吞掉自身锚点）
KEY_MONTH_MANAGE = "X月份，管理承载力分析XXXXXXXXXX"

REPORT_UNITS = list(C.REPORT_UNITS)

_CN_DIGITS = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6}


# =============================================================================
# 通用小工具
# =============================================================================


def _parse_date(value: str) -> date:
    return date.fromisoformat(value.strip())


def _fmt_d(d: date) -> str:
    return f"{d.month}月{d.day}日"


def _overlaps(p, lo: date, hi: date) -> bool:
    return p.start <= hi and p.end >= lo


def _risk_num(risk: Optional[str]) -> Optional[int]:
    """风险等级取数值：优先阿拉伯数字，回退中文数字；无则 None。"""
    if not risk:
        return None
    risk = str(risk)
    m = re.search(r"(\d+)", risk)
    if m:
        return int(m.group(1))
    for ch in risk:
        if ch in _CN_DIGITS:
            return _CN_DIGITS[ch]
    return None


def _iso_week(d: date) -> Tuple[date, date]:
    mon = d - timedelta(days=d.weekday())
    return mon, mon + timedelta(days=6)


def _month_bounds(d: date) -> Tuple[date, date]:
    lo = date(d.year, d.month, 1)
    hi = (date(d.year + 1, 1, 1) if d.month == 12
          else date(d.year, d.month + 1, 1)) - timedelta(days=1)
    return lo, hi


def _month_weeks(d: date) -> List[Tuple[int, date, date]]:
    """月内周次划分：[(序号, 起, 止)]，起止收口在月边界内，最多 5 周。

    初始按「与报告月相交的 ISO 周」枚举（周一为周首，首尾与月外交叉的
    周收口到月边界）。若因此得出 6 周（如 31 日月的首尾各出现一个跨月
    碎片周），把末尾碎片周并入上一周，保证与月报模板固定 5 个周槽对齐。
    """
    lo, hi = _month_bounds(d)
    weeks: List[Tuple[int, date, date]] = []
    cursor = lo - timedelta(days=lo.weekday())  # 月首所在 ISO 周的周一
    ordinal = 1
    while cursor <= hi:
        sun = cursor + timedelta(days=6)
        weeks.append((ordinal, max(cursor, lo), min(sun, hi)))
        cursor = sun + timedelta(days=1)
        ordinal += 1
    while len(weeks) > 5:  # 尾部碎片周（通常 ≤2 天）并入上一周
        last = weeks.pop()
        prev = weeks[-1]
        weeks[-1] = (prev[0], prev[1], last[2])
    return weeks


def _fmt_plan_time(p) -> str:
    """高风险作业时间：日期 + 单元格内换行 +（N天）。
    单日“9月1日”；同月多日“9月1日-8日”；跨月“8月28日-9月2日”。
    （N天）单独换行显示，不再紧跟在日期同行。"""
    if p.start == p.end:
        days = 1
        date_part = f"{_fmt_d(p.start)}"
    elif p.start.month == p.end.month:
        days = (p.end - p.start).days + 1
        date_part = f"{p.start.month}月{p.start.day}日-{p.end.day}日"
    else:
        days = (p.end - p.start).days + 1
        date_part = f"{_fmt_d(p.start)}-{_fmt_d(p.end)}"
    return f"{date_part}\v（{days}天）"


# =============================================================================
# 数据准备
# =============================================================================


def _load_source(path: str, input_cfg: Dict[str, Any]):
    """按扩展名加载计划源（JSON 规范化 或 Excel），返回 WorkDataset。"""
    if str(path).lower().endswith((".xls", ".xlsm", ".xlsx")):
        return read_xls_dataset(path, input_cfg)
    return read_dataset_file(path)


def _load_result(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    if not isinstance(data, dict):
        raise ValueError("结果 JSON 顶层必须是对象")
    return data


def _unit_cap_by_day(result: Dict[str, Any]) -> Dict[str, Dict[str, float]]:
    """日期×单位承载力：优先用结果 matrix，否则从 days.teams.unit 收敛。
    单位值 = 该单位当日各班组「班组承载力」的加权均值（与 build_matrix_day 同口径）。
    """
    matrix = result.get("matrix")
    if matrix:
        return {d: {u: float(v["承载力"]) for u, v in units.items()}
                for d, units in matrix.items()}
    out: Dict[str, Dict[str, float]] = {}
    for day_res in result.get("days") or []:
        loads: Dict[str, List[float]] = {}
        for t in (day_res.get("teams") or {}).values():
            loads.setdefault(t.get("unit") or C.ORG_CATEGORY_OTHER,
                             []).append(float(t.get("班组承载力") or 0.0))
        caps: Dict[str, float] = {}
        for u, vals in loads.items():
            total = sum(vals)
            caps[u] = round(sum(v * (v / total) for v in vals) if total
                            else (max(vals) if vals else 0.0), 1)
        out[day_res["date"]] = caps
    return out


def _mgmt_aggregate(result: Dict[str, Any], lo: date, hi: date) -> Dict[str, Any]:
    """区间（日/周/月）管理承载力聚合：F_管理 = Σ实需人天 ÷ Σ可供人天。

    从 management.days 逐日聚合（对任何 period 粒度都成立）；可供人天按
    单位全部在册人数 × 天数 还原（新口径，与班组/负责人承载力同构）。
    天数只计工作日（与计算层 _mgmt_range_summary 用工作日日历一致），
    周末 / 法定节假日不进分母；分子实需按区间全部自然日累加（有计划的周末/节假日照计）。
    区间内无工作日时天数改用自然日（与图表 period_manage_caps_from_days 一致）。
    """
    mgmt = result.get("management") or {}
    if not mgmt.get("enabled"):
        return {}
    days = mgmt.get("days") or {}
    managers = mgmt.get("managers") or {}

    def _staff(unit: str) -> float:
        info = managers.get(unit)
        if info and info.get("在册人数"):
            return float(info["在册人数"])
        return 0.0  # 无人员档案不瞎设满编数（分母按 0 计，避免虚增承载力）

    day_list = []
    cur = lo
    span_days = 0
    while cur <= hi:
        day_list.append(cur.isoformat())
        if C.is_workday(cur):
            span_days += 1
        cur += timedelta(days=1)
    if span_days <= 0:
        span_days = len(day_list)
    demand: Dict[str, float] = {}
    for d in day_list:
        for u, info in (days.get(d) or {}).items():
            demand[u] = demand.get(u, 0.0) + float(info.get("实需人天") or 0.0)
    if not demand:
        return {}
    total_demand = sum(demand.values())
    total_avail = sum(_staff(u) * span_days for u in demand)
    cap = (total_demand / total_avail) * 100 if total_avail > 0 else 0.0
    # 各单位单独口径（最高者）供句子点名单
    unit_caps = {u: (v / (_staff(u) * span_days)) * 100
                 for u, v in demand.items()
                 if _staff(u) * span_days > 0}
    top_units = sorted(unit_caps.items(), key=lambda kv: kv[1], reverse=True)
    return {"总实需人天": round(total_demand, 1),
            "总可供人天": round(total_avail, 1),
            "管理承载力": round(cap, 1),
            "预警": classify_alert(cap),
            "单位承载力": unit_caps,
            "较高单位": top_units[:3]}


_MGMT_TAIL = {
    "满载": "已接近或达到满载水平，请相关单位强化现场安全管控",
    "停工(超满载)": "已超过满载水平，请相关单位立即优化调整作业安排",
    "重载": "处于重载水平，请做好现场到位监督",
}


def _manage_sentence(prefix: str, agg: Dict[str, Any]) -> str:
    """生成管理承载力评述句（不带句尾句号，模板会补）。"""
    if not agg:
        return f"{prefix}管理承载力总体满足现场作业管控需求"
    cap = agg["管理承载力"]
    tail = _MGMT_TAIL.get(agg["预警"], "现场管理人员满足作业管控需求")
    extra = ""
    if agg.get("较高单位") and agg["较高单位"][0][1] >= 75:
        u, c = agg["较高单位"][0]
        extra = f"，其中{u}达{c:.0f}%"
    return f"{prefix}管理承载力为{cap:.0f}%{extra}，{tail}"


# =============================================================================
# 内容生成
# =============================================================================


def _period_plans(dataset, lo: date, hi: date) -> list:
    return [p for p in dataset.plans if _overlaps(p, lo, hi)]


def _unit_of_plan(normalizer: UnitNormalizer, p) -> str:
    return normalizer.unit_of((p.company, p.work_area, p.team))


def _build_summary(plans, normalizer) -> Dict[str, int]:
    """概况句计数：日计划项数 / 二级 / 三级。"""
    total = len(plans)
    two = sum(1 for p in plans if p.risk == "二级")
    three = sum(1 for p in plans if p.risk == "三级")
    return {"total": total, "二级": two, "三级": three}


def _select_detail_plans(plans) -> list:
    """三级及以上风险（二级/三级）计划全部列出，动态扩行、不设行数上限。

    明细表须与概况句「二级X项 三级Y项」完全一致：有多少条列多少条，Word
    表格按行数自动扩行（fillReport._fill_table_data 动态 Rows.Add），不再做
    超限优先筛选或固定行数抽样截断，保证「文说显示、表就显示」。
    """
    return [p for p in plans
            if (lv := _risk_num(p.risk)) is not None and 1 <= lv <= 3]


def _build_detail_rows(plans, normalizer, with_leader: bool = False,
                       with_members: bool = False,
                       with_time: bool = False) -> List[List[str]]:
    """明细表行：三级及以上风险（二级/三级）全部列出，动态扩行不设上限。

    人员分析：with_leader 输出工作负责人真名（不脱敏），with_members 输出
    班组成员人数（= 计划“所需自有人员人数”）。月报表不传二者，实现“月/年
    不输出人员”。
    """
    rows: List[List[str]] = []
    for p in _select_detail_plans(plans):
        unit = _unit_of_plan(normalizer, p)
        row: List[str] = [str(len(rows) + 1), unit, p.name, str(p.risk or "")]
        if with_leader:
            row.append(p.leader or "")
        if with_members:
            row.append(str(p.members))
        if with_time:
            row.append(_fmt_plan_time(p))
        rows.append(row)
    return rows


def _unit_profile(dataset, normalizer, lo: date, hi: date,
                  cap_by_unit: Dict[str, float]) -> List[Tuple[str, float, List[str]]]:
    """单位承载评述画像：[(单位, 周期平均承载力, 主要工作计划名)]，按承载降序。"""
    works: Dict[str, List[str]] = {}
    for p in _period_plans(dataset, lo, hi):
        u = _unit_of_plan(normalizer, p)
        if p.name not in works.setdefault(u, []):
            works[u].append(p.name)
    profile: List[Tuple[str, float, List[str]]] = []
    for unit, cap in cap_by_unit.items():
        names = works.get(unit, [])
        profile.append((unit, cap, names))
    profile.sort(key=lambda kv: kv[1], reverse=True)
    return profile


def _works_text(names: List[str], limit: int = 8) -> str:
    """主要工作并集文案：“A、B、C等工作”；为空时用通用描述。"""
    if not names:
        return "设备巡视、缺陷消除等日常工作"
    picked = names[:limit]
    tail = "等工作" if len(names) > limit else "等工作"
    return "、".join(picked) + tail


def _build_job_block(profile: List[Tuple[str, float, List[str]]], full_hint: bool = True) -> str:
    """生成作业承载力评述段（满载/重载分句）。full_hint 控制是否含“作业承载力”字样。"""
    full = [(u, w) for u, cap, w in profile if cap >= 90]
    heavy = [(u, w) for u, cap, w in profile if 75 <= cap < 90]
    parts: List[str] = []
    if full:
        units = "、".join(u for u, _ in full)
        works_src = [w for _, w in full]
        works = _works_text([n for ws in works_src for n in ws])
        label = "作业承载力满载（90%以上）" if full_hint else "满载（90%以上）"
        parts.append(f"{units}{label}，主要开展{works}。")
    if heavy:
        units = "、".join(u for u, _ in heavy)
        works = _works_text([n for _, w in heavy for n in w])
        parts.append(f"{units}作业承载力重载（75%以上），主要开展{works}。")
    if not parts:
        # 无满载/重载单位时按实际峰值分级：全区间 <50% 为轻载，50%~75% 为适中（避免轻载期误写“适中”）
        caps = [cap for _, cap, _ in profile]
        if caps and max(caps) < 50:
            return "本期各单位作业承载力总体轻载、作业饱满度不高，未出现满载（90%以上）或重载（75%以上）单位。"
        return "本期各单位作业承载力总体适中，未出现满载（90%以上）或重载（75%以上）单位。"
    return "".join(parts)


def _week_over_analysis(cap_by_day: Dict[str, Dict[str, float]],
                        lo: date, hi: date) -> str:
    """周报日分析句（P42）：列超100% / 重载单位及日期，按连续日期归并。

    v2.3：枚举区间**全部自然日**（含周六周日/法定节假日有计划的天），
    与矩阵表口径一致——周末/节假日当天有作业即正常点名。"""
    over: Dict[str, List[date]] = {}
    heavy: Dict[str, List[date]] = {}
    has_data = False
    peak = 0.0
    cur = lo
    while cur <= hi:
        dkey = cur.isoformat()
        for u, cap in (cap_by_day.get(dkey) or {}).items():
            has_data = True
            peak = max(peak, cap)
            if cap > 100:
                (over.setdefault(u, [])).append(cur)
            elif 75 <= cap <= 100:
                (heavy.setdefault(u, [])).append(cur)
        cur += timedelta(days=1)

    def _fmt_days(u: str, days: List[date]) -> str:
        days = sorted(days)
        groups: List[str] = []
        start = prev = days[0]
        for d in days[1:]:
            if d == prev + timedelta(days=1):
                prev = d
                continue
            groups.append(_fmt_days_group(start, prev))
            start = prev = d
        groups.append(_fmt_days_group(start, prev))
        # group 自带完整日期（含「日」），外层只需括号，避免「（9月1日-4日日）」叠字
        return f"{u}（{'、'.join(groups)}）"

    def _fmt_days_group(a: date, b: date) -> str:
        if a == b:
            return f"{a.month}月{a.day}日"
        if a.month == b.month:
            return f"{a.month}月{a.day}日-{b.day}日"
        return f"{a.month}月{a.day}日-{b.month}月{b.day}日"

    parts = [f"{_fmt_days(u, ds)}超100%" for u, ds in sorted(over.items())]
    parts += [f"{_fmt_days(u, ds)}重载" for u, ds in sorted(heavy.items())]
    if not parts:
        # 无超100%/重载时按区间峰值分级（保留「各单位日计划如表所示」锚点前缀）
        if has_data and peak < 50:
            return "各单位日计划如表所示，各日承载力总体轻载，请各单位持续加强现场安全管控。"
        return "各单位日计划如表所示，各日承载力总体适中，请各单位持续加强现场安全管控。"
    return "各单位日计划如表所示：" + "；".join(parts) + "，请进一步优化调整工作计划安排，确保现场安全。"


def _manage_sentence_per_week(result: Dict[str, Any],
                              weeks: List[Tuple[int, date, date]]) -> List[Dict[str, Any]]:
    """月报逐周管理承载力 paragraph_replace 项（真实 F_管理数据）。

    v2.17：阅读序图号——图1 类型、图2 全月作业、图3 全月管理，
    第 N 周管理图号 = 2N+3（5/7/9/11/13）。无句号，模板段自带「。」保留。
    """
    specs: List[Dict[str, Any]] = []
    for ordinal, w_lo, w_hi in weeks:
        agg = _mgmt_aggregate(result, w_lo, w_hi)
        specs.append({
            "anchor": ANCHOR_MANAGE,
            "occurrence": ordinal,
            "new_text": (_manage_sentence(f"第{ordinal}周", agg)
                         + f"，如图 {2 * ordinal + 3}所示"),
        })
    return specs


def _matrix_table(unit_cap_by_day: Dict[str, Dict[str, float]],
                  lo: date, hi: date) -> Dict[str, Any]:
    """周报日承载力矩阵表：行=日期（一行一个日期）、列=分类(单位)（值取整）。

    用户明确要求「一行一个日期，每个列是分类」：每个自然日占一行，单位
    （班组/专业分类）横向铺开成一列一分类，列头即该分类名；第一列标签
    「日期」。如此即便统计跨两周也不过多行、纵向排布、不横向撑爆 A4 版面；
    列头分类名在窄格中按需自动竖排 2~3 行显示。

    v2.3：行枚举**全部自然日**（lo..hi，含周六周日/法定节假日）——周末/节假日
    当天有计划则显示承载力数值（与计算层 days/matrix 一致），无计划显示 '-'；
    不再把周末有计划的行滤掉。

    单位列由 matrix 实际单位收敛，优先按 REPORT_UNITS 固定顺序排列，
    未在清单中的单位（如组织架构归集的“其他”）排在末尾 —— 避免用
    REPORT_UNITS 硬查 matrix 导致整列空值。行头日期：跨月周带月份，
    同月周只标日。另附 anchor（表前引导段关键词）供 fillReport 在转置
    后列头全为分类名、无法按表头唯一匹配时，精确锚定本表。

    v2.6：spec 附 max_units_per_table=8 分块标记 —— 单位列数超过 8 时
    fillReport 自动把本表纵向拆成上下堆叠的多个子表（每块 ≤8 个单位列，
    首列「日期」行键在每块重复），防止 17+ 单位列总宽撑爆版心；
    单位数 ≤8 时保持单表。
    """
    union: List[str] = []
    for caps in unit_cap_by_day.values():
        for u in caps:
            if u not in union:
                union.append(u)
    ordered = [u for u in REPORT_UNITS if u in union]
    ordered += [u for u in union if u not in ordered]
    days: List[date] = []
    cur = lo
    while cur <= hi:  # v2.3：全部自然日（含周末/节假日，无计划显示 '-'）
        days.append(cur)
        cur += timedelta(days=1)
    cross_month = lo.month != hi.month

    def _hdr(d: date) -> str:
        return f"{d.month}月{d.day}日" if cross_month else f"{d.day}日"

    columns = ["日期"] + ordered
    rows: List[List[Any]] = []
    for d in days:
        row: List[Any] = [_hdr(d)]
        for u in ordered:
            v = (unit_cap_by_day.get(d.isoformat()) or {}).get(u)
            row.append(int(round(v)) if v is not None else "-")
        rows.append(row)
    return {"columns": columns, "rows": rows,
            "anchor": "各单位日计划如表所示",
            "max_units_per_table": 6}


def _week_mean_caps(cap_by_day: Dict[str, Dict[str, float]],
                    w_lo: date, w_hi: date) -> Dict[str, float]:
    """周/多日各单位承载力：Σ自然日日值 / 工作日天数（与 calc_period 同口径）。"""
    return period_unit_caps_from_daily(cap_by_day, w_lo, w_hi)


def _weekly_job_paragraph(cap_by_day: Dict[str, Dict[str, float]], dataset,
                          normalizer, w_lo: date, w_hi: date,
                          fig_no: int) -> str:
    """周作业承载力评述段（整段替换模板硬编码示例用，含句尾句号与图号引用）。"""
    mean = _week_mean_caps(cap_by_day, w_lo, w_hi)
    profile = _unit_profile(dataset, normalizer, w_lo, w_hi, mean)
    block = _build_job_block(profile, full_hint=False).rstrip("。")
    return f"{block}。各单位承载力如图 {fig_no}所示。"


# =============================================================================
# 各模板参数组装
# =============================================================================


def _advice_from_numbers(period: str, summary: Dict[str, int],
                         profile: List[Tuple[str, float, List[str]]],
                         mgmt: Dict[str, Any]) -> str:
    """建议章只引用本次算出的项数、单位和百分比，不采用外部自由正文。"""
    bits = [f"{period}日计划{summary.get('total', 0)}项"]
    if summary.get("二级"):
        bits.append(f"二级风险{summary['二级']}项")
    if summary.get("三级"):
        bits.append(f"三级风险{summary['三级']}项")
    text = "，".join(bits) + "。"
    ranked = sorted(
        ((u, cap) for u, cap, _ in profile if cap > 0),
        key=lambda item: item[1],
        reverse=True,
    )
    if ranked:
        top = "、".join(f"{u}{cap:.1f}%" for u, cap in ranked[:3])
        peak_u, peak_c = ranked[0]
        if peak_c >= 90:
            tone = f"峰值单位{peak_u}为{peak_c:.1f}%，已达满载，请压缩该单位同期作业。"
        elif peak_c >= 75:
            tone = f"峰值单位{peak_u}为{peak_c:.1f}%，处于重载，请统筹该单位作业安排。"
        elif peak_c >= 50:
            tone = f"峰值单位{peak_u}为{peak_c:.1f}%，承载力适中。"
        else:
            tone = f"峰值单位{peak_u}为{peak_c:.1f}%，承载力轻载。"
        text += f"单位承载力居前的是{top}。{tone}"
    cap = mgmt.get("管理承载力") if mgmt else None
    if isinstance(cap, (int, float)):
        text += f"管理承载力为{float(cap):.0f}%。"
    text += "请各单位按实际承载力安排现场作业。"
    return text


def build_daily_params(result, dataset, normalizer,
                       report_date: date,
                       advice: Optional[str] = None,
                       lo: Optional[date] = None,
                       hi: Optional[date] = None) -> Dict[str, Any]:
    # v2.8：自定义时间范围给定时按其过滤（runReports 已校验日报范围为单日）
    d_lo = lo if lo is not None else report_date
    d_hi = hi if hi is not None else report_date
    p = _period_plans(dataset, d_lo, d_hi)
    summary = _build_summary(p, normalizer)
    cap_by_day = _unit_cap_by_day(result)
    caps = cap_by_day.get(d_lo.isoformat()) or {}
    profile = _unit_profile(dataset, normalizer, d_lo, d_hi, caps)
    mgmt = _mgmt_aggregate(result, d_lo, d_hi)
    replace = {
        "2026年X月X日": f"2026年{report_date.month}月{report_date.day}日",
        "日计划X项": f"日计划{summary['total']}项",
        "二级风险作业X项": f"二级风险作业{summary['二级']}项",
        "三级风险作业X项": f"三级风险作业{summary['三级']}项",
        KEY_JOBBLOCK_A: _build_job_block(profile),
        KEY_ADVICE_BODY: _advice_from_numbers("当日", summary, profile, mgmt),
    }
    detail_rows = _build_detail_rows(p, normalizer,
                                     with_leader=True, with_members=True)
    if not detail_rows:
        # 本周期无三级及以上（或超限）计划可列：概况句不再以「…明细表如下：」
        # 悬挂，直接自然收尾；空明细表整体删除由 fillReport 按 rows=[] 处理。
        replace["三级及以上风险工作明细表如下："] = (
            "当日三级及以上风险作业不单列明细。")
        replace[KEY_ADVICE_BODY] = _advice_from_numbers(
            "当日", summary, profile, mgmt)
    return {
        "replace": replace,
        "paragraph_replace": [
            {"anchor": ANCHOR_MANAGE, "occurrence": 1,
             "new_text": _manage_sentence("本日", mgmt)},
        ],
        "tables": [
            {
                "columns": ["序号", "单位", "作业内容", "风险等级",
                            "工作负责人", "班组成员人数"],
                "rows": detail_rows,
            }
        ],
    }


def build_weekly_params(result, dataset, normalizer,
                        report_date: date,
                        advice: Optional[str] = None,
                        lo: Optional[date] = None,
                        hi: Optional[date] = None) -> Dict[str, Any]:
    # v2.8：自定义时间范围给定时按范围出报告（概况/明细/评述/日分析/矩阵/
    # 管理承载全部收敛到范围，范围外数据不进汇总）；周次命名仍按锚定日推导。
    if lo is not None and hi is not None:
        wk_lo, wk_hi = lo, hi
    else:
        wk_lo, wk_hi = _iso_week(report_date)
    ordinal = next((o for o, a, b in _month_weeks(report_date) if a <= report_date <= b), 1)
    p = _period_plans(dataset, wk_lo, wk_hi)
    summary = _build_summary(p, normalizer)
    cap_by_day = _unit_cap_by_day(result)
    caps = _week_mean_caps(cap_by_day, wk_lo, wk_hi)
    profile = _unit_profile(dataset, normalizer, wk_lo, wk_hi, caps)
    mgmt = _mgmt_aggregate(result, wk_lo, wk_hi)
    m = report_date.month
    replace = {
        "2026年X月份第X周": f"2026年{m}月份第{ordinal}周",
        # 只去掉周报概况这一处日期范围的全角括号；月报各周标题仍保留括号。
        KEY_WEEK_TITLE_GENERIC:
            f"{m}月份第{ordinal}周{_fmt_d(wk_lo)}至{_fmt_d(wk_hi)}",
        "日计划X项": f"日计划{summary['total']}项",
        "二级风险作业X项": f"二级风险作业{summary['二级']}项",
        "三级风险作业X项": f"三级风险作业{summary['三级']}项",
        KEY_JOBBLOCK_A: _build_job_block(profile),
        KEY_WEEK_DAILY_ANALYSIS: _week_over_analysis(cap_by_day, wk_lo, wk_hi),
        KEY_ADVICE_BODY: _advice_from_numbers("本周", summary, profile, mgmt),
    }
    detail_rows = _build_detail_rows(p, normalizer,
                                     with_leader=True, with_members=True,
                                     with_time=True)
    if not detail_rows:
        replace["三级及以上风险工作明细表如下："] = (
            "本周三级及以上风险作业不单列明细。")
        replace[KEY_ADVICE_BODY] = _advice_from_numbers(
            "本周", summary, profile, mgmt)
    # 周跨月时标题周也跨月（罕见），仍按报告星期号处理即可
    return {
        "replace": replace,
        "paragraph_replace": [
            # 模板 P37 自带「2.本周」标题，此处前缀留空避免叠出「本周本周」
            {"anchor": ANCHOR_MANAGE, "occurrence": 1,
             "new_text": _manage_sentence("", mgmt)},
        ],
        "tables": [
            {
                "columns": ["序号", "单位", "作业内容", "风险等级",
                            "工作负责人", "班组成员人数", "高风险作业时间"],
                "rows": detail_rows,
            },
            _matrix_table(cap_by_day, wk_lo, wk_hi),
        ],
    }


def build_monthly_params(result, dataset, normalizer,
                         report_date: date,
                         advice: Optional[str] = None,
                         lo: Optional[date] = None,
                         hi: Optional[date] = None) -> Dict[str, Any]:
    m = report_date.month
    # v2.8：自定义时间范围给定时按其过滤（runReports 已校验月报范围=整自然月）
    if lo is not None and hi is not None:
        lo, hi = lo, hi
    else:
        lo, hi = _month_bounds(report_date)
    weeks = _month_weeks(report_date)
    p = _period_plans(dataset, lo, hi)
    summary = _build_summary(p, normalizer)
    cap_by_day = _unit_cap_by_day(result)

    # 月总单位承载评述（Σ自然日 / 工作日天数，与图表同口径）
    caps = period_unit_caps_from_daily(cap_by_day, lo, hi)
    profile = _unit_profile(dataset, normalizer, lo, hi, caps)
    w1 = weeks[0] if weeks else None

    mgmt = _mgmt_aggregate(result, lo, hi)

    replace: Dict[str, str] = {
        "2026年X月份": f"2026年{m}月份",
        "X月份": f"{m}月份",
        "日计划X项": f"日计划{summary['total']}项",
        "二级风险作业X项": f"二级风险作业{summary['二级']}项",
        "三级风险作业X项": f"三级风险作业{summary['三级']}项",
        KEY_JOBBLOCK_A: _build_job_block(profile),
        # v2.12：全月管理评述句尾引用图2 全月管理承力图（无句号，模板段自带「。」保留）
        KEY_MONTH_MANAGE: _manage_sentence(f"全月", mgmt) + "，如图 3所示",
        KEY_ADVICE_BODY: _advice_from_numbers("本月", summary, profile, mgmt),
    }
    detail_rows = _build_detail_rows(p, normalizer,
                                     with_leader=False, with_time=True)
    if not detail_rows:
        replace["三级及以上风险工作明细表如下："] = (
            "当月三级及以上风险作业不单列明细。")
        replace[KEY_ADVICE_BODY] = _advice_from_numbers(
            "本月", summary, profile, mgmt)
    # 第1周标题（通用占位）与 2~5 周标题（模板硬编码 5 月示例）
    if w1:
        replace[KEY_WEEK_TITLE_GENERIC] = (
            f"{m}月份第{w1[0]}周（{_fmt_d(w1[1])}至{_fmt_d(w1[2])}）")
        # 月模板第1周标题写死「第1周」而非「第X周」，用更长 key 抢在「X月份」前整段填日期，
        # 否则残留“X月X日至X日”占位（原始 bug：KEY_WEEK_TITLE_GENERIC 匹配失败）。
        replace["X月份第1周（X月X日至X日）"] = (
            f"{m}月份第{w1[0]}周（{_fmt_d(w1[1])}至{_fmt_d(w1[2])}）")
    for i, hard in enumerate(KEY_WEEK_TITLES_HARD, start=2):
        if len(weeks) >= i:
            o, w_lo, w_hi = weeks[i - 1]
            replace[hard] = f"{m}月份第{o}周（{_fmt_d(w_lo)}至{_fmt_d(w_hi)}）"

    # 周作业承载力评述（第1~5周）：整段覆盖模板硬编码的占位/5 月示例段落。
    # 周槽序号 section 对应稳定标题行'（1）作业承载力分析'第 section 处命中 →
    # 替换其下一段；图号（v2.17）：第 N 周作业图号 = 2N+2：4/6/8/10/12。
    job_specs: List[Dict[str, Any]] = []
    for section in range(1, min(5, len(weeks)) + 1):
        _, w_lo, w_hi = weeks[section - 1]
        job_specs.append({
            "anchor": "（1）作业承载力分析",
            "occurrence": section,
            "next_paragraph": True,
            "new_text": _weekly_job_paragraph(cap_by_day, dataset, normalizer,
                                              w_lo, w_hi, 2 * section + 2),
        })

    return {
        "replace": replace,
        "paragraph_replace": _manage_sentence_per_week(result, weeks) + job_specs,
        "tables": [
            {
                "columns": ["序号", "单位", "作业内容", "风险等级", "高风险作业时间"],
                "rows": detail_rows,
            }
        ],
        "trim_month_weeks": len(weeks),
    }


_BUILDERS = {
    "报告-日": build_daily_params,
    "报告-周": build_weekly_params,
    "报告-月": build_monthly_params,
}


# =============================================================================
# AI 辅助建议注入（--advice）
# =============================================================================

# 占位符 key（模板文末「四、辅助建议与风险管控提示」章占位段；由 fillReport
# 全文 replace 注入）。键=模板占位段完整壳文本（含引导语与「建议占位符：」），
# 整段替换成纯建议正文，避免只换内层 token 时残留「建议占位符：」字样。
# 模板占位段原文（三份报告模板一致）：
#   （本期辅助建议由研判人员结合承载力数据综合提出。建议占位符：ADVICE_BODY_2026）
KEY_ADVICE_BODY = "（本期辅助建议由研判人员结合承载力数据综合提出。建议占位符：ADVICE_BODY_2026）"

# 未提供 --advice / 该模板缺条目时的兜底默认句（保证占位永不残留）
DEFAULT_ADVICE: Dict[str, str] = {
    "报告-日": "本期各单位承载力总体平稳。请各单位持续关注重点日承载力变化，"
              "对承载力偏高单位提前统筹作业计划、合理调配人力，确保作业安全可控。",
    "报告-周": "本周总体承载力处于平稳区间。请各工区、班组持续跟踪重载、满载单位，"
              "统筹安排下周作业计划，避免同类高风险作业集中开展。",
    "报告-月": "本月总体承载力平稳可控。请各单位结合月度承载力分布，前瞻统筹下月作业"
              "计划，对连续高强度作业的班组与工作负责人合理安排轮换与休息，确保作业安全。",
}


def _resolve_advice(template: str, advice: Optional[Dict[str, str]]) -> str:
    """从 --advice JSON（{模板名: 建议正文}）解析当前模板建议文本。

    未提供该模板条目 / 条目为空 / 未传 --advice 时回退统一默认句。
    已传 --advice 但缺本模板键时打告警（禁止静默用兜底冒充已分析）。
    """
    if isinstance(advice, dict):
        text = advice.get(template)
        if isinstance(text, str) and text.strip():
            return text.strip()
        LOGGER.warning(
            "--advice 缺少模板「%s」的真实建议，将回退 DEFAULT_ADVICE；"
            "正式交付须为 报告-日/报告-周/报告-月 均填写分析正文。",
            template)
    return DEFAULT_ADVICE.get(template, DEFAULT_ADVICE["报告-日"])


_REQUIRED_ADVICE_TEMPLATES = ("报告-日", "报告-周", "报告-月")

# 落款机构默认名（用户确认后原样传入 --office 即视为已确认）
DEFAULT_OFFICE = "国网青岛供电公司安全生产委员会办公室"


def validate_advice_for_templates(
        templates: List[str], advice_path: Optional[str]) -> Optional[str]:
    """校验 --advice 覆盖本次将生成的每一份报告模板（单份/两份/--all 同一规则）。"""
    need = [t for t in templates if t in _REQUIRED_ADVICE_TEMPLATES]
    if not need:
        return None
    if not advice_path:
        return ("出报告必须提供 --advice，且 JSON 含本次模板的非空建议正文："
                + "、".join(need)
                + "（禁止缺省回退 DEFAULT_ADVICE）")
    payload = load_advice_payload(advice_path)
    if not isinstance(payload, dict):
        return f"--advice 无法读取或格式错误：{advice_path}"
    missing = []
    for key in need:
        text = payload.get(key)
        if not (isinstance(text, str) and text.strip()):
            missing.append(key)
    if missing:
        return ("--advice 缺少真实建议条目："
                + "、".join(missing)
                + "（本次生成的每一份报告都必须有基于当次 calc 的分析正文）")
    return None


def validate_advice_for_all(advice_path: Optional[str]) -> Optional[str]:
    """兼容旧名：等价于校验日/周/月三份。"""
    return validate_advice_for_templates(list(_REQUIRED_ADVICE_TEMPLATES), advice_path)


def load_advice_payload(advice_path: Optional[str]) -> Optional[Dict[str, str]]:
    """读取 --advice JSON（可选）。文件缺失/损坏时告警返回 None，由解析器回退默认句。"""
    if not advice_path:
        return None
    p = os.path.abspath(advice_path)
    if not os.path.exists(p):
        LOGGER.warning("--advice 文件不存在：%s，本模板建议用默认句。", p)
        return None
    with open(p, "r", encoding="utf-8") as fh:
        payload = json.load(fh)
    if not isinstance(payload, dict):
        LOGGER.warning("--advice 应为 {模板名: 建议正文} 的 JSON 对象，收到 %s，"
                       "本模板建议用默认句。", type(payload).__name__)
        return None
    return payload


# =============================================================================
# 控制器（CLI 入口）
# =============================================================================


def build_main_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="genReportParams.py",
        description="报告指标 → fillReport 填充参数生成器（技能「安全作业承载力分析」）",
        add_help=True,
    )
    parser.add_argument("--result", dest="result_path", required=True,
                        metavar="PATH", help="calcCapacity 输出 JSON")
    parser.add_argument("--source", dest="source_path", required=True,
                        metavar="PATH", help="计划源：.json 规范化 或 .xls/.xlsm/.xlsx")
    parser.add_argument("--template", dest="template", required=True,
                        choices=list(_BUILDERS.keys()), help="报告模板")
    parser.add_argument("--report-date", dest="report_date", default=None,
                        metavar="YYYY-MM-DD", help="报告锚定日期（缺省取数据最后一天）")
    parser.add_argument("--lo", dest="lo", default=None, metavar="YYYY-MM-DD",
                        help="自定义时间范围起始日（须与 --hi 成对）：报告按该范围判定，"
                             "范围外的数据不进报告汇总")
    parser.add_argument("--hi", dest="hi", default=None, metavar="YYYY-MM-DD",
                        help="自定义时间范围结束日（须与 --lo 成对）")
    parser.add_argument("--config", dest="config_path", default=None,
                        metavar="PATH", help="capacity_config.json 路径")
    parser.add_argument("--advice", dest="advice_path", default=None,
                        metavar="PATH",
                        help="AI 辅助建议 JSON（可选）：{模板名: 建议正文}，如 "
                             "{\"报告-日\":\"…\",\"报告-周\":\"…\",\"报告-月\":\"…\"}；"
                             "注入模板文末「四、辅助建议与风险管控提示」占位段；"
                             "缺条目/未传时用统一默认句，保证占位不残留")
    parser.add_argument("--out", dest="out_path", required=True, metavar="PATH",
                        help="输出 fillReport 参数 JSON")
    return parser


def _setup_logging() -> None:
    logging.basicConfig(level=logging.INFO, stream=sys.stderr,
                        format="[%(levelname)s] %(message)s")


def main(argv: Optional[List[str]] = None) -> int:
    args = build_main_parser().parse_args(argv)
    _setup_logging()
    err = {"ok": False, "errorCode": "UNEXPECTED", "message": ""}

    # 1) 配置与数据源加载（与 calcCapacity 同路径，保证口径一致）
    try:
        cfg = load_config(args.config_path)
        apply_formula_config(cfg)
        input_cfg = cfg.get(CONFIG_SECTION_INPUT) or DEFAULT_INPUT_CONFIG
        org_cfg = input_cfg.get("组织架构映射") or {}
        normalizer = UnitNormalizer(org_cfg)
        # 组织架构：与 calcCapacity 同口径，配置启用「优先读取源表组织架构」时读 sheet 归列
        sheet_rows = load_org_sheet_rows(args.source_path, input_cfg)
        if sheet_rows:
            normalizer.load_sheet(sheet_rows)
        dataset = _load_source(args.source_path, input_cfg)
        # 与 calcCapacity 同口径：匹配不到组织架构的计划（归「其他」）不统计，
        # 使明细表条目数与正文/矩阵单位列严格一致
        if normalizer.enabled:
            dataset = filter_matchable(dataset, normalizer)
        result = _load_result(args.result_path)
    except Exception as exc:  # noqa: BLE001
        err.update({"errorCode": "FILE_NOT_FOUND" if "存在" in str(exc) else "INPUT_INVALID",
                    "message": f"加载配置/数据源失败：{exc}"})
        LOGGER.error("%s", err["message"])
        print(json.dumps(err, ensure_ascii=False))
        return 3

    # 2) 报告锚定日期：缺省取数据最后一天
    try:
        if args.report_date:
            report_date = _parse_date(args.report_date)
        else:
            days = result.get("days") or []
            if not days:
                raise ValueError("结果 JSON 无 days，无法确定报告日期")
            report_date = _parse_date(max(d["date"] for d in days))
        # v2.8：自定义时间范围（--lo/--hi 成对）
        lo = hi = None
        if args.lo or args.hi:
            if not (args.lo and args.hi):
                raise ValueError("--lo/--hi 必须成对指定（YYYY-MM-DD）")
            lo = _parse_date(args.lo)
            hi = _parse_date(args.hi)
            if lo > hi:
                raise ValueError(f"范围起始日 {lo} 晚于结束日 {hi}")
        builder = _BUILDERS[args.template]
        advice_payload = load_advice_payload(args.advice_path)
        advice_text = _resolve_advice(args.template, advice_payload)
        params = builder(result, dataset, normalizer, report_date,
                         advice=advice_text, lo=lo, hi=hi)
    except Exception as exc:  # noqa: BLE001
        err.update({"errorCode": "INPUT_INVALID", "message": f"生成参数失败：{exc}"})
        LOGGER.error("%s", err["message"])
        print(json.dumps(err, ensure_ascii=False))
        return 3

    # 3) 原子写盘
    out_as_doc = args.out_path.lower().endswith(".doc")
    if out_as_doc:  # 参数输出应为 .json
        err.update({"errorCode": "INPUT_INVALID",
                    "message": "--out 应为 .json 路径（fillReport 参数文件）"})
        print(json.dumps(err, ensure_ascii=False))
        return 2
    try:
        atomic_write_text(os.path.abspath(args.out_path),
                          json.dumps(params, ensure_ascii=False, indent=2))
    except Exception as exc:  # noqa: BLE001
        err.update({"errorCode": "OUTPUT_WRITE_FAILED", "message": f"参数写盘失败：{exc}"})
        LOGGER.error("%s", err["message"])
        print(json.dumps(err, ensure_ascii=False))
        return 4

    print(json.dumps({"ok": True, "template": args.template,
                      "report_date": report_date.isoformat(),
                      "out": os.path.abspath(args.out_path)},
                     ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
