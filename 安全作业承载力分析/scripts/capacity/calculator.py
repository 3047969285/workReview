# -*- coding: utf-8 -*-
"""
承载力量化计算核心（calculator：日 / 周期承载力量化与四级预警）
==============================================================

计算入口一览（想看“怎么算”直接跳到对应函数）：
  - 班组承载力怎么算   -> `calc_team_capacity()`   （内部细分 calc_team_staff /
                           calc_leader_capacity / calc_member_capacity）
  - 工作负责人承载力   -> `calc_leader_capacity()`
  - 班组成员承载力     -> `calc_member_capacity()`
  - 工区承载力（加权） -> `calc_area_capacity()`
  - 日承载力           -> `calc_day()`（入口：逐班组聚合）
  - 周期（周/月/年）   -> `calc_period()`（分桶见 `calc_period_buckets`）
  - 日期×单位矩阵      -> `build_matrix_day()`
  - 周月图/评述口径    -> `period_unit_caps_from_daily()` / `period_manage_caps_from_days()`
  - 四级预警分级       -> `coefficients.classify_alert()`（见 coefficients 模块）

`CapacityService` 为薄调度层：准备数据（β 索引 / 组织架构归一化器）-> 调上述纯函数
-> 组装输出 JSON。输出契约 keys（persons_beta / days / periods / matrix）与原口径
逐键一致，供报告生成直接消费。
"""
from __future__ import annotations

from datetime import date
from typing import Any, Dict, List, Optional, Tuple

from . import constants as C
from .coefficients import build_person_beta, classify_alert, risk_theta
from .org import UnitNormalizer, filter_matchable
from .roster import (build_roster_index, filter_roster_matchable, iter_plan_team_slices,
                      resolve_team, roster_staff, split_team_tokens)
from .models import AbsenceRecord, PersonRecord, PlanRecord, WorkDataset

# 单日班组聚合中间量
TeamAcc = Dict[str, Any]

# =============================================================================
# 辅助
# =============================================================================


def _days_between(start: date, end: date) -> List[date]:
    """闭区间内逐日展开。"""
    days: List[date] = []
    cursor = start
    while cursor <= end:
        days.append(cursor)
        cursor = cursor.fromordinal(cursor.toordinal() + 1)
    return days


def _workdays_between(start: date, end: date) -> List[date]:
    """闭区间内逐日展开，仅保留工作日（剔除周末与法定节假日，见 constants.is_workday）。"""
    return [d for d in _days_between(start, end) if C.is_workday(d)]


def period_unit_caps_from_daily(
    cap_by_day: Dict[str, Dict[str, float]],
    start: date, end: date,
) -> Dict[str, float]:
    """多天单位承载力：Σ(各自然日日承载力%) / 工作日天数。

    日承载力% = 当日工时/当日分母；分母按日稳定时，等价于
    Σ自然日工时 / (工作日天数 × 日分母)，与 calc_period 一致。
    无计划日分子按 0；工作日无论有无计划都计入分母天数。
    区间内无工作日（例如日报落在周六）时，分母改用自然日天数，避免图上变成 0。
    """
    n_work = len(_workdays_between(start, end))
    if n_work <= 0:
        n_work = len(_days_between(start, end))
    sums: Dict[str, float] = {}
    for d in _days_between(start, end):
        for unit, val in (cap_by_day.get(d.isoformat()) or {}).items():
            sums[unit] = sums.get(unit, 0.0) + float(val)
    if n_work <= 0:
        return {unit: 0.0 for unit in sums}
    return {unit: total / n_work for unit, total in sums.items()}


def period_manage_caps_from_days(
    days: Dict[str, Any],
    start: date, end: date,
) -> Dict[str, float]:
    """多天管理承载力：Σ自然日实需人天 / (在册人数 × 工作日天数) × 100。

    与 _mgmt_range_summary / _mgmt_aggregate 同口径。
    区间内无工作日时分母改用自然日天数。
    """
    n_work = len(_workdays_between(start, end))
    if n_work <= 0:
        n_work = len(_days_between(start, end))
    demand: Dict[str, float] = {}
    staff: Dict[str, float] = {}
    for d in _days_between(start, end):
        for unit, info in (days.get(d.isoformat()) or {}).items():
            if not isinstance(info, dict):
                continue
            demand[unit] = demand.get(unit, 0.0) + float(info.get("实需人天") or 0.0)
            if unit not in staff and info.get("在册人数") is not None:
                staff[unit] = float(info.get("在册人数") or 0.0)
    out: Dict[str, float] = {}
    for unit, dem in demand.items():
        avail = staff.get(unit, 0.0) * n_work
        out[unit] = (dem / avail) * 100.0 if avail > 0 else 0.0
    return out


def _day_absent_names(absences: Tuple[AbsenceRecord, ...], d: date) -> set:
    """当日不在岗（在起止范围内）人员姓名集合。"""
    return {a.name for a in absences if a.start <= d <= a.end}


def _on_day_plans(plans: Tuple[PlanRecord, ...], d: date) -> List[PlanRecord]:
    """当日生效的作业计划（计划起止区间覆盖该日）。"""
    return [p for p in plans if p.start <= d <= p.end]


def _team_denom(staff: float) -> float:
    """班组日分母：可用人数 × 标准日工时，带分母保护。"""
    return C.STD_HOURS_PER_DAY * max(staff, C.DENOM_EPS)


# =============================================================================
# 班组承载力（“班组怎么算”入口）
# =============================================================================


def calc_team_staff(
    roster: Dict[str, Dict[str, Any]], team: str, absent: set
) -> float:
    """班组可用人数（人员档案该班组在册 Σβ − 当日不在岗 β）。

    日/周期计算前已用 iter_plan_team_slices 按班拆开，这里看到的是单个班组。
    对不上档案的班组在拆分时跳过；一个班组都对不上的计划已由 filter_roster_matchable 去掉。
    """
    return max(roster_staff(roster, team, absent), 0.0)


def calc_leader_capacity(
    team_acc: TeamAcc, staff: float,
    roster: Dict[str, Dict[str, Any]], team: str,
) -> float:
    """工作负责人承载力指数 f_leader：负责人需求工时 ÷ 负责人可用人天；上限钳制。

    负责人可用人天 = 该(多)班组人员档案中工作负责人真名数 × 标准日工时（真实在册，
    非默认值）；工作负责人为 0 时按极小分母计，百分比会很高，套用配置后的预警落在满载，不是停工档。
    """
    resolved = resolve_team(roster, team)
    leader_count = len(resolved["leaders"]) if resolved else 0
    avail_denom = C.STD_HOURS_PER_DAY * max(leader_count, C.DENOM_EPS)
    return min(team_acc["负责人工时"] / max(avail_denom, C.DENOM_EPS), C.LEADER_LOAD_CAP)


def calc_member_capacity(team_acc: TeamAcc, staff: float) -> float:
    """班组成员承载力指数 f_member：成员工时 ÷ 班组日分母。"""
    denom = _team_denom(staff)
    return team_acc["成员工时"] / denom if team_acc["成员工时"] else 0.0


def _team_kind_for(roster: Dict[str, Dict[str, Any]], team: str) -> str:
    """该班组在人员档案里的班组类型。没有类型或几种类型并列时，用配置默认。"""
    found: List[str] = []
    for token in split_team_tokens(team):
        kind = (roster.get(token) or {}).get("team_kind")
        if kind and kind not in found:
            found.append(kind)
    if len(found) == 1:
        return found[0]
    return C.TEAM_KIND_DEFAULT


def calc_team_capacity(
    team: str, team_acc: TeamAcc,
    roster: Dict[str, Dict[str, Any]],
    normalizer: UnitNormalizer, absent: set,
) -> Dict[str, Any]:
    """【班组承载力】入口：负责人 / 班组成员 / 班组综合，及四级预警与单位归列。

    班组综合按该班组自己的「班组类型」（人员档案），不是全局一个默认值：
    检修施工=max(负责人, 成员)；运检合一=max(负责人, 成员)+巡视；变电运维=操作+巡视。
    档案里没有类型时用配置「默认班组类型」。调控等未单列的类型走检修施工这一支。
    分母是这个班组自己的在册人数，不是多个班组在册人数之和。
    """
    staff = calc_team_staff(roster, team, absent)
    denom = _team_denom(staff)
    f_leader = calc_leader_capacity(team_acc, staff, roster, team)
    f_member = calc_member_capacity(team_acc, staff)
    f_patrol = team_acc["巡视工时"] / denom if team_acc["巡视工时"] else 0.0

    kind = _team_kind_for(roster, team)
    f_team = max(f_leader, f_member)
    if kind == "运检合一":
        f_team = max(f_leader, f_member) + f_patrol
    elif kind == "变电运维":
        f_team = team_acc["操作工时"] / denom + f_patrol

    return {
        "工作负责人承载力": round(f_leader * 100, 1),
        "班组成员承载力": round(f_member * 100, 1),
        "班组承载力": round(f_team * 100, 1),
        "预警": classify_alert(f_team * 100),
        "当日作业项数": len(team_acc["负责人需求"]),
        "负责人": [x["beta"] for x in team_acc["负责人需求"]],
        # 组织架构归一化到报告矩阵单位列（未启用组织架构时为“其他”）
        "unit": normalizer.unit_of(
            (team_acc.get("company"), team_acc.get("work_area"), team)),
    }


def calc_area_capacity(teams_res: Dict[str, Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """【工区承载力】入口：权重 = 班组承载力占当日各班组的比重（与现网口径一致）。"""
    loads = [team["班组承载力"] for team in teams_res.values()]
    if not loads:
        return None
    total = sum(loads)
    if not total:
        return None  # total 为 0（全部轻载 0%）：遵循原口径不输出 area 键
    area = sum(v * (v / total) for v in loads)
    return {"工区承载力": round(area, 1), "预警": classify_alert(area)}


def calc_day(
    dataset: WorkDataset, normalizer: UnitNormalizer,
    beta_map: Dict[str, Dict[str, Dict[str, float]]],
    roster: Dict[str, Dict[str, Any]], d: date,
) -> Dict[str, Any]:
    """【日承载力】入口：单日逐班组聚合负责人 / 班组成员 / 巡视 / 操作需求，再生成工区加权。

    每人每日“实际工时”≤8h 截断（口径：一个人工作时间超过 8 小时按 8 小时计，ΣT>8 时按
    比例压缩）：
      - 负责人按人分组：单日跨多个作业 ΣT>8h 时，加权需求按比例压缩到 8h；
      - 班组成员人均：按“当日班组工作日小时合计 ΣT”为每人实际工时（班组人员同做
        本班组全部作业），成员工时 / 巡视 / 操作占用同口径按 ΣT 比例压缩。
    """
    day_plans = _on_day_plans(dataset.plans, d)
    absent = _day_absent_names(dataset.absences, d)
    teams: Dict[str, TeamAcc] = {}

    for raw_plan in day_plans:
        for plan in iter_plan_team_slices(raw_plan, roster):
            team_acc = teams.setdefault(plan.team, {
                "负责人需求": [], "负责人工时": 0.0, "成员工时": 0.0,
                "巡视工时": 0.0, "操作工时": 0.0, "成员人数": 0,
                "负责人分组": {},
                "成员日小时": 0.0, "巡视日小时": 0.0, "操作日小时": 0.0,
                "company": None, "work_area": None,
            })
            # 记录班组首个可用的公司 / 工区，供组织架构归一化
            if team_acc["company"] is None and plan.company:
                team_acc["company"] = plan.company
            if team_acc["work_area"] is None and plan.work_area:
                team_acc["work_area"] = plan.work_area
            t = plan.hours
            theta = risk_theta(plan.risk, plan.plan_type, plan.content)
            # 工作负责人列 = 需要的负责人数量：负责人需求按该数量计（β 读配置默认值）
            leader_beta = C.DEFAULT_BETA
            team_acc["负责人需求"].append({"hours": t, "theta": theta, "beta": leader_beta,
                                           "count": plan.leader_count})
            # 负责人按"人"分组累计（原始小时 + 加权需求），供 8h 截断；需求按负责人数量计
            leader_key = plan.leader or ""
            lg = team_acc["负责人分组"].setdefault(leader_key, {"hours": 0.0, "weighted": 0.0})
            lg["hours"] += t
            lg["weighted"] += t * theta * leader_beta * plan.leader_count
            # 组员效能统一取 1.0 简化；有档案时可按实际人员项累加
            team_acc["成员工时"] += t * theta * plan.members
            team_acc["成员日小时"] += t  # 班组当日作业小时合计（每人实际工时基数）
            team_acc["成员人数"] += plan.members
            if plan.plan_type in C.PATROL_TYPES:
                team_acc["巡视工时"] += t * plan.members
                team_acc["巡视日小时"] += t
            if plan.plan_type in C.OPERATION_TYPES:
                team_acc["操作工时"] += t * plan.members
                team_acc["操作日小时"] += t

    # 第二遍：每人每日“实际工时”≤8h 截断（ΣT>8 → 按 8/ΣT 比例压缩）
    for team_acc in teams.values():
        lead = 0.0
        for _, lg in team_acc["负责人分组"].items():
            lead += lg["weighted"] * (
                min(1.0, C.MAX_WORK_HOURS / lg["hours"]) if lg["hours"] > 0 else 1.0)
        team_acc["负责人工时"] = lead
        for raw_key, weighted_key in (("成员日小时", "成员工时"),
                                      ("巡视日小时", "巡视工时"),
                                      ("操作日小时", "操作工时")):
            raw = team_acc[raw_key]
            if raw > 0:
                team_acc[weighted_key] *= min(1.0, C.MAX_WORK_HOURS / raw)

    result: Dict[str, Any] = {"date": str(d), "teams": {}}
    for team, team_acc in teams.items():
        result["teams"][team] = calc_team_capacity(
            team, team_acc, roster, normalizer, absent)

    area = calc_area_capacity(result["teams"])
    if area is not None:
        result["area"] = area
    return result


# =============================================================================
# 周期承载力（周 / 月 / 年）
# =============================================================================


def calc_period(dataset: WorkDataset, start: date, end: date,
                roster: Optional[Dict[str, Dict[str, Any]]] = None) -> Dict[str, Any]:
    """【周期承载力】入口：周 / 月 / 年度聚合。

    分子：区间内全部自然日的计划工时（周六日/法定节假日有计划的天照常累加）。
    分母：仅工作日 × 当日在册 Σβ × 标准日工时（is_workday）。
    再取 max(负责人, 成员)（现网简化口径，周期不做每人 8h 截断）。
    """
    agg = {"负责人工时": 0.0, "成员工时": 0.0}
    denom_sum = 0.0
    for d in _days_between(start, end):
        if C.is_workday(d):
            absent = _day_absent_names(dataset.absences, d)
            day_beta_sum = sum(
                person.beta
                for person in dataset.persons
                if person.name not in absent
            )
            staff = max(day_beta_sum, 0.0)
            denom_sum += C.STD_HOURS_PER_DAY * max(staff, C.DENOM_EPS)
        for raw_plan in _on_day_plans(dataset.plans, d):
            for plan in iter_plan_team_slices(raw_plan, roster):
                theta = risk_theta(plan.risk, plan.plan_type, plan.content)
                # 负责人需求按"需要的负责人数量"计（β 读配置默认值；多班组只记在第一条）
                agg["负责人工时"] += plan.hours * theta * C.DEFAULT_BETA * plan.leader_count
                agg["成员工时"] += plan.hours * theta * plan.members

    f_leader = agg["负责人工时"] / max(denom_sum, C.DENOM_EPS)
    f_member = agg["成员工时"] / max(denom_sum, C.DENOM_EPS)
    f_team = max(f_leader, f_member)
    return {
        "start": str(start), "end": str(end),
        "工作负责人承载力": round(f_leader * 100, 1),
        "班组成员承载力": round(f_member * 100, 1),
        "班组承载力": round(f_team * 100, 1),
        "预警": classify_alert(f_team * 100),
    }


def calc_period_buckets(start: date, end: date, mode: str) -> List[Tuple[str, date, date]]:
    """按 周(ISO)/月/年 把起止区间切成有序分桶，返回 (键, 桶首, 桶末)。"""
    buckets: Dict[str, List[date]] = {}
    for d in _days_between(start, end):
        if mode == "周":
            y, w, _ = d.isocalendar()
            key = f"{y}年第{w}周"
        elif mode == "月":
            key = f"{d.year}年{d.month}月"
        else:  # 年
            key = f"{d.year}年"
        buckets.setdefault(key, []).append(d)
    return [(k, ds[0], ds[-1]) for k, ds in buckets.items()]


def build_matrix_day(day_res: Dict[str, Any]) -> Dict[str, Any]:
    """组织架构矩阵一行（单个日期）：当日各班组承载力按单位列（中心/公司）收敛。

    单位承载力 = 该单位当日各班组「班组承载力」的加权均值（权重=各班组承载力占比，
    与工区聚合同一口径），预警按加权结果分级，供报告“日期 × 单位”矩阵表直接消费。
    """
    loads: Dict[str, List[float]] = {}
    for team_res in day_res.get("teams", {}).values():
        unit = team_res.get("unit") or C.ORG_CATEGORY_OTHER
        loads.setdefault(unit, []).append(team_res.get("班组承载力", 0.0))
    out: Dict[str, Any] = {}
    for unit, values in loads.items():
        total = sum(values)
        pct = sum(v * (v / total) for v in values) if total else (max(values) if values else 0.0)
        out[unit] = {"承载力": round(pct, 1), "预警": classify_alert(pct)}
    return out


# =============================================================================
# 管理承载力（F_管理 = 实需人天 ÷ 可供人天）
# =============================================================================


def _manage_occupancy(risk: Optional[str]) -> float:
    """单个现场当日管理占用（人天）：作业风险等级命中「需同进同出」集合占 2，一般到位占 1。"""
    if risk and risk in C.SAME_ENTRY_RISKS:
        return C.OCCUPANCY_SAME_ENTRY
    return C.OCCUPANCY_NORMAL


def managers_by_unit(
    persons: Tuple[PersonRecord, ...], normalizer: UnitNormalizer
) -> Dict[str, int]:
    """各单位“在家管理人员数”：人员档案中角色/职务命中「管理岗关键字」者，
    按 单位(工区) 优先、县公司兜底 归到报告单位列。"""
    out: Dict[str, int] = {}
    for p in persons:
        if not any(k in p.role for k in C.MANAGER_ROLE_KEYWORDS):
            continue
        unit = normalizer.unit_of((p.county, p.work_area))
        out[unit] = out.get(unit, 0) + 1
    return out


def staff_by_unit(
    persons: Tuple["PersonRecord", ...], normalizer: UnitNormalizer
) -> Dict[str, float]:
    """各单位全部在册人员 Σβ（新口径管理承载力分母基数，与班组/负责人承载力同构）。

    归列口径与 managers_by_unit 一致：人员 单位(工区) 优先、县公司兜底。
    """
    out: Dict[str, float] = {}
    for p in persons:
        unit = normalizer.unit_of((p.county, p.work_area))
        out[unit] = out.get(unit, 0.0) + p.beta
    return out


def management_demand_by_day(
    dataset: WorkDataset, normalizer: UnitNormalizer
) -> Dict[date, Dict[str, float]]:
    """逐日各单位实需人天：每现场按风险等级占 2（同进同出）或 1（一般到位），
    多日作业逐日计；单位键 = 作业 公司/工区 归一化到报告单位列。

    v2.3 口径：按计划覆盖的**全部自然日**展开（含周六周日/法定节假日有计划的天），
    供日维度展示与 management.days 逐日值使用；周/月区间聚合见 _mgmt_range_summary
    （分子=全部自然日实需，分母天数只计工作日）。"""
    day_map: Dict[date, Dict[str, float]] = {}
    for plan in dataset.plans:
        unit = normalizer.unit_of((plan.company, plan.work_area))
        occup = _manage_occupancy(plan.risk)
        for d in _days_between(plan.start, plan.end):
            day_map.setdefault(d, {}).setdefault(unit, 0.0)
            day_map[d][unit] += occup
    return day_map


def _mgmt_unit_summary(
    unit: str, demand: float, days_count: int,
    managers: Dict[str, int], staff_map: Dict[str, float],
) -> Dict[str, Any]:
    """单单位区间管理承载力汇总（新口径）：F_管理 = 实需人天 ÷ (单位全部在册人数 × 天数)。

    分母与班组成员/负责人承载力同构：该单位全部可用在册人员（Σβ，来自人员档案真实
    在册；无档案则 0 -> 判为不可供，不再用默认满编数回退）。
    """
    mgr = managers.get(unit, 0)  # 展示用：管理岗计数（真实，无档案则 0，不再默认 10）
    staff = staff_map.get(unit, 0.0)  # 分母基数：全部在册 Σβ（真实，无档案则 0 -> 判为不可供）
    avail = staff * days_count
    cap = (demand / avail) * 100 if avail > 0 else 0.0
    return {
        "实需人天": round(demand, 1),
        "可供人天": round(avail, 1),
        "在册人数": round(staff, 1),
        "在家管理人员数": mgr,
        "管理承载力": round(cap, 1),
        "预警": classify_alert(cap),
    }


def _mgmt_range_summary(
    start: date, end: date, day_map: Dict[date, Dict[str, float]],
    managers: Dict[str, int], staff_map: Dict[str, float],
) -> Dict[str, Dict[str, Any]]:
    """区间聚合：分子按全部自然日叠加实需，分母天数只计工作日。

    单位并集 = 区间内出现单位 ∪ 有管理人员的单位（含仅周末/节假日有作业的单位）。
    """
    all_days = _days_between(start, end)
    workdays = _workdays_between(start, end)
    units = set(managers)
    for d in all_days:
        units.update((day_map.get(d) or {}).keys())
    n_work = len(workdays)
    return {u: _mgmt_unit_summary(
        u, sum((day_map.get(d) or {}).get(u, 0.0) for d in all_days),
        n_work, managers, staff_map)
        for u in sorted(units)}


def build_management(
    dataset: WorkDataset, normalizer: UnitNormalizer, period: str = "日",
) -> Dict[str, Any]:
    """管理承载力全量输出：managers(单位计数) + overall(全区间) + periods(周期分桶) + days(逐日)。"""
    managers = managers_by_unit(dataset.persons, normalizer)
    staff_map = staff_by_unit(dataset.persons, normalizer)
    day_map = management_demand_by_day(dataset, normalizer)
    if not dataset.plans:
        return {"enabled": True,
                "managers": {u: {"在家管理人员数": n,
                                 "在册人数": round(staff_map.get(u, 0.0), 1)}
                             for u, n in managers.items()},
                "overall": {}, "periods": [], "days": {}}
    lo = min(p.start for p in dataset.plans)
    hi = max(p.end for p in dataset.plans)
    buckets = (calc_period_buckets(lo, hi, period)
               if period in ("周", "月", "年") else [(None, lo, hi)])
    periods = [{"period": key, "start": str(b), "end": str(e),
                "units": _mgmt_range_summary(b, e, day_map, managers, staff_map)}
               for key, b, e in buckets]
    days = {}
    for d in sorted(day_map):
        units = {}
        for unit, demand in day_map[d].items():
            staff = staff_map.get(unit, 0.0)
            avail = staff
            cap = (demand / avail) * 100 if avail > 0 else 0.0
            units[unit] = {"实需人天": round(demand, 1),
                           "在册人数": round(staff, 1),
                           "管理承载力": round(cap, 1), "预警": classify_alert(cap)}
        days[str(d)] = units
    return {
        "enabled": True,
        "managers": {u: {"在家管理人员数": n,
                         "在册人数": round(staff_map.get(u, 0.0), 1)}
                     for u, n in managers.items()},
        "overall": _mgmt_range_summary(lo, hi, day_map, managers, staff_map),
        "periods": periods,
        "days": days,
    }


# =============================================================================
# 薄调度层
# =============================================================================


class CapacityService:
    """承载力量化调度（薄）：准备数据（β 索引 / 组织架构归一化器）-> 调纯函数 -> 组装输出 JSON。"""

    def __init__(self, dataset: WorkDataset,
                 org_cfg: Optional[Dict[str, Any]] = None,
                 manage_cfg: Optional[Dict[str, Any]] = None,
                 org_sheet_rows: Optional[List[Dict[str, Any]]] = None) -> None:
        self.dataset = dataset
        self.beta_map, self.leaders = build_person_beta(dataset.persons)
        # 班组档案索引：分母（班组可用人数 / 负责人可用人天）与负责人真名均取自真实人员档案
        self.roster = build_roster_index(dataset.persons)
        # 组织架构归一化器（是否启用由配置「组织架构映射.是否启用」决定）
        self.normalizer = UnitNormalizer(org_cfg)
        # 配置「优先读取源表组织架构」时，用源表组织架构行做数据驱动归列（关键词兜底）
        if org_sheet_rows:
            self.normalizer.load_sheet(org_sheet_rows)
        # 匹配不到组织架构标准单位的计划（归「其他」）不统计：先剔除再算日/矩阵/管理，
        # 保证单位列与正文/明细口径一致（用户口径：匹配不到组织架构的不统计）
        if self.normalizer.enabled:
            self.dataset = filter_matchable(self.dataset, self.normalizer)
        # 一个班组都匹配不到的计划整条剔除。多班组只是部分对不上时保留，
        # 计算时按班组拆开，跳过没对上的人。
        self.dataset = filter_roster_matchable(self.dataset, self.roster)
        # 管理承载力启用开关（配置「输入解析.管理承载力.是否启用」）
        self.manage_enabled = bool((manage_cfg or {}).get("是否启用", False))

    def calc_day(self, d: date) -> Dict[str, Any]:
        return calc_day(self.dataset, self.normalizer, self.beta_map, self.roster, d)

    def calc_period(self, start: date, end: date) -> Dict[str, Any]:
        return calc_period(self.dataset, start, end, self.roster)

    def summarize(self, period: str = "日") -> Dict[str, Any]:
        """全量输出：persons_beta + 逐日 + 周期分桶 + (可选)日期×单位矩阵。

        period 取值：日(默认，整体周期一条) / 周 / 月 / 年（逐桶输出 periods，
        每桶含 period 键）；组织架构「是否启用」时额外输出 matrix（日期×单位列）。
        输出契约稳定，供报告生成消费。
        """
        out: Dict[str, Any] = {"persons_beta": {}, "days": [], "periods": []}
        for team, members in self.beta_map.items():
            out["persons_beta"][team] = {name: beta for name, beta in members.items()}

        # v2.3 口径：日维度（days / matrix / management.days）按**计划覆盖的全部自然日**
        # 展开，周六周日与法定节假日有计划的分班照常参与日计算（“计划中有周末/节假日不要
        # 去掉”）——不在此处剔除。周/月/年累加只对分母用工作日日历：calc_period 与
        # management 区间聚合的分母走 is_workday，分子仍按自然日计划累加。
        day_set = {
            day
            for plan in self.dataset.plans
            for day in _days_between(plan.start, plan.end)
        }
        matrix: Optional[Dict[str, Any]] = None
        if self.normalizer.enabled:
            matrix = {}
        for day in sorted(day_set):
            day_res = self.calc_day(day)
            out["days"].append(day_res)
            if matrix is not None:
                matrix[str(day)] = build_matrix_day(day_res)

        if self.dataset.plans:
            starts = [p.start for p in self.dataset.plans]
            ends = [p.end for p in self.dataset.plans]
            if period in ("周", "月", "年"):
                for key, b_start, b_end in calc_period_buckets(min(starts), max(ends), period):
                    bucket = self.calc_period(b_start, b_end)
                    bucket["period"] = key
                    out["periods"].append(bucket)
            else:
                out["periods"].append(self.calc_period(min(starts), max(ends)))
        if matrix is not None:
            out["matrix"] = matrix
        out["period_mode"] = period if period in ("日", "周", "月", "年") else "日"
        if self.manage_enabled:
            out["management"] = build_management(self.dataset, self.normalizer, period=period)
        return out
