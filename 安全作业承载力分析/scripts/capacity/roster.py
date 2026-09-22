# -*- coding: utf-8 -*-
"""
班组人员档案匹配（roster：计划 班组 -> 真实在册人数 / 工作负责人真名）
====================================================================

用户口径（实事求是，不得瞎写数据）：
  - 计划「现场施工单位班组」可能是单班组，也可能是「A-2、B-4」多班组拼接（-N 表示
    从该班组抽调 N 人）。分母（班组可用人数 / 工作负责人可用人天）必须来自人员档案
    的**真实在册人数**，不再用任何"默认满编人数"回退。
  - 计划「工作负责人」列填的是「需要的工作负责人数量」（不是姓名）。真名从该计划
    班组/工区/公司对应的人员档案（角色含「工作负责人」者）匹配取得。
  - 计划点了多个班组时，按班组拆开分别计算，不再把各班组在册人数加总成一个分母。
    某个班组在人员档案中匹配不到时，只跳过该班组上的人，其它班组照算。
    一个班组都匹配不到，整条计划才剔除。

本模块提供：
  - `clean_team_token / split_team_tokens`：班组名清洗（去 -N 尾缀）与多班组拆分；
  - `build_roster_index`：人员档案 -> {清洗后班组名: RosterTeam}；
  - `split_team_assignments`：班组串 -> [(清洗后班组名, 本计划抽调人数或 None)]；
  - `resolve_team`：计划班组串 -> {班组列表, 在册人数, 工作负责人真名, 全部成员名}，
    任一班组未命中返回 None（函件等仍用这一严格口径）；
  - `iter_plan_team_slices`：多班组计划拆成每个已匹配班组一条，供日/周期工时计算；
  - `roster_staff`：按 班组 + 当日不在岗 计算可用人数（供日分母）；
  - `filter_roster_matchable`：一个班组都匹配不到的计划整条剔除；多班组里只是部分对不上的保留。
"""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import replace
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple

from . import constants as C
from .models import PersonRecord, PlanRecord, WorkDataset

# 多班组分隔符：顿号 / 逗号（全半角）/ 分号 / 斜杠 / 空白
_SPLIT_RE = re.compile(r"[、,，;；/\s]+")
# 班组名尾缀抽调人数（如「变电一次检修二班-2」的 -2，兼容全角/破折号）
_SUFFIX_RE = re.compile(r"[-－—–]\s*\d+\s*$")
_SUFFIX_NUM_RE = re.compile(r"[-－—–]\s*(\d+)\s*$")


def clean_team_token(token: Any) -> str:
    """单个班组 token 清洗：strip + 去尾部「-数字」抽调人数。"""
    return _SUFFIX_RE.sub("", str(token or "").strip()).strip()


def split_team_tokens(team_raw: Any) -> List[str]:
    """把计划班组串拆成去重后的清洗班组名列表（保持出现顺序）。"""
    return [name for name, _ in split_team_assignments(team_raw)]


def split_team_assignments(team_raw: Any) -> List[Tuple[str, Optional[int]]]:
    """拆开计划班组串。

    返回 ``[(清洗后班组名, 本计划从该班组抽调的人数或 None), ...]``，按出现顺序去重。
    「班组-2」里的 2 是这条计划在该班组上的人数，不是班组名。同一班组出现多次时人数相加。
    没有「-数字」尾缀时，人数为 None（调用方不要把整张计划的总人数抄到每个班组上）。
    """
    order: List[str] = []
    counts: Dict[str, Optional[int]] = {}
    for part in _SPLIT_RE.split(str(team_raw or "").strip()):
        raw = part.strip()
        if not raw:
            continue
        matched = _SUFFIX_NUM_RE.search(raw)
        name = clean_team_token(raw)
        if not name:
            continue
        headcount = int(matched.group(1)) if matched else None
        if name not in counts:
            order.append(name)
            counts[name] = headcount
            continue
        prev = counts[name]
        if headcount is None:
            continue
        counts[name] = headcount if prev is None else prev + headcount
    return [(name, counts[name]) for name in order]


def _resolve_team_kind(counts: Dict[str, int]) -> Optional[str]:
    """班组类型：只有一种非空取值时用它；没填返回 None；几种类型并列第一也返回 None。"""
    if not counts:
        return None
    top = max(counts.values())
    winners = [kind for kind, n in counts.items() if n == top]
    if len(winners) == 1:
        return winners[0]
    return None


def build_roster_index(persons: Iterable[PersonRecord]) -> Dict[str, Dict[str, Any]]:
    """人员档案 -> 班组索引。

    返回 ``{清洗后班组名: {"beta_by_name": {姓名: β}, "leaders": [工作负责人真名...],
    "areas": {工区: 人数}, "counties": {县公司: 人数}, "team_kind": 班组类型或 None}}``。
    β 在 parse_person 里已经按「强制统一β」决定：开着时每人是配置默认值，关掉时是档案作业素养能力。
    ``team_kind`` 取该班组非空「班组类型」的唯一多数；没有或并列时为 None，计算时回退配置默认。
    """
    index: Dict[str, Dict[str, Any]] = {}
    for p in persons:
        key = clean_team_token(p.team)
        if not key:
            continue
        e = index.setdefault(
            key, {"beta_by_name": {}, "leaders": [], "areas": {}, "counties": {},
                  "kind_counts": Counter()})
        e["beta_by_name"][p.name] = e["beta_by_name"].get(p.name, 0.0) + float(p.beta)
        if p.is_leader and p.name not in e["leaders"]:
            e["leaders"].append(p.name)
        if p.work_area:
            e["areas"][p.work_area] = e["areas"].get(p.work_area, 0) + 1
        if p.county:
            e["counties"][p.county] = e["counties"].get(p.county, 0) + 1
        if p.team_kind:
            e["kind_counts"][p.team_kind] += 1
    for e in index.values():
        e["team_kind"] = _resolve_team_kind(e.pop("kind_counts"))
    total = sum(len(e["beta_by_name"]) for e in index.values())
    C.LOGGER.info("班组档案索引：%d 个班组 / %d 人（供真实分母与负责人真名匹配）",
                  len(index), total)
    return index


def resolve_team(roster: Dict[str, Dict[str, Any]], team_raw: Any) -> Optional[Dict[str, Any]]:
    """把计划班组串解析为真实档案汇总。

    返回 ``{"teams": [班组名...], "staff": 在册Σβ, "leaders": [工作负责人真名...],
    "names": [全部成员真名...]}``；班组串为空或**任一班组未命中**档案 -> 返回 None
    （调用方据此把该计划剔除，不计算）。
    """
    tokens = split_team_tokens(team_raw)
    if not tokens:
        return None
    staff = 0.0
    leaders: List[str] = []
    names: List[str] = []
    for t in tokens:
        e = roster.get(t)
        if e is None:
            return None
        staff += sum(e["beta_by_name"].values())
        for n in e["leaders"]:
            if n not in leaders:
                leaders.append(n)
        for n in e["beta_by_name"]:
            if n not in names:
                names.append(n)
    return {"teams": tokens, "staff": staff, "leaders": leaders, "names": names}


def iter_plan_team_slices(plan: PlanRecord,
                          roster: Optional[Dict[str, Dict[str, Any]]] = None
                          ) -> List[PlanRecord]:
    """多班组计划按已匹配班组拆成多条，供日承载力与周期工时分别计算。

    单班组计划原样返回，人数仍用「所需自有人员人数」，分母仍是该班组在册人数。
    多个班组时：只保留人员档案里对得上的班组；对不上的班组跳过，不把整条计划丢掉。
    每个班组的作业人数用班组名后面的「-数字」（本计划从该班组抽出的人），
    不用各班组在册人数相加，也不把整张计划的总人数抄到每一个班组上。
    「工作负责人」列是整张计划的一个数量，没有姓名或身份证号能把它分到某个班组，
    所以只记在按出现顺序第一个匹配上的班组，其余拆分条的负责人数量为 0。
    """
    assigns = split_team_assignments(plan.team)
    if len(assigns) <= 1 or not roster:
        return [plan]
    matched = [(name, n) for name, n in assigns if name in roster]
    if not matched:
        return []
    slices: List[PlanRecord] = []
    for i, (name, headcount) in enumerate(matched):
        members = int(headcount) if headcount is not None else 0
        slices.append(replace(
            plan,
            team=name,
            members=members,
            leader_count=plan.leader_count if i == 0 else 0,
        ))
    return slices


def _plan_has_matched_team(roster: Dict[str, Dict[str, Any]], team_raw: Any) -> bool:
    """至少一个班组 token 能在人员档案里对上。"""
    assigns = split_team_assignments(team_raw)
    return any(name in roster for name, _ in assigns)


def roster_staff(roster: Dict[str, Dict[str, Any]], team_raw: Any,
                 absent: Optional[Set[str]] = None) -> float:
    """计划班组当日可用人数 = Σ(各班组成员 β − 当日不在岗 β)。

    未命中档案的班组按 0 计（不参与分母）；调用方应先用 `resolve_team` 保证班组
    全部命中后再取分母，本函数仅作日维度扣减不在岗的细粒度取值。
    """
    absent = absent or set()
    total = 0.0
    for t in split_team_tokens(team_raw):
        e = roster.get(t)
        if e is None:
            continue
        total += sum(b for n, b in e["beta_by_name"].items() if n not in absent)
    return total


def filter_roster_matchable(dataset: WorkDataset,
                            roster: Dict[str, Dict[str, Any]]) -> WorkDataset:
    """去掉一个班组都匹配不到人员档案的计划（无变化则原样返回）。

    多个班组里只有一部分对不上时，计划保留；计算时 `iter_plan_team_slices`
    只跳过没对上的班组。人员档案为空时无法匹配任何班组，全部计划剔除
    （杜绝用默认满编数瞎算）。
    """
    if not dataset.plans:
        return dataset
    if not roster:
        C.LOGGER.warning("人员档案为空：无法匹配任何班组，全部计划剔除（不计算）")
        return replace(dataset, plans=())
    kept = tuple(p for p in dataset.plans if _plan_has_matched_team(roster, p.team))
    dropped = len(dataset.plans) - len(kept)
    if dropped:
        C.LOGGER.info("班组档案匹配：剔除完全未匹配计划 %d 条（保留 %d 条）", dropped, len(kept))
    return dataset if dropped == 0 else replace(dataset, plans=kept)
