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
  - 计划的任一班组在人员档案中匹配不到 -> 该计划**整条剔除**（不计算、不进报告/函件）。

本模块提供：
  - `clean_team_token / split_team_tokens`：班组名清洗（去 -N 尾缀）与多班组拆分；
  - `build_roster_index`：人员档案 -> {清洗后班组名: RosterTeam}；
  - `resolve_team`：计划班组串 -> {班组列表, 在册人数, 工作负责人真名, 全部成员名}，
    任一班组未命中返回 None；
  - `roster_staff`：按 班组 + 当日不在岗 计算可用人数（供日分母）；
  - `filter_roster_matchable`：剔除班组匹配不到的计划，返回新 dataset。
"""
from __future__ import annotations

import re
from dataclasses import replace
from typing import Any, Dict, Iterable, List, Optional, Set

from . import constants as C
from .models import PersonRecord, WorkDataset

# 多班组分隔符：顿号 / 逗号（全半角）/ 分号 / 斜杠 / 空白
_SPLIT_RE = re.compile(r"[、,，;；/\s]+")
# 班组名尾缀抽调人数（如「变电一次检修二班-2」的 -2，兼容全角/破折号）
_SUFFIX_RE = re.compile(r"[-－—–]\s*\d+\s*$")


def clean_team_token(token: Any) -> str:
    """单个班组 token 清洗：strip + 去尾部「-数字」抽调人数。"""
    return _SUFFIX_RE.sub("", str(token or "").strip()).strip()


def split_team_tokens(team_raw: Any) -> List[str]:
    """把计划班组串拆成去重后的清洗班组名列表（保持出现顺序）。"""
    tokens: List[str] = []
    seen: Set[str] = set()
    for part in _SPLIT_RE.split(str(team_raw or "").strip()):
        tok = clean_team_token(part)
        if tok and tok not in seen:
            seen.add(tok)
            tokens.append(tok)
    return tokens


def build_roster_index(persons: Iterable[PersonRecord]) -> Dict[str, Dict[str, Any]]:
    """人员档案 -> 班组索引。

    返回 ``{清洗后班组名: {"beta_by_name": {姓名: β}, "leaders": [工作负责人真名...],
    "areas": {工区: 人数}, "counties": {县公司: 人数}}}``。β 当前统一按 1.0（FORCE_BETA），
    故 ``beta_by_name`` 之和即该班组真实在册人数。
    """
    index: Dict[str, Dict[str, Any]] = {}
    for p in persons:
        key = clean_team_token(p.team)
        if not key:
            continue
        e = index.setdefault(
            key, {"beta_by_name": {}, "leaders": [], "areas": {}, "counties": {}})
        e["beta_by_name"][p.name] = e["beta_by_name"].get(p.name, 0.0) + float(p.beta)
        if p.is_leader and p.name not in e["leaders"]:
            e["leaders"].append(p.name)
        if p.work_area:
            e["areas"][p.work_area] = e["areas"].get(p.work_area, 0) + 1
        if p.county:
            e["counties"][p.county] = e["counties"].get(p.county, 0) + 1
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
    """剔除班组任一 token 匹配不到人员档案的计划（无变化则原样返回）。

    用户口径：匹配不到的人员 / 班组不计算、不进入报告与函件。人员档案为空时无法匹配
    任何班组，全部计划剔除（杜绝用默认满编数瞎算）。
    """
    if not dataset.plans:
        return dataset
    if not roster:
        C.LOGGER.warning("人员档案为空：无法匹配任何班组，全部计划剔除（不计算）")
        return replace(dataset, plans=())
    kept = tuple(p for p in dataset.plans if resolve_team(roster, p.team) is not None)
    dropped = len(dataset.plans) - len(kept)
    if dropped:
        C.LOGGER.info("班组档案匹配：剔除未匹配计划 %d 条（保留 %d 条）", dropped, len(kept))
    return dataset if dropped == 0 else replace(dataset, plans=kept)
