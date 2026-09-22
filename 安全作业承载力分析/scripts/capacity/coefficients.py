# -*- coding: utf-8 -*-
"""
系数解析与预警分级（coefficients：β 人员效能 / θ 作业风险 / 四级预警）
====================================================================

- `risk_theta()`：作业风险系数 θ（工作内容或作业类型取值为「装表接电」时 0.2；无风险计 0；未知按四级回退）；
- `build_person_beta()`：由人员档案构建 班组->姓名->β 索引与 班组->负责人β 列表；
- `classify_alert()`：按 `ALERT_RULES` 降序命中。`apply_formula_config` 之后阈值只有 90/75/50，大于 90 为满载，大于 100% 仍是满载，没有停工档。
"""
from __future__ import annotations

from typing import Tuple

from . import constants as C
from .models import PersonRecord


def _is_meter_install(value: Any) -> bool:
    """取值去掉首尾空白后正好是「装表接电」。含「装表」但不是这个词的不算。"""
    return str(value).strip() == "装表接电" if value not in (None, "") else False


def risk_theta(risk_level: Any, plan_type: Any = None, content: Any = None) -> float:
    """作业风险系数 θ：无风险等级计 0；未知等级按四级(1.0)回退。

    工作内容是「装表接电」，或作业类型的取值是「装表接电」时，不看作业风险等级，
    θ 取 RISK_COEF「装表接电」（配置为 0.2）。其它含「装表」的文字按作业风险等级取值。
    """
    if _is_meter_install(content) or _is_meter_install(plan_type):
        return float(C.RISK_COEF.get("装表接电", 0.2))
    if risk_level is None or risk_level in ("", "无"):
        return C.RISK_COEF_NONE
    return C.RISK_COEF.get(str(risk_level), C.RISK_COEF_DEFAULT)


def build_person_beta(
    persons: Tuple[PersonRecord, ...]
) -> Tuple[Dict[str, Dict[str, Dict[str, float]]], Dict[str, List[float]]]:
    """由人员档案构建：{班组: {姓名: {"beta": v}}} 与 {班组: [负责人β,...]}。"""
    beta_map: Dict[str, Dict[str, Dict[str, float]]] = {}
    leaders: Dict[str, List[float]] = {}
    for person in persons:
        beta_map.setdefault(person.team, {})[person.name] = {"beta": person.beta}
        if person.is_leader:
            leaders.setdefault(person.team, []).append(person.beta)
    return beta_map, leaders


def classify_alert(pct: float) -> str:
    """四级预警分级：按 ALERT_RULES 从高到低命中。

    配置加载后只有 90/75/50 三档，大于 90（含大于 100）为满载，没有停工档。
    未套配置时 constants 里的预置表仍可能含「停工(超满载)」，那不是现网口径。
    """
    for threshold, level in C.ALERT_RULES:
        if pct > threshold:
            return level
    return C.ALERT_LIGHT
