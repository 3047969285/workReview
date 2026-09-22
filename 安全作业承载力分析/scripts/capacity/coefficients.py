# -*- coding: utf-8 -*-
"""
系数解析与预警分级（coefficients：β 人员效能 / θ 作业风险 / 四级预警）
====================================================================

- `risk_theta()`：作业风险系数 θ（工作内容或作业类型取值为「装表接电」时 0.2；无风险计 0；未知按四级回退）；
- `build_person_beta()`：由人员档案构建 班组->姓名->β 索引与 班组->负责人β 列表；
- `classify_alert()`：四级预警分级（>100 判停工管控）。
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
    θ 取 RISK_COEF「装表接电」（配置为 0.2）。作业类型里只是含有「装表」不再改判五级。
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
    """四级预警分级：命中即返回（>100 判停工管控）。"""
    for threshold, level in C.ALERT_RULES:
        if pct > threshold:
            return level
    return C.ALERT_LIGHT
