# -*- coding: utf-8 -*-
"""
领域模型（domain：校验后的不可变数据结构）
====================================

承载力量化计算统一流转的数据载体，解析层产出、计算层消费。
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Optional, Tuple


@dataclass(frozen=True)
class PersonRecord:
    """人员档案（真实姓名）。"""

    name: str
    team: str
    beta: float
    is_leader: bool = False
    # 人员所属 市/县 公司（county）与 单位（工区）/ 职务（角色）：供管理承载力统计
    # “在家管理人员数”并归一化到报告单位列；未提供时为空串（计算层回退单位/默认人数）
    county: str = ""
    work_area: str = ""
    role: str = ""


@dataclass(frozen=True)
class PlanRecord:
    """作业计划（已拆分并规范化起止日期与时长）。"""

    name: str
    team: str
    leader: Optional[str]
    # 计划「工作负责人」列填的是"需要的工作负责人数量"（非姓名）：解析为整数数量，
    # 空/非数字按 1（每张工作票至少 1 名工作负责人）。负责人承载力分子按此数量计需求。
    leader_count: int
    members: int
    hours: float
    risk: Optional[str]
    plan_type: str
    start: date
    end: date
    # 业主单位公司 / 工区：供组织架构归一化到报告矩阵的单位列（未启用时恒为 None）
    company: Optional[str] = None
    work_area: Optional[str] = None


@dataclass(frozen=True)
class AbsenceRecord:
    """请假 / 出差 / 参会 / 培训（不在岗）记录。"""

    name: str
    start: date
    end: date


@dataclass(frozen=True)
class WorkDataset:
    """规范化后的完整输入数据集，是上层计算的唯一数据来源。"""

    persons: Tuple[PersonRecord, ...]
    plans: Tuple[PlanRecord, ...]
    absences: Tuple[AbsenceRecord, ...]
