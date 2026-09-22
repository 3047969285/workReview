# -*- coding: utf-8 -*-
"""
数据解析与校验（repository 层：原始 JSON/接口/表格行 -> 领域模型）
===============================================================

- `parse_person / parse_plan / parse_absence`：单条对象解析校验（R1 输入校验先行）；
- `build_dataset`：{persons, plans, absence} 规范化结构 -> WorkDataset；
- `read_dataset_file`：输入 JSON 文件读取 + 文件级 errorCode 映射；
- `_rows_to_plans / _rows_from_interface / read_interface_dataset`：接口 /
  与 Excel 同构的表格行（列名->值）解析。
"""
from __future__ import annotations

import json
import os
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

from . import constants as C
from .constants import (ABSENCE_FIELD_ALIASES, FIELD_ALIASES,
                        PERSON_FIELD_ALIASES, DEFAULT_BETA, FORCE_BETA,
                        STD_HOURS_PER_DAY, InputDataError, FileAccessError,
                        FileContentError)
from .config import _cfg_get, _PLAN_CANON, DEFAULT_INPUT_CONFIG
from .models import AbsenceRecord, PersonRecord, PlanRecord, WorkDataset


def _pick(raw: Dict[str, Any], aliases: Tuple[str, ...]) -> Any:
    """按别名优先级取值；未命中返回 None。"""
    for key in aliases:
        if key in raw and raw[key] is not None:
            return raw[key]
    return None


def parse_date(value: Any) -> Optional[date]:
    """日期解析：兼容 datetime / date / 常见字符串格式；解析失败返回 None。"""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S",
                "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def to_float(value: Any, field_desc: str) -> float:
    """将数值字段转换为 float；非法时抛 InputDataError（R1 明确失败）。"""
    try:
        return float(value)
    except (TypeError, ValueError) as exc:
        raise InputDataError(
            f"字段「{field_desc}」取值无法解析为数值：{value!r}", {"field": field_desc}
        ) from exc


def coerce_date(value: Any, mode: str) -> Any:
    """日期规整：datetime/date 原样；Excel 序列号按 1899-12-30 基准转；其余（字符串）交给 parse_date。"""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value is None or str(value).strip() == "":
        return None
    if mode == "excel序列号":
        try:
            serial = float(value)
            return (datetime(1899, 12, 30) + timedelta(days=serial)).date()
        except (TypeError, ValueError):
            pass
    return value


def clamp_hours(hours: float) -> float:
    """作业时长 T 规整：不足 1 小时按 1 小时，超过 8 小时按 8 小时。"""
    return min(max(hours, C.MIN_WORK_HOURS), C.MAX_WORK_HOURS)


def _require_date(value: Any, field_desc: str) -> date:
    result = parse_date(value)
    if result is None:
        raise InputDataError(f"字段「{field_desc}」无法解析为日期：{value!r}",
                             {"field": field_desc})
    return result


def _parse_bool(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    return str(value).strip().lower() in ("是", "true", "1", "yes")


def parse_person(raw: Any, idx: int) -> PersonRecord:
    """解析单条人员档案；失败抛 InputDataError 并指明出错行。"""
    if not isinstance(raw, dict):
        raise InputDataError(f"人员档案第 {idx} 条不是对象：{raw!r}", {"item": idx})
    name = _pick(raw, PERSON_FIELD_ALIASES["name"])
    if not name:
        raise InputDataError(f"人员档案第 {idx} 条缺少姓名（name）", {"item": idx})
    team = str(_pick(raw, PERSON_FIELD_ALIASES["team"]) or "未分配")
    # 人员效能系数 β：当前阶段先统一按 1.0 代替；
    # 接入真实人员效能数据源后，置 FORCE_BETA=False 即恢复从输入字段逐人读取
    if FORCE_BETA:
        beta = DEFAULT_BETA
    else:
        # β 仅在缺省（None）时取 1.0；值为 0 需原样保留，避免改变原口径
        beta_raw = _pick(raw, PERSON_FIELD_ALIASES["beta"])
        beta = to_float(DEFAULT_BETA if beta_raw is None else beta_raw,
                        f"人员{name}的beta")
    if beta <= 0:
        raise InputDataError(f"人员「{name}」的 β 需为正数，当前：{beta}",
                             {"name": name})
    return PersonRecord(name=str(name), team=team, beta=round(beta, 4),
                        is_leader=_parse_bool(_pick(raw, PERSON_FIELD_ALIASES["is_leader"])),
                        county=str(_pick(raw, PERSON_FIELD_ALIASES["county"]) or ""),
                        work_area=str(_pick(raw, PERSON_FIELD_ALIASES["work_area"]) or ""),
                        role=str(_pick(raw, PERSON_FIELD_ALIASES["role"]) or ""))


def parse_plan(raw: Any, idx: int) -> PlanRecord:
    """解析单条作业计划；校验必填字段并规整时长 / 起止日期。"""
    if not isinstance(raw, dict):
        raise InputDataError(f"作业计划第 {idx} 条不是对象：{raw!r}", {"item": idx})
    name = str(_pick(raw, FIELD_ALIASES["name"]) or f"作业计划{idx}")
    team = _pick(raw, FIELD_ALIASES["team"])
    if not team:
        raise InputDataError(f"作业计划第 {idx} 条缺少所属班组（team）", {"item": idx})
    start = _require_date(_pick(raw, FIELD_ALIASES["start"]), f"作业计划第{idx}条.start")
    end = _require_date(_pick(raw, FIELD_ALIASES["end"]), f"作业计划第{idx}条.end")
    if end < start:
        raise InputDataError(
            f"作业计划第 {idx} 条结束日期早于开始日期：{start} > {end}",
            {"item": idx, "start": str(start), "end": str(end)})

    hours_raw = _pick(raw, FIELD_ALIASES["hours"])
    if hours_raw in (None, ""):
        hours_raw = None  # 未填(含空串)按标准日工时回退
    hours = clamp_hours(to_float(hours_raw if hours_raw is not None else STD_HOURS_PER_DAY,
                                 f"作业计划第{idx}条.hours"))

    members_raw = _pick(raw, FIELD_ALIASES["members"])
    try:
        members = int(members_raw) if members_raw not in (None, "") else 0
    except (TypeError, ValueError) as exc:
        raise InputDataError(
            f"作业计划第 {idx} 条所需自有人员人数非法：{members_raw!r}", {"item": idx}
        ) from exc
    if members < 0:
        raise InputDataError(
            f"作业计划第 {idx} 条所需自有人员人数不能为负：{members_raw!r}", {"item": idx})

    leader = _pick(raw, FIELD_ALIASES["leader"])
    leader_str = str(leader).strip() if leader not in (None, "") else None
    # 工作负责人列 = "需要的工作负责人数量"（业务口径：填数量，非姓名）。
    # 数字 -> 该数量；非数字（历史数据可能填姓名）-> 视为 1 名负责人。空按 1。
    leader_count = 1
    if leader_str:
        try:
            n = int(round(float(leader_str)))
            leader_count = n if n >= 1 else 1
        except (TypeError, ValueError):
            leader_count = 1
    risk = _pick(raw, FIELD_ALIASES["risk"])
    plan_type = str(_pick(raw, FIELD_ALIASES["type"]) or "")
    company_raw = _pick(raw, FIELD_ALIASES["company"])
    work_area_raw = _pick(raw, FIELD_ALIASES["work_area"])
    return PlanRecord(name=name, team=str(team), leader=leader_str,
                      leader_count=leader_count, members=members, hours=hours,
                      risk=str(risk) if risk not in (None, "") else None,
                      plan_type=plan_type, start=start, end=end,
                      company=str(company_raw) if company_raw not in (None, "") else None,
                      work_area=str(work_area_raw) if work_area_raw not in (None, "") else None)


def parse_absence(raw: Any, idx: int) -> AbsenceRecord:
    """解析单条不在岗（请假 / 出差）记录。"""
    if not isinstance(raw, dict):
        raise InputDataError(f"不在岗记录第 {idx} 条不是对象：{raw!r}", {"item": idx})
    name = _pick(raw, ABSENCE_FIELD_ALIASES["name"])
    if not name:
        raise InputDataError(f"不在岗记录第 {idx} 条缺少姓名（name）", {"item": idx})
    start = _require_date(_pick(raw, ABSENCE_FIELD_ALIASES["start"]), f"不在岗第{idx}条.start")
    end = _require_date(_pick(raw, ABSENCE_FIELD_ALIASES["end"]), f"不在岗第{idx}条.end")
    if end < start:
        raise InputDataError(
            f"不在岗记录第 {idx} 条结束日期早于开始日期：{start} > {end}",
            {"item": idx, "start": str(start), "end": str(end)})
    return AbsenceRecord(name=str(name), start=start, end=end)


def build_dataset(raw: Dict[str, Any]) -> WorkDataset:
    """将原始数据结构（persons / plans / absence）解析校验为 WorkDataset。"""
    persons_raw = raw.get("persons", []) or []
    plans_raw = raw.get("plans", []) or []
    absences_raw = raw.get("absence", []) or []
    if not isinstance(persons_raw, list) or not isinstance(plans_raw, list) \
            or not isinstance(absences_raw, list):
        raise InputDataError("输入结构非法：persons / plans / absence 均须为数组")
    persons = tuple(parse_person(p, i + 1) for i, p in enumerate(persons_raw))
    plans = tuple(parse_plan(p, i + 1) for i, p in enumerate(plans_raw))
    absences = tuple(parse_absence(a, i + 1) for i, a in enumerate(absences_raw))
    return WorkDataset(persons=persons, plans=plans, absences=absences)


def read_dataset_file(path: str) -> WorkDataset:
    """读取输入 JSON 文件并解析；文件级错误映射为统一 errorCode。"""
    if not os.path.exists(path):
        raise FileAccessError(f"输入文件不存在：{path}", path)
    if os.path.isdir(path):
        raise FileAccessError(f"输入路径是目录而非文件：{path}", path)
    try:
        with open(path, "r", encoding="utf-8") as fh:
            raw = json.load(fh)
    except UnicodeDecodeError as exc:
        raise FileContentError("输入文件编码无法识别（需 UTF-8）", path, str(exc)) from exc
    except json.JSONDecodeError as exc:
        raise FileContentError("输入文件不是合法的 JSON", path, str(exc)) from exc
    except OSError as exc:
        raise FileAccessError(f"输入文件读取失败：{exc}", path) from exc
    if not isinstance(raw, dict):
        raise FileContentError("输入 JSON 顶层必须是对象（含 persons / plans / absence）",
                               path)
    return build_dataset(raw)


def _rows_to_plans(col_rows: List[Dict[str, Any]], input_cfg: Dict[str, Any]) -> List[Dict[str, Any]]:
    """按列映射把表格行(列名->值)解析为作业计划对象(领域字段键)；空排期/无班组行跳过。"""
    colmap = _cfg_get(input_cfg, "作业计划列映射", {}) or {}
    date_mode = str(_cfg_get(input_cfg, "时间列格式", "excel序列号"))
    plans: List[Dict[str, Any]] = []
    skipped = 0
    for row in col_rows:
        raw: Dict[str, Any] = {}
        for field, colname in colmap.items():
            value = row.get(colname)
            if field in ("计划开始时间", "计划结束时间"):
                value = coerce_date(value, date_mode)
            if isinstance(value, float) and value == int(value):
                value = int(value)
            raw[_PLAN_CANON.get(field, field)] = value
        if raw.get("start") in (None, "") and raw.get("end") in (None, ""):
            skipped += 1
            continue
        team = raw.get("team") or raw.get("workArea")
        if not team:
            skipped += 1
            continue
        raw["team"] = str(team)
        plans.append(raw)
    C.LOGGER.info("已解析作业计划 %d 条，跳过无排期/无班组行 %d 条", len(plans), skipped)
    return plans


def _rows_from_interface(obj: Any) -> Optional[List[Dict[str, Any]]]:
    """提取接口 JSON 中的表格行(列名->值)；无法识别返回 None。"""
    if isinstance(obj, list) and obj and isinstance(obj[0], dict):
        return list(obj)
    if isinstance(obj, dict) and isinstance(obj.get("rows"), list):
        rows = obj["rows"]
        if rows and isinstance(rows[0], dict):
            return list(rows)
        if isinstance(obj.get("header"), list):
            hdr = [str(h) for h in obj["header"]]
            return [
                {hdr[i]: (row[i] if i < len(row) else None) for i in range(len(hdr))}
                for row in rows
            ]
    return None


def read_interface_dataset(obj: Any, input_cfg: Optional[Dict[str, Any]] = None) -> WorkDataset:
    """接口返回 JSON：规范化结构(persons/plans/absence) 或 与 Excel 同构的表格行(含表头列名)。"""
    input_cfg = input_cfg or DEFAULT_INPUT_CONFIG
    if isinstance(obj, dict) and any(k in obj for k in ("persons", "plans", "absence")):
        return build_dataset(obj)
    col_rows = _rows_from_interface(obj)
    if col_rows is None:
        raise InputDataError(
            "接口数据无法识别：应为 {persons, plans, absence} 规范化结构，"
            "或 {rows: [{“列名”: 值}, ...]}/[{“列名”: 值}, ...] 表格行结构")
    plans = _rows_to_plans(col_rows, input_cfg)
    C.LOGGER.info("数据源接口：表格行 %d 条（按配置列映射解析）", len(col_rows))
    return build_dataset({"plans": plans, "persons": [], "absence": []})
