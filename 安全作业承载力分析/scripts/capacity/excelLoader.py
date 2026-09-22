# -*- coding: utf-8 -*-
"""
Excel 数据源加载（excelLoader：.xls / .xlsm / .xlsx -> WorkDataset）
====================================================================

按扩展名选引擎：.xls/.xlsm 用 xlrd，.xlsx 用 openpyxl（均按内网镜像安装）。
主表 / 人员表解析统一按配置的列映射（config.DEFAULT_INPUT_CONFIG 或配置文件）。
"""
from __future__ import annotations

import os
from typing import Any, Dict, List, Optional, Tuple

from . import constants as C
from .constants import (DEFAULT_BETA, FORCE_BETA, InputDataError,
                        FileAccessError, FileContentError)
from .config import _cfg_get, DEFAULT_INPUT_CONFIG
from .parsing import _rows_to_plans, build_dataset
from .models import WorkDataset


def _sheet_nrows(sh: Any) -> int:
    """跨引擎行数：xlrd(nrows) / openpyxl(max_row)。"""
    return sh.nrows if hasattr(sh, "nrows") else sh.max_row


def _sheet_ncols(sh: Any) -> int:
    """跨引擎列数：xlrd(ncols) / openpyxl(max_column)。"""
    return sh.ncols if hasattr(sh, "ncols") else sh.max_column


def _sheet_cell(sh: Any, r: int, c: int) -> Any:
    """跨引擎取单元格值：xlrd(cell_value，0-based) / openpyxl(cell().value，1-based)。"""
    if hasattr(sh, "cell_value"):
        return sh.cell_value(r, c)
    return sh.cell(r + 1, c + 1).value


def _find_workbook_sheet(wb: Any, want: str, kind: str) -> Any:
    """跨引擎按名取工作表：xlrd(sheet_by_name) / openpyxl(wb[name])。"""
    names = wb.sheet_names() if hasattr(wb, "sheet_names") else list(wb.sheetnames)
    if want not in names:
        raise InputDataError(f"{kind}工作表「{want}」不存在，可选：{names}",
                             {"sheets": names, "want": want})
    return wb.sheet_by_name(want) if hasattr(wb, "sheet_by_name") else wb[want]


def _xls_rowdicts(sh: Any, header_row: int, data_start: int) -> List[Dict[str, Any]]:
    """把工作表转成 [ {列名: 单元格值} ... ]（日期单元格按各自引擎已转 datetime/date）。"""
    header2idx: Dict[str, int] = {}
    for c in range(_sheet_ncols(sh)):
        h = str(_sheet_cell(sh, header_row, c) or "").strip()
        if h:
            header2idx[h] = c
    out: List[Dict[str, Any]] = []
    for r in range(data_start, _sheet_nrows(sh)):
        row = {name: _sheet_cell(sh, r, idx) for name, idx in header2idx.items()}
        out.append(row)
    return out


def _build_xls_persons(wb: Any, input_cfg: Dict[str, Any]) -> List[Dict[str, Any]]:
    """按“人员档案”配置跨引擎解析人员表。姓名/班组/角色为必填，β/工区/职务可选。"""
    pa = _cfg_get(input_cfg, "人员档案", {}) or {}
    sheet = str(_cfg_get(pa, "工作表", "人员信息"))
    sh = _find_workbook_sheet(wb, sheet, "人员档案")
    hr = int(_cfg_get(pa, "表头行号", 0) or 0)
    dr = int(_cfg_get(pa, "数据起始行号", hr + 1) or (hr + 1))
    cmap = _cfg_get(pa, "人员列映射", {}) or {}
    header2idx = {str(_sheet_cell(sh, hr, c) or "").strip(): c for c in range(_sheet_ncols(sh))}
    # 必填列：姓名 / 所属班组 / 是否工作负责人；其余（β、单位工区、职务）为可选列
    required = [cmap[k] for k in ("姓名", "所属班组", "是否工作负责人") if k in cmap]
    missing = [cn for cn in required if str(cn) not in header2idx]
    if missing:
        raise InputDataError(f"人员表缺少列：{missing}（可在配置文件“人员档案.人员列映射”调整）",
                             {"missing": missing})
    keywords = list(_cfg_get(pa, "角色含工作负责人关键字", ["工作负责人"]))
    out: List[Dict[str, Any]] = []
    for r in range(dr, _sheet_nrows(sh)):
        name = _sheet_cell(sh, r, header2idx[str(cmap["姓名"])])
        if name in (None, ""):
            continue
        team = _sheet_cell(sh, r, header2idx[str(cmap["所属班组"])])
        if team in (None, ""):
            team = "未分配"
        role_mark = str(_sheet_cell(sh, r, header2idx[str(cmap["是否工作负责人"])]))
        beta_col = cmap.get("人员效能系数β")
        if beta_col and str(beta_col) in header2idx:
            bv = _sheet_cell(sh, r, header2idx[str(beta_col)])
            beta = DEFAULT_BETA if bv in (None, "") or str(bv).strip() == "无" else float(bv)
        else:
            beta = DEFAULT_BETA
        is_leader = any(k in role_mark for k in keywords)
        def _opt(column_key: str) -> str:
            col = cmap.get(column_key)
            if not col or str(col) not in header2idx:
                return ""
            v = _sheet_cell(sh, r, header2idx[str(col)])
            return str(v).strip() if v not in (None, "") else ""
        # 县公司 / 单位（工区）为可选列，供管理承载力归列；职务与角色合并进 role 供管理岗判定
        out.append({"name": str(name), "team": str(team), "beta": float(beta),
                    "is_leader": is_leader,
                    "county": _opt("县公司"),
                    "work_area": _opt("单位（工区）"),
                    "role": f"{role_mark} {_opt('职务')}".strip()})
    C.LOGGER.info("人员档案解析 %d 条（来自工作表「%s」）", len(out), sheet)
    return out


def _parse_workbook(wb: Any, input_cfg: Dict[str, Any], path: str,
                    source_label: str) -> WorkDataset:
    """按配置解析工作簿的主表与人员表 -> WorkDataset（xlrd/openpyxl 共用公共尾部）。"""
    main_sheet = str(_cfg_get(input_cfg, "主工作表", "秋检计划"))
    sh = _find_workbook_sheet(wb, main_sheet, "作业计划")
    hr = int(_cfg_get(input_cfg, "表头行号", 0) or 0)
    dr = int(_cfg_get(input_cfg, "数据起始行号", hr + 1) or (hr + 1))
    col_rows = _xls_rowdicts(sh, hr, dr)
    plans = _rows_to_plans(col_rows, input_cfg)
    pa = _cfg_get(input_cfg, "人员档案", {}) or {}
    manage_cfg = _cfg_get(input_cfg, "管理承载力", {}) or {}
    manage_enabled = bool(_cfg_get(manage_cfg, "是否启用", False))
    # 人员档案现为真实分母的**必要数据源**（班组可用人数 / 工作负责人真名均来自此），
    # 不再以 β/管理承载力开关为条件：Excel 一律尝试加载；表缺失/非法时优雅降级为空
    # （此时计划将因班组匹配不到而整条剔除，见 roster.filter_roster_matchable）。
    try:
        persons = _build_xls_persons(wb, input_cfg)
    except C.InputDataError as exc:
        C.LOGGER.warning("人员档案不可用（%s）：承载力分母将无法匹配，相关计划将被剔除",
                         exc.message)
        persons = []
    C.LOGGER.info("数据源 %s：%s（主表「%s」，%d 行；人员档案 %d 条）",
                  source_label, path, main_sheet, len(col_rows), len(persons))
    return build_dataset({"plans": plans, "persons": persons, "absence": []})


def _open_workbook(path: str) -> Tuple[Any, str]:
    """跨引擎打开工作簿并返回 (wb, 引擎名："xlsx"/"xls")；.xlsx 会话须调用方 close。"""
    ext = os.path.splitext(str(path).lower())[1]
    if ext == ".xlsx":
        try:
            import openpyxl  # 仅 .xlsx 输入时才依赖
        except ImportError as exc:
            raise FileContentError("处理 .xlsx 需要 openpyxl 库（内网镜像：pip install openpyxl）",
                                   path, str(exc)) from exc
        try:
            return openpyxl.load_workbook(path, read_only=True, data_only=True), "xlsx"
        except Exception as exc:
            raise FileContentError(f"无法打开 Excel 文件：{exc}", path, str(exc)) from exc
    try:
        import xlrd  # 仅 .xls/.xlsm 输入时才依赖，保持标准库优先
    except ImportError as exc:
        raise FileContentError("处理 .xls/.xlsm 需要 xlrd 库（内网镜像：pip install xlrd）",
                               path, str(exc)) from exc
    try:
        return xlrd.open_workbook(path), "xls"
    except Exception as exc:  # xlrd 对非法文件可能抛 XLRDError 等
        raise FileContentError(f"无法打开 Excel 文件：{exc}", path, str(exc)) from exc


def read_sheet_rows(path: str, sheet_name: str, header_row: int = 0,
                    data_start: Optional[int] = None,
                    input_cfg: Optional[Dict[str, Any]] = None,
                    required: bool = False) -> List[Dict[str, Any]]:
    """读取工作簿中指定工作表的原始行（列名 -> 单元格值）。

    供需要原始展示列（如“电压等级”）或辅助表（“组织架构”）的功能复用；
    缺表时 required=True 抛 InputDataError，否则返回空表（优雅降级）。
    """
    if data_start is None:
        data_start = header_row + 1
    wb, kind = _open_workbook(path)
    try:
        names = wb.sheet_names() if hasattr(wb, "sheet_names") else list(wb.sheetnames)
        if sheet_name not in names:
            if required:
                raise InputDataError(f"工作表「{sheet_name}」不存在，可选：{names}",
                                     {"sheets": names, "want": sheet_name})
            return []
        sh = wb.sheet_by_name(sheet_name) if hasattr(wb, "sheet_by_name") else wb[sheet_name]
        return _xls_rowdicts(sh, header_row, data_start)
    finally:
        if kind == "xlsx":
            try:
                wb.close()
            except Exception:
                pass


def read_xls_persons(path: str, input_cfg: Dict[str, Any]) -> List[Dict[str, Any]]:
    """读取 Excel「人员信息」sheet -> 人员原始记录 [{name,team,beta,is_leader,county,work_area,role}]。

    供需要独立构建班组档案索引的功能（如 genLetterParams 的负责人真名匹配）复用；
    表缺失/非法时抛 InputDataError，由调用方决定是否降级。
    """
    if not os.path.exists(path):
        raise FileAccessError(f"输入文件不存在：{path}", path)
    wb, kind = _open_workbook(path)
    try:
        return _build_xls_persons(wb, input_cfg)
    finally:
        if kind == "xlsx":
            try:
                wb.close()
            except Exception:
                pass


def read_xls_dataset(path: str, input_cfg: Optional[Dict[str, Any]] = None) -> WorkDataset:
    """读取 Excel 作业计划 -> WorkDataset。按扩展名选引擎：.xls/.xlsm 用 xlrd，.xlsx 用 openpyxl。"""
    if not os.path.exists(path):
        raise FileAccessError(f"输入文件不存在：{path}", path)
    if os.path.isdir(path):
        raise FileAccessError(f"输入路径是目录而非文件：{path}", path)
    input_cfg = input_cfg or DEFAULT_INPUT_CONFIG

    wb, kind = _open_workbook(path)
    try:
        return _parse_workbook(wb, input_cfg, path, f"Excel(.{kind})")
    finally:
        if kind == "xlsx":
            try:
                wb.close()
            except Exception:
                pass
