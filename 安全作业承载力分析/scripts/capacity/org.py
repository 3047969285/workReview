# -*- coding: utf-8 -*-
"""
组织架构归一化（org：班组/工区/公司 -> 报告矩阵单位列）
========================================================

`UnitNormalizer.unit_of()`：把计划的 公司 / 工区 / 班组 收敛到报告矩阵表的
单位列（中心/公司，REPORT_UNITS 17 列）。能否启用由配置「输入解析.组织架构映射.
是否启用」控制（默认 false，不改变既有输出）；全部未命中归入「其他」。

数据驱动归列（可选）：`load_sheet()` 载入源 xls「组织架构」sheet 行后，先用
「班组名 + 计划公司/工区上下文」做精确消歧：同名班组在 sheet 多个县公司子树里
重复出现（如 变电运维一班 在 即墨/胶州/市公司直属 各有 1 行），靠计划公司的
县公司别名、再靠工区别名逐级收敛到唯一行；命中的行「县公司简称」∈单位清单直接归
该县公司，否则对行内「单位（工区）」等名再走关键词规则。二者解耦：sheet 目录只负责
「班组/工区名 + 上下文 → 单位（工区）」，「单位（工区）→ 归列单位」仍由 `_match` 完成。
上下文消歧不出的（如外部承接单位、公司不在 sheet）回退 `_match` 既有单值逻辑
（单位清单精确 / 关键词规则 / 最长包含），两路共用同一套单位清单，保证口径一致。
"""
from __future__ import annotations

import hashlib
from collections import defaultdict
from dataclasses import replace
from typing import Any, Dict, List, Optional, Tuple

from . import constants as C
from .config import _cfg_get

# 同一进程里日/周/月会各自 new 一个 UnitNormalizer，但归列只取决于单位清单、
# 规则和同一份组织架构行。按这个指纹记住 unit_of 结果，避免十几万次重复打分。
_UNIT_OF_MEMO: Dict[tuple, str] = {}
_COMP_KEY_MEMO: Dict[str, str] = {}
_ROW_PREP: Dict[int, Tuple[str, Tuple[str, str]]] = {}
_MATCH_MISS = object()

# 「组织架构」sheet 中参与目录构建的列（班组简/全名 → 工区简/全名）
_SHEET_INDEX_KEYS = ("班组", "班组简称", "单位（工区）", "工区简称")
# 命中 sheet 行后再按该顺序对行内名称做关键词归列
_SHEET_UNIT_KEYS = ("单位（工区）", "工区简称", "班组", "班组简称")
# 公司名归一：去「国网 / 供电公司」前后缀后取中间的区/县/市名（如 黄岛区、即墨区）
_COMPANY_STRIP = ("国网", "供电公司")


def _comp_key(value: str) -> str:
    """公司名 -> 可比对的短键：国网平度市供电公司 → 平度市。"""
    hit = _COMP_KEY_MEMO.get(value)
    if hit is not None:
        return hit
    s = value.strip()
    for p in _COMPANY_STRIP:
        s = s.replace(p, "")
    s = s.strip()
    _COMP_KEY_MEMO[value] = s
    return s


def _row_score_parts(row: Dict[str, Any]) -> Tuple[str, Tuple[str, str]]:
    """一行组织架构的打分材料只算一次：县公司短键、工区名二元组。"""
    rid = id(row)
    hit = _ROW_PREP.get(rid)
    if hit is not None:
        return hit
    prepared = (
        _comp_key(str(row.get("县公司") or "")),
        (str(row.get("单位（工区）") or ""), str(row.get("工区简称") or "")),
    )
    _ROW_PREP[rid] = prepared
    return prepared


def _text_score(needle: str, haystacks: List[str]) -> int:
    """needle 与最贴近 haystack 的包含长度分（一方含另一方，取短者长度）。"""
    needle = needle.strip()
    if not needle:
        return 0
    best = 0
    for h in haystacks:
        h = h.strip()
        if not h:
            continue
        if needle in h:
            best = max(best, len(needle))
        elif h in needle:
            best = max(best, len(h))
    return best


# 班组归属规则（内置兜底，由 REPORT_UNITS 归纳）：config 瘦身（v1.6）后不再从
# 「组织架构映射.班组归属规则」读取静态清单，这里作为 源表「组织架构」sheet 未命中 /
# 未载入时的稳定回退；sheet 命中时按 县公司简称 / 工区简称 优先数据驱动归列。
# 顺序即优先级：县公司关键词在前，避免被「城阳中心」等中心词干扰。
_BUILTIN_RULES: List[Tuple[Tuple[str, ...], str]] = [
    (("国网黄岛", "黄岛"), "黄岛公司"),
    (("国网胶州", "胶州"), "胶州公司"),
    (("国网即墨", "即墨"), "即墨公司"),
    (("国网平度", "平度"), "平度公司"),
    (("国网莱西", "莱西"), "莱西公司"),
    (("变电运维",), "变电运维"),
    (("变电检修",), "变电检修"),
    (("输配电", "输电"), "输电中心"),
    (("带电",), "带电中心"),
    (("电缆",), "电缆中心"),
    (("市南",), "市南中心"),
    (("市北",), "市北中心"),
    (("崂山",), "崂山中心"),
    (("李沧",), "李沧中心"),
    (("城阳",), "城阳中心"),
    (("安装",), "安装公司"),
    (("送变电",), "送变电"),
]

# 公司名去前缀回退（无 sheet 简称时压缩长公司名用）
_SHORT_STRIP = ("国网", "青岛", "供电公司", "电力工程安装有限公司", "有限公司", "分公司")


class UnitNormalizer:
    """组织架构归一化：命中优先级 = 源表(sheet)行按 公司/工区 上下文化解 -> 单位清单精确匹配 -> 归属规则(包含词,按顺序) -> 单位清单包含匹配(取最长)。"""

    def __init__(self, org_cfg: Optional[Dict[str, Any]] = None) -> None:
        org_cfg = org_cfg or {}
        self.enabled = bool(_cfg_get(org_cfg, "是否启用", False))
        units = _cfg_get(org_cfg, "单位清单", None)
        self.units: Tuple[str, ...] = tuple(str(u) for u in units) if units else C.REPORT_UNITS
        # 归属规则：config 瘦身后不再内置静态清单，仅当旧配置仍显式携带时才读取，
        # 否则用代码内置 _BUILTIN_RULES（见本模板底部定义）
        self.rules: List[Tuple[Tuple[str, ...], str]] = []
        rules = _cfg_get(org_cfg, "班组归属规则", None)
        if isinstance(rules, list):
            for rule in rules:
                if not isinstance(rule, dict):
                    continue
                inc = _cfg_get(rule, "包含", None)
                owner = _cfg_get(rule, "归属", None)
                if inc and owner:
                    self.rules.append((tuple(str(x) for x in inc), str(owner)))
        if not self.rules:
            self.rules = list(_BUILTIN_RULES)
        self._rules_token = tuple((inc, owner) for inc, owner in self.rules)
        # 「组织架构」sheet 目录：名 -> 全部同名行（同名班组多子树并存，须上下文消歧）
        self._by_name: Dict[str, List[Dict[str, Any]]] = {}
        self._has_sheet = False
        self._match_cache: Dict[str, Any] = {}
        self._memo_token: tuple = ()
        self._bind_memo()

    def has_sheet(self) -> bool:
        """是否已载入「组织架构」sheet 行（上下文消歧可用）。"""
        return self._has_sheet

    def load_sheet(self, rows: Optional[List[Dict[str, Any]]] = None) -> "UnitNormalizer":
        """载入「组织架构」sheet 行（市公司/县公司/县公司简称/单位（工区）/工区简称/班组/班组简称）。

        建立「班组/工区名 -> 全部同名行」目录；同名行不做首见折叠，归列阶段用
        计划公司/工区上下文消歧（见 _resolve_context）。行缺失/非 dict 优雅跳过。
        """
        index: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
        for row in rows or []:
            if not isinstance(row, dict):
                continue
            for key in _SHEET_INDEX_KEYS:
                v = str(row.get(key) or "").strip()
                if v:
                    index[v].append(row)
        self._by_name = dict(index)
        self._has_sheet = True
        self._match_cache.clear()
        self._bind_memo()
        return self

    def _bind_memo(self) -> None:
        """归列指纹：单位清单 + 规则 + 组织架构行内容。日/周/月共用同一份记忆。"""
        parts: List[str] = []
        for name in sorted(self._by_name):
            parts.append(name)
            for row in self._by_name[name]:
                parts.append(str(id(row)))
                parts.append(str(row.get("县公司") or ""))
                parts.append(str(row.get("县公司简称") or ""))
                parts.append(str(row.get("单位（工区）") or ""))
                parts.append(str(row.get("工区简称") or ""))
                parts.append(str(row.get("班组") or ""))
                parts.append(str(row.get("班组简称") or ""))
        sheet_fp = hashlib.sha1("\0".join(parts).encode("utf-8")).hexdigest() if parts else ""
        self._memo_token = (self.units, self._rules_token, sheet_fp)

    def unit_of(self, values: Tuple[Any, ...]) -> str:
        """按 公司 -> 工区 -> 班组 归一化到报告单位列。

        sheet 已载入且计划班组名命中目录时，用「公司/工区上下文」消歧同名行后精确归列；
        否则回退 `_match` 既有单值逻辑（单位清单精确 / 关键词规则 / 最长包含）。
        相同三元组在本进程只算一次，日/周/月和矩阵过滤共用结果。
        """
        key = (self._memo_token, tuple(
            "" if v in (None, "") else str(v).strip() for v in values))
        cached = _UNIT_OF_MEMO.get(key)
        if cached is not None:
            return cached
        company, work_area, team = self._slice(values)
        if self._has_sheet and team:
            category = self._resolve_context(company, work_area, team)
            if category is not None:
                _UNIT_OF_MEMO[key] = category
                return category
        fallback = C.ORG_CATEGORY_OTHER
        for value in values:
            category = self._match(str(value).strip()) if value not in (None, "") else None
            if category is not None:
                fallback = category
        _UNIT_OF_MEMO[key] = fallback
        return fallback

    @staticmethod
    def _slice(values: Tuple[Any, ...]) -> Tuple[str, str, str]:
        """把 2/3 元组 values 切为 (公司, 工区, 班组) 字符串（缺省为空串）。"""
        v = tuple(values)
        company = str(v[0] or "").strip() if len(v) > 0 and v[0] not in (None, "") else ""
        work_area = str(v[1] or "").strip() if len(v) > 1 and v[1] not in (None, "") else ""
        team = str(v[2] or "").strip() if len(v) > 2 and v[2] not in (None, "") else ""
        return company, work_area, team

    def _resolve_context(self, company: str, work_area: str, team: str) -> Optional[str]:
        """按 班组名 → 公司消歧 → 工区消歧 逐级收敛到唯一 sheet 行并归列。

        优先用班组名（_by_name 目录）；班组不在目录时退化用工区名再试一遍（同一套
        公司消歧）。任一阶段收敛出唯一行即返回归列单位；收敛不出返回 None 交回 _match。
        """
        row = self._find_row(company, work_area, team)
        return self._unit_from_row(row) if row is not None else None

    def _find_row(self, company: str, work_area: str,
                  team: str) -> Optional[Dict[str, Any]]:
        """按 班组名 → 公司消歧 → 工区消歧 逐级收敛到唯一 sheet 行；无则 None。"""
        for name, wa in ((team, work_area), (work_area, "")):
            cands = self._by_name.get(name or "")
            if not cands:
                continue
            row = self._pick_candidate(cands, company, wa)
            if row is not None:
                return row
        return None

    def short_name(self, values: Tuple[Any, ...]) -> Tuple[str, str, str]:
        """取 公司 / 工区 / 班组 简称（sheet「简称」列优先，缺省回退源名/去前缀）。

        - 公司简称：命中 sheet 行取「县公司简称」（市公司直属行无县公司简称 ->
          用「单位（工区）」全名压缩）；未命中回退源公司名去「国网/供电公司」前缀。
        - 工区简称：sheet「工区简称」，缺省取「单位（工区）」全名。
        - 班组简称：sheet「班组简称」，缺省取「班组」全名。
        供报告 / 函件展示统一用简称（需求：公司名/工区名/班组名都用简称）。
        """
        company, work_area, team = self._slice(values)
        row = self._find_row(company, work_area, team) if self._has_sheet else None
        if row is None:
            return (self._short_company(company),
                    work_area or "-", team or "-")
        county = str(row.get("县公司简称") or "").strip()
        c = county if county and county != "市公司" else self._short_company(company)
        wa = str(row.get("工区简称") or "").strip() or \
            str(row.get("单位（工区）") or "").strip() or work_area or "-"
        tm = str(row.get("班组简称") or "").strip() or \
            str(row.get("班组") or "").strip() or team or "-"
        return c, wa, tm

    @staticmethod
    def _short_company(value: str) -> str:
        """公司名压缩：去「国网/青岛/供电公司…」前缀后取主干（如 黄岛、送变电）。"""
        s = str(value or "").strip()
        for p in _SHORT_STRIP:
            s = s.replace(p, "")
        s = s.strip("省市区县公司 ")
        return s or (str(value or "").strip() or "-")

    def _pick_candidate(self, cands: List[Dict[str, Any]],
                        company: str, work_area: str) -> Optional[Dict[str, Any]]:
        """在候选同名行里按 公司别名 + 工区别名 打分，唯一且分>0 才返回。

        即便同名行在 sheet 里唯一出现，也必须与计划公司子树重叠才采纳：同名外部
        施工队在多个县公司下分属不同业主（如 配电工程四队 sheet 仅记在平度，但
        直属计划也用该名），脱离公司上下文会把直属计划错归到县公司。
        """
        best: List[Dict[str, Any]] = []
        best_score = 0
        company_key = _comp_key(company) if company else ""
        for row in cands:
            score = 0
            county_key, areas = _row_score_parts(row)
            if company:
                score += _text_score(company_key, [county_key])
            if work_area:
                score += _text_score(work_area, list(areas))
            if score > best_score:
                best, best_score = [row], score
            elif score and score == best_score:
                best.append(row)
        if not best or best_score <= 0 or len(best) != 1:
            return None
        return best[0]

    def _unit_from_row(self, row: Dict[str, Any]) -> Optional[str]:
        """sheet 命中行 -> 最终归列单位：
        县公司简称∈单位清单逐级直归县公司；否则对行内名称走关键词/包含匹配。
        """
        county_abbr = str(row.get("县公司简称") or "").strip()
        if county_abbr and county_abbr in self.units:
            return county_abbr
        for key in _SHEET_UNIT_KEYS:
            name = str(row.get(key) or "").strip()
            if name:
                category = self._match(name)
                if category is not None:
                    return category
        return None

    def _match(self, value: str) -> Optional[str]:
        if not value:
            return None
        cached = self._match_cache.get(value, _MATCH_MISS)
        if cached is not _MATCH_MISS:
            return cached
        found: Optional[str] = None
        if value in self.units:
            found = value
        else:
            for inc, owner in self.rules:
                if any(k in value for k in inc):
                    found = owner
                    break
            if found is None:
                hits = [u for u in self.units if u in value or value in u]
                if hits:
                    found = max(hits, key=len)  # 取最长命中，避免“变电检修”误吞“变电检修一班”
        self._match_cache[value] = found
        return found


def load_org_sheet_rows(source_path: Optional[str],
                        input_cfg: Dict[str, Any]) -> Optional[List[Dict[str, Any]]]:
    """配置启用「组织架构映射.优先读取源表组织架构」且源为 Excel 时，返回「组织架构」sheet 行。

    sheet 缺失或非 Excel 源时优雅返回 None（保持纯关键词归列，不阻断流水线）。
    """
    if not source_path or not str(source_path).lower().endswith((".xls", ".xlsm", ".xlsx")):
        return None
    org = _cfg_get(input_cfg, "组织架构映射", {}) or {}
    if not (org.get("是否启用") and _cfg_get(org, "优先读取源表组织架构", False)):
        return None
    from .excelLoader import read_sheet_rows  # 局部导入避免顶层循环依赖
    sheet = str(_cfg_get(org, "组织架构工作表", "组织架构"))
    hr = int(_cfg_get(org, "表头行号", 0) or 0)
    dr = int(_cfg_get(org, "数据起始行号", hr + 1) or (hr + 1))
    return read_sheet_rows(source_path, sheet, header_row=hr, data_start=dr,
                           required=False)


def filter_matchable(dataset, normalizer: "UnitNormalizer"):
    """剔除无法归入组织架构标准单位列的计划（`unit_of` 归「其他」的项），返回新 dataset。

    用户口径：匹配不到组织架构的计划不统计、不进入矩阵/报告/函件。调用方须在
    `normalizer.enabled` 为真时调用；persons / absence 原样保留（β 索引按班组
    聚合，与单位归列无关）。无剔除时直接返回原对象（省一次 dataclasses.replace）。
    """
    kept = tuple(
        p for p in dataset.plans
        if normalizer.unit_of((p.company, p.work_area, p.team)) != C.ORG_CATEGORY_OTHER
    )
    return dataset if len(kept) == len(dataset.plans) else replace(dataset, plans=kept)
