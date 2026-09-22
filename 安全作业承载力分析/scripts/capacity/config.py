# -*- coding: utf-8 -*-
"""
配置加载（config：公式参数与输入解析均可配置，改配置不改代码）
============================================================

- `load_config()`：读技能目录 `config/capacity_config.json`（可用 `--config` 指定）；
- `apply_formula_config()`：把「公式与计算参数」应用到 `constants.py` 模块级系数
  （全局可变系数集中改一处），其它模块一律从 `C.*` 读取；
- `DEFAULT_INPUT_CONFIG`：内置缺省的输入解析配置，配置文件缺失时按此解析.
"""
from __future__ import annotations

import json
import os
from datetime import date
from typing import Any, Dict, Optional

from . import constants as C
from .constants import FileAccessError, FileContentError

CONFIG_FILE_NAME = "capacity_config.json"
CONFIG_SECTION_FORMULA = "公式与计算参数"
CONFIG_SECTION_INPUT = "输入解析"

# 内置缺省的输入解析配置（与技能目录 config/capacity_config.json 的“输入解析”节保持一致）
DEFAULT_INPUT_CONFIG: Dict[str, Any] = {
    "主工作表": "秋检计划",
    "表头行号": 0,
    "数据起始行号": 2,
    "时间列格式": "excel序列号",
    "作业计划列映射": {
        "计划名称": "工程名称",
        "作业班组": "现场施工单位班组",
        "计划开始时间": "计划开始时间",
        "计划结束时间": "计划结束时间",
        "作业时长(小时)": "工作预计时长（小时,仅为单日工作填写）",
        "作业风险等级": "作业风险等级",
        "工作负责人": "工作负责人",
        "所需自有人员人数": "所需自有人员人数（不含工作负责人）",
        "作业类型": "作业类型",
        "业主单位公司": "业主单位（公司）",
        "业主单位工区": "业主单位（工区）",
        "工程专业": "工程专业",
        "工作内容": "工作内容",
        "工作地点": "工作地点（精确到区/县）",
    },
    "人员档案": {
        "是否加载": False,
        "工作表": "人员信息",
        "表头行号": 0,
        "数据起始行号": 1,
        "人员列映射": {
            "姓名": "姓名",
            "所属班组": "班组",
            "是否工作负责人": "角色",
            "人员效能系数β": "作业素养能力（无需填写）",
            # 以下为可选列：管理承载力统计“在家管理人员数”用（县公司兜底、工区优先归列）
            "县公司": "县公司",
            "单位（工区）": "单位（工区）",
            "职务": "职务",
            "班组类型": "班组类型",
        },
        "角色含工作负责人关键字": ["工作负责人"],
    },
    "管理承载力": {
        "是否启用": True,
        "说明": "是否启用=true 时计算并输出管理承载力（F_管理=实需人天÷可供人天，可供=在家管理人员数×每人可到现场数）。",
    },
    "组织架构映射": {
        "是否启用": False,
        "工作表": "组织架构",
        "优先读取源表组织架构": False,
        "组织架构工作表": "组织架构",
        "表头行号": 0,
        "数据起始行号": 1,
        "说明": "是否启用=true 时把计划的公司、工区、班组归一化到报告矩阵的单位列，结果 JSON 增加 matrix(日期×单位)。单位清单不写在配置里。优先读取源表「组织架构」：县公司简称属于报告单位列时归该县公司，市公司直属单位按行内「单位（工区）」或「工区简称」归列。sheet 未命中时用代码内置关键词。",
    },
}

# 作业计划逻辑字段 -> 领域模型字段(与 FIELD_ALIASES 主键一致)
_PLAN_CANON: Dict[str, str] = {
    "计划名称": "name", "作业班组": "team", "计划开始时间": "start", "计划结束时间": "end",
    "作业时长(小时)": "hours", "作业风险等级": "risk", "工作负责人": "leader",
    "所需自有人员人数": "members", "作业类型": "type", "业主单位公司": "company",
    "业主单位工区": "workArea", "工程专业": "specialty", "工作内容": "content",
    "工作地点": "location",
}


def _cfg_get(node: Dict[str, Any], key: str, default: Any = None) -> Any:
    """多级安全取值。"""
    if not isinstance(node, dict):
        return default
    return node.get(key, default)


def default_config_path() -> Optional[str]:
    """技能目录 config 下的缺省配置文件；向上查找（scripts/capacity 可能嵌套于根目录），不存在返回 None。"""
    cur = os.path.dirname(os.path.abspath(__file__))
    for _ in range(8):
        cand = os.path.join(cur, "config", CONFIG_FILE_NAME)
        if os.path.isfile(cand):
            return cand
        parent = os.path.dirname(cur)
        if parent == cur:
            break
        cur = parent
    return None


def load_config(path: Optional[str] = None) -> Dict[str, Any]:
    """加载配置 JSON；显式指定但缺失/非法 -> 明确异常；未指定且无缺省文件 -> 空配置沿用内置默认。"""
    cfg_path = path or default_config_path()
    if cfg_path is None:
        return {}
    if not os.path.isfile(cfg_path):
        raise FileAccessError(f"配置文件不存在：{cfg_path}", cfg_path)
    try:
        with open(cfg_path, "r", encoding="utf-8") as fh:
            cfg = json.load(fh)
    except UnicodeDecodeError as exc:
        raise FileContentError("配置文件编码无法识别（需 UTF-8）", cfg_path, str(exc)) from exc
    except json.JSONDecodeError as exc:
        raise FileContentError("配置文件不是合法的 JSON", cfg_path, str(exc)) from exc
    except OSError as exc:
        raise FileAccessError(f"配置文件读取失败：{exc}", cfg_path) from exc
    if not isinstance(cfg, dict):
        raise FileContentError("配置文件顶层必须是对象（含“公式与计算参数”/“输入解析”节）", cfg_path)
    return cfg


def apply_formula_config(cfg: Dict[str, Any]) -> None:
    """把配置文件的“公式与计算参数”应用到 constants 模块级系数表；未配置项沿用内置默认。

    通过 `C.x = ...` 修改 constants 模块属性（拆分后统一从 `C.*` 读取），不再用 global。
    """
    f = _cfg_get(cfg, CONFIG_SECTION_FORMULA, {}) or {}
    C.STD_HOURS_PER_DAY = float(_cfg_get(f, "标准日工时(小时)", C.STD_HOURS_PER_DAY))
    bounds = _cfg_get(f, "工时边界(小时)", {}) or {}
    C.MIN_WORK_HOURS = float(_cfg_get(bounds, "最小", C.MIN_WORK_HOURS))
    C.MAX_WORK_HOURS = float(_cfg_get(bounds, "最大", C.MAX_WORK_HOURS))
    C.DEFAULT_BETA = float(_cfg_get(f, "默认人员效能系数β", C.DEFAULT_BETA))
    C.FORCE_BETA = bool(_cfg_get(f, "强制统一β", C.FORCE_BETA))
    C.LEADER_LOAD_CAP = float(_cfg_get(f, "负责人承载力上限", C.LEADER_LOAD_CAP))
    C.DENOM_EPS = float(_cfg_get(f, "分母保护值", C.DENOM_EPS))
    coef = _cfg_get(f, "作业风险系数θ", None)
    if isinstance(coef, dict) and coef:
        C.RISK_COEF = {str(k): float(v) for k, v in coef.items()}
    C.RISK_COEF_DEFAULT = float(_cfg_get(f, "未知风险等级回退θ", C.RISK_COEF_DEFAULT))
    C.RISK_COEF_NONE = float(_cfg_get(f, "无风险等级θ", C.RISK_COEF_NONE))
    C.TEAM_KIND_DEFAULT = str(_cfg_get(f, "默认班组类型", C.TEAM_KIND_DEFAULT))
    patrol = _cfg_get(f, "巡视作业类型", None)
    if isinstance(patrol, list):
        C.PATROL_TYPES = tuple(str(x) for x in patrol)
    opera = _cfg_get(f, "操作作业类型", None)
    if isinstance(opera, list):
        C.OPERATION_TYPES = tuple(str(x) for x in opera)
    rules = _cfg_get(f, "预警分级", None)
    if isinstance(rules, list) and rules:
        C.ALERT_RULES = tuple((float(r["阈值"]), str(r["级别"])) for r in rules)
    C.ALERT_LIGHT = str(_cfg_get(f, "轻载文案", C.ALERT_LIGHT))
    C.MASK_NAMES = bool(_cfg_get(f, "输出姓名脱敏", C.MASK_NAMES))
    # 工作日日历（剔除周末 + 法定节假日）：周末按开关剔除、节假日走可维护清单。
    # v2.3：is_workday 仅用于周期（周/月/年）累加分母与 management 区间聚合分母；
    # 日维度（days / matrix / management.days）直接按计划覆盖的自然日展开，不受此开关影响。
    nw = _cfg_get(f, "非工作日", {}) or {}
    C.EXCLUDE_WEEKENDS = bool(_cfg_get(nw, "是否剔除周末", C.EXCLUDE_WEEKENDS))
    hols = _cfg_get(nw, "法定节假日", None)
    if isinstance(hols, list):
        hs = set()
        for x in hols:
            try:
                hs.add(date.fromisoformat(str(x).strip()))
            except ValueError:
                continue  # 跳过非法格式的节假日行
        C.STATUTORY_HOLIDAYS = frozenset(hs)
    # 管理承载力（F_管理）公式系数
    manage = _cfg_get(f, "管理承载力", {}) or {}
    C.MANAGE_SITES_PER_PERSON = float(_cfg_get(manage, "每人可到现场数", C.MANAGE_SITES_PER_PERSON))
    C.OCCUPANCY_SAME_ENTRY = float(_cfg_get(manage, "同进同出占用(人天)", C.OCCUPANCY_SAME_ENTRY))
    C.OCCUPANCY_NORMAL = float(_cfg_get(manage, "一般到位占用(人天)", C.OCCUPANCY_NORMAL))
    same_risks = _cfg_get(manage, "需同进同出风险等级", None)
    if isinstance(same_risks, list):
        C.SAME_ENTRY_RISKS = tuple(str(x) for x in same_risks)
    mgr_kw = _cfg_get(manage, "管理岗关键字", None)
    if isinstance(mgr_kw, list) and mgr_kw:
        C.MANAGER_ROLE_KEYWORDS = tuple(str(x) for x in mgr_kw)
