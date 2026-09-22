# -*- coding: utf-8 -*-
"""
常量、系数表、错误码与统一异常（定义层）
====================================

所有公式系数、字段别名、错误码 / 退出码与统一异常都集中在此模块，
是其它模块的"唯一事实源"。`capacity/config.py` 的 `apply_formula_config()`
通过本模块的模块级变量覆盖公式系数（可调参数），计算模块一律从这里读取。

注释标 [班组承载力] [工区承载力] 等的中文锚点即各容量计算入口，
想看"班组怎么算"直接跳到 `capacity/calculator.py` 的 calc_team_capacity()。
"""
from __future__ import annotations

import logging
from typing import Any, Dict, Tuple

# 统一日志器：写入 stderr，保持 stdout 只输出结果
LOGGER = logging.getLogger("calcCapacity")

# =============================================================================
# 一、公式系数与常量（与现网《安全承载力量化分析方法》口径一致，禁止随意改动）
# =============================================================================

# 标准日工时（小时）
STD_HOURS_PER_DAY = 8
# 作业时长 T 规整边界：不足 1 小时按 1 小时，超过 8 小时按 8 小时
MIN_WORK_HOURS = 1.0
MAX_WORK_HOURS = 8.0
# 人员效能系数 β。FORCE_BETA / DEFAULT_BETA 是套配置前的初值。
# apply_formula_config 会用配置「强制统一β」「默认人员效能系数β」覆盖。
# true：parse_person 让每个人都用默认 β。false：用档案「作业素养能力」，缺省才用默认 β。
# 不是把 β 永久写死为 1.0。当前配置文件里「强制统一β」为 true。
DEFAULT_BETA = 1.0
FORCE_BETA = True
# 工作负责人承载力上限钳制：避免除数过小导致数值异常放大
LEADER_LOAD_CAP = 999.0
# 分母保护值：规避除零，同时保持与原口径一致的极小分母
DENOM_EPS = 0.0001

# 作业风险系数 θ
RISK_COEF: Dict[str, float] = {
    "二级": 1.5,
    "三级": 1.2,
    "四级": 1.0,
    "五级": 0.8,
    "装表接电": 0.2,
    "装表": 0.2,
}
# 未知作业风险等级的回退 θ（按最保守一档四级）
RISK_COEF_DEFAULT = 1.0
# 无风险等级（空 / "无"）时，任务不计入承载力
RISK_COEF_NONE = 0.0

# ---- 管理承载力（F_管理 = 实需人天 ÷ 可供人天）----
# 每人可同时到场的现场数：分母 = 在家管理人员数 × 该值（主体一人可去两个现场）
MANAGE_SITES_PER_PERSON = 2
# 现场占用（人天）：需同进同出的现场占 2，一般到位占 1
OCCUPANCY_SAME_ENTRY = 2
OCCUPANCY_NORMAL = 1
# 需同进同出的作业风险等级集合（业务口径：作业风险等级 3 级以上 = 二级 / 三级）
SAME_ENTRY_RISKS: Tuple[str, ...] = ("二级", "三级")
# 管理岗判定关键字：命中「角色 / 职务」任一即计入该单位在家管理人员数
MANAGER_ROLE_KEYWORDS: Tuple[str, ...] = (
    "管理", "主任", "副主任", "书记", "经理", "总经理", "副总经理",
    "专责", "所长", "班长",
)

# ---- 工作日日历（默认剔除：周末 + 法定节假日）----
# 口径（v2.0）：周期（周/月/年）累加的分母用“工作日日集”，剔除周六周日与法定节假日。
# 口径调整（v2.3）：**日维度放开周末节假日**——days / matrix / management.days 按计划覆盖的
# 全部自然日展开（周六周日/法定节假日若当天有计划照常参与日计算，“计划中有周末/节假日不要去掉、
# 周末承载力和日常一样”）；周末/节假日剔除**只**发生在周期累加：calc_period 的分母、management
# 区间聚合的分母仍按工作日日历（分子 remain=计划实际安排，见 calculator）。
# 配置「非工作日.是否剔除周末」=true 剔除周末，「非工作日.法定节假日」清单剔除法 定节假日
#（按国务院当年放假安排维护）。若需完全按计划时间计算，把开关置 false 且清空节假日清单即可。
EXCLUDE_WEEKENDS = True
NON_WORK_WEEKDAYS: Tuple[int, ...] = (5, 6)  # date.weekday(): 5=周六, 6=周日
STATUTORY_HOLIDAYS: frozenset = frozenset()  # frozenset[date]


def is_workday(d) -> bool:
    """周期（周/月/年）累加可计入分母的日期：剔除周末（启用时）与清单内法定节假日。

    被剔除的日子不进入 calc_period 累加分母与 management 区间聚合的分母天数；
    这些日子上的计划工时 / 实需人天仍计入分子。日维度（days / matrix /
    management.days）调用方不使用本函数，直接按自然日展开。
    """
    if EXCLUDE_WEEKENDS and d.weekday() in NON_WORK_WEEKDAYS:
        return False
    return d not in STATUTORY_HOLIDAYS

# 班组类型（现网口径暂按"检修施工"组合；预留运检合一 / 变电运维扩展）
TEAM_KIND_DEFAULT = "检修施工"
# 作业类型识别集合
PATROL_TYPES: Tuple[str, ...] = ("巡视",)
OPERATION_TYPES: Tuple[str, ...] = ("倒闸", "操作")

# 四级预警分级规则：(阈值下界, 级别)，按降序匹配，命中即返回。
# 下面这组是 apply_formula_config 之前的预置表，仍含「停工(超满载)」。
# 加载 capacity_config.json 后会被预警分级 90/75/50 整表替换：大于 90 为满载，
# 大于 100% 也是满载，没有停工档。不要把预置表当成现网口径；现网以配置为准。
ALERT_RULES: Tuple[Tuple[float, str], ...] = (
    (100.0, "停工(超满载)"),
    (90.0, "满载"),
    (75.0, "重载"),
    (50.0, "适度"),
)
ALERT_LIGHT = "轻载"
# 输出是否对姓名脱敏(姓 + *)。计算内部仍保留原始姓名，便于未来按档案匹配 β。
# 用户口径调整（v1.5）：名字不需要脱敏，报告 / 函件一律输出真名，默认关闭脱敏。
# 敏感字段(身份证号/联系方式/地址等)在解析层即不进入。
MASK_NAMES = False

# 报告矩阵表的单位清单（中心 / 公司，与官方周报模板表头一致）。
# 组织架构归一化启用时，把计划所属的 公司/工区/班组 收敛到这 17 个单位列；
# 未命中的归入"其他"。
REPORT_UNITS: Tuple[str, ...] = (
    "变电运维", "变电检修", "输电中心", "带电中心", "电缆中心",
    "市南中心", "市北中心", "崂山中心", "李沧中心", "城阳中心",
    "安装公司", "送变电", "胶州公司", "即墨公司", "黄岛公司",
    "平度公司", "莱西公司",
)
ORG_CATEGORY_OTHER = "其他"

# 字段别名：(主键, 次键, ...) —— 兼容 31 字段作业计划模板中部分中文列名；
# 主键优先，次键仅为缺省回退，保证旧输入行为不变、新接入更兼容
FIELD_ALIASES: Dict[str, Tuple[str, ...]] = {
    # 作业班组
    "team": ("team", "班组", "现场施工单位班组"),
    # 作业开始 / 结束时间
    "start": ("start", "计划开始时间"),
    "end": ("end", "计划结束时间"),
    # 作业风险等级
    "risk": ("risk", "作业风险等级"),
    # 作业时长（小时）
    "hours": ("hours", "时长", "工作预计时长(小时)"),
    # 工作负责人
    "leader": ("leader", "工作负责人", "负责人"),
    # 所需自有人员人数（不含工作负责人）
    "members": ("members", "所需自有人员人数", "所需自有人员人数(不含工作负责人)"),
    # 作业类型（巡视 / 倒闸 / 操作 等）
    "type": ("type", "作业类型"),
    # 工作内容（装表接电的 θ 看这一列的取值，也看作业类型）
    "content": ("content", "工作内容"),
    # 工程名称
    "name": ("name", "工程名称"),
    # 业主单位公司 / 工区（用于组织架构归一化到报告矩阵中心 / 公司列）
    "company": ("company", "company_name", "公司", "业主单位公司"),
    "work_area": ("work_area", "workArea", "工区", "业主单位工区"),
}
# 人员档案字段别名
PERSON_FIELD_ALIASES: Dict[str, Tuple[str, ...]] = {
    "name": ("name", "姓名", "人员姓名"),
    "team": ("team", "所属班组", "班组"),
    "beta": ("beta", "人员效能系数β", "人员效能系数"),
    "is_leader": ("is_leader", "是否工作负责人", "是否负责人"),
    # 人员所属单位(工区) / 职务：管理承载力统计“在家管理人员数”用（可选）
    "work_area": ("work_area", "单位（工区）", "单位工区", "工区"),
    "role": ("role", "职务", "角色"),
    # 班组类型：检修施工 / 运检合一 / 变电运维。计算时按班组取值，不再只用全局默认。
    "team_kind": ("team_kind", "班组类型"),
    # 人员所属 市/县 公司：管理承载力按“县公司兜底、工区优先”归列（可选）
    "county": ("county", "县公司", "市公司"),
}
# 请假/出差字段别名
ABSENCE_FIELD_ALIASES: Dict[str, Tuple[str, ...]] = {
    "name": ("name", "姓名", "人员姓名"),
    "start": ("start", "开始时间", "开始日期", "start_date"),
    "end": ("end", "结束时间", "结束日期", "end_date"),
}

# =============================================================================
# 二、错误码与统一异常（失败必须携带 errorCode）
# =============================================================================

ERR_USAGE = "USAGE_ERROR"
ERR_INPUT_INVALID = "INPUT_INVALID"
ERR_FILE_NOT_FOUND = "FILE_NOT_FOUND"
ERR_FILE_INVALID = "FILE_INVALID"
ERR_OUTPUT_WRITE_FAILED = "OUTPUT_WRITE_FAILED"
ERR_UNEXPECTED = "UNEXPECTED"

EXIT_OK = 0
EXIT_USAGE = 2
EXIT_INPUT = 3
EXIT_IO = 4
EXIT_UNEXPECTED = 1

# 错误码 -> 退出码映射
_ERROR_EXIT: Dict[str, int] = {
    ERR_USAGE: EXIT_USAGE,
    ERR_INPUT_INVALID: EXIT_INPUT,
    ERR_FILE_NOT_FOUND: EXIT_IO,
    ERR_FILE_INVALID: EXIT_IO,
    ERR_OUTPUT_WRITE_FAILED: EXIT_IO,
    ERR_UNEXPECTED: EXIT_UNEXPECTED,
}


class CapacityError(Exception):
    """承载力量化引擎统一异常基类：携带错误码与上下文，供上层定位与响应。"""

    code = ERR_UNEXPECTED

    def __init__(self, message: str, context: Dict[str, Any] | None = None) -> None:
        super().__init__(message)
        self.message = message
        self.context = context or {}

    def to_payload(self) -> Dict[str, Any]:
        """输出层统一响应体（失败必须带 errorCode，参照成熟方案统一契约）。"""
        return {"ok": False, "errorCode": self.code, "message": self.message}


class UsageError(CapacityError):
    """参数 / 用法错误。"""

    code = ERR_USAGE


class InputDataError(CapacityError):
    """输入数据校验失败：字段缺失、类型非法、日期无法解析等。"""

    code = ERR_INPUT_INVALID


class FileAccessError(CapacityError):
    """输入文件不存在或不可读。"""

    code = ERR_FILE_NOT_FOUND

    def __init__(self, message: str, path: str | None = None) -> None:
        super().__init__(message, {"path": path})


class FileContentError(CapacityError):
    """输入文件内容非法：JSON 解析失败或结构不正确。"""

    code = ERR_FILE_INVALID

    def __init__(self, message: str, path: str | None = None,
                 detail: str | None = None) -> None:
        super().__init__(message, {"path": path, "detail": detail})


class OutputWriteError(CapacityError):
    """结果文件原子写盘失败。"""

    code = ERR_OUTPUT_WRITE_FAILED

    def __init__(self, message: str, path: str | None = None,
                 detail: str | None = None) -> None:
        super().__init__(message, {"path": path, "detail": detail})
