# -*- coding: utf-8 -*-
"""
承载力量化计算引擎（技能「安全作业承载力分析」内置测算工具）—— CLI 入口
===================================================================

计算入口索引（想看“怎么算”直接跳到对应模块/函数）：
    | 想看                             | 跳到                                              |
    |----------------------------------|---------------------------------------------------|
    | 班组承载力怎么算                  | capacity.calculator.calc_team_capacity()          |
    | 工作负责人承载力                  | capacity.calculator.calc_leader_capacity()        |
    | 班组成员承载力                    | capacity.calculator.calc_member_capacity()        |
    | 工区承载力（加权）                | capacity.calculator.calc_area_capacity()          |
    | 周期（周/月/年）聚合              | capacity.calculator.calc_period()                 |
    | 日期×单位矩阵                     | capacity.calculator.build_matrix_day()            |
    | 四级预警分级                      | capacity.coefficients.classify_alert()            |
    | 组织架构归列                      | capacity.org.UnitNormalizer.unit_of()             |
    | β / θ 系数                       | capacity.coefficients.build_person_beta()/risk_theta() |
    | 输入加载（JSON/接口/Excel）        | capacity.parsing / capacity.excelLoader             |
    | 公式系数 / 配置文件               | capacity.config.apply_formula_config()            |

依据《安全承载力量化分析方法》实现：
  - 人员效能折算系数 β（配置「强制统一β」在 parse_person 读取；当前 true 时每人用默认 1.0，false 时用档案「作业素养能力」）
  - 作业风险系数 θ（工作内容或作业类型去空白后正好是「装表接电」时取「装表接电」0.2；空或「无」为 0；其余按作业风险等级）
  - 日承载力（工作负责人 / 班组成员 / 班组 / 工区）
  - 周 / 月度 / 年度聚合（分子按自然日计划工时、分母按工作日叠加）
  - 四级预警分级（配置套用后：轻载 ≤50 / 适度 50~75 / 重载 75~90 / 满载 >90；大于 100% 仍为满载，无停工档）

------------------------------------------------------------
用法与退出码
------------------------------------------------------------
  python calcCapacity.py data.json                 # 规范化 JSON(persons/plans/absence)，stdout 输出结果 JSON
  python calcCapacity.py workplan.xls [--config]   # Excel 作业计划：按配置读主表/列映射，日/周/月自动展开
  python calcCapacity.py --api '{"rows":[...]}'    # 接口返回 JSON(规范化结构 或 与 Excel 同构的表格行)
  python calcCapacity.py --xls workplan.xls --out r.json   # 指定按 Excel 读取并原子写盘
  python calcCapacity.py workplan.xlsx --period 周  # 支持 .xlsx；周期分桶输出(日/周/月/年)
  python calcCapacity.py --demo                    # 使用内置样例演示

  # 公式与输入解析均可在 config/capacity_config.json 中调整(脚本自动加载，也可用 --config 指定)
  # 组织架构归一化到报告矩阵单位列：config「输入解析.组织架构映射.是否启用=true」后输出 matrix

  退出码：0=成功  2=用法错误  3=输入数据错误  4=文件/IO 错误  1=未知异常
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from typing import Any, Dict, List, Optional

from capacity import constants as C
from capacity.config import (CONFIG_SECTION_INPUT, DEFAULT_INPUT_CONFIG,
                             apply_formula_config, load_config)
from capacity.calculator import CapacityService
from capacity.excelLoader import read_xls_dataset
from capacity.org import load_org_sheet_rows
from capacity.output import atomic_write_text, mask_result_names, read_demo
from capacity.parsing import build_dataset, read_dataset_file, read_interface_dataset

LOGGER = C.LOGGER


def build_main_parser() -> argparse.ArgumentParser:
    """命令行参数解析：保持 `--demo` / 数据文件 / `--out` 用法不变，新增 Excel / 接口 / 配置。"""
    parser = argparse.ArgumentParser(
        prog="calcCapacity.py",
        description="承载力量化计算引擎（技能「安全作业承载力分析」内置测算工具）",
        add_help=True,
    )
    parser.add_argument("data_file", nargs="?", default=None,
                        help="输入数据文件：.json(规范化 persons/plans/absence) 或 .xls/.xlsm(Excel 作业计划)")
    parser.add_argument("--xls", dest="xls_path", default=None, metavar="PATH",
                        help="直接按 Excel 读取作业计划（按配置的主表/列映射）")
    parser.add_argument("--api", dest="api_json", default=None, metavar="JSON",
                        help="接口返回的 JSON 字符串：规范化结构 或 与 Excel 同构的表格行")
    parser.add_argument("--config", dest="config_path", default=None, metavar="PATH",
                        help="配置文件(capacity_config.json)路径；缺省自动找技能目录 config/ 下同名文件")
    parser.add_argument("--period", dest="period", choices=["日", "周", "月", "年"], default="日",
                        help="统计周期粒度：日(默认，整体周期一条) / 周 / 月 / 年（逐桶输出 periods）")
    parser.add_argument("--demo", action="store_true",
                        help="使用内置样例演示")
    parser.add_argument("--out", dest="out_path", default=None, metavar="PATH",
                        help="将结果 JSON 原子写盘到指定路径（可选）")
    return parser


def _setup_logging() -> None:
    logging.basicConfig(level=logging.INFO, stream=sys.stderr,
                        format="[%(levelname)s] %(message)s")


def _load_dataset(args: argparse.Namespace, input_cfg: Dict[str, Any]):
    """按 CLI 参数路由到对应数据源，返回规范化 WorkDataset。"""
    if args.demo:
        LOGGER.info("使用内置样例计算")
        return build_dataset(read_demo())
    if args.xls_path:
        return read_xls_dataset(args.xls_path, input_cfg)
    if args.api_json:
        try:
            obj = json.loads(args.api_json)
        except json.JSONDecodeError as exc:
            raise C.InputDataError(
                f"--api 参数不是合法 JSON：{exc}", {"raw": args.api_json[:200]}) from exc
        return read_interface_dataset(obj, input_cfg)
    path = args.data_file
    if str(path).lower().endswith((".xls", ".xlsm", ".xlsx")):
        return read_xls_dataset(path, input_cfg)
    return read_dataset_file(path)


def main(argv: Optional[List[str]] = None) -> int:
    """程序入口：参数路由 -> 加载配置 -> 读取/解析(JSON/Excel/接口) -> 计算 -> 输出。"""
    args = build_main_parser().parse_args(argv)
    _setup_logging()

    # 用法校验：--demo / 数据文件 / --xls / --api 至少其一
    if not args.demo and not args.data_file and not args.xls_path and not args.api_json:
        print(__doc__)
        return C.EXIT_USAGE

    try:
        cfg = load_config(args.config_path)
        apply_formula_config(cfg)
        input_cfg = cfg.get(CONFIG_SECTION_INPUT) or DEFAULT_INPUT_CONFIG
        org_cfg = input_cfg.get("组织架构映射") or {}
        manage_cfg = input_cfg.get("管理承载力") or {}

        dataset = _load_dataset(args, input_cfg)

        # 组织架构：配置启用「优先读取源表组织架构」且源为 Excel 时，读「组织架构」sheet 归列
        source_path = args.xls_path or args.data_file
        org_sheet_rows = load_org_sheet_rows(source_path, input_cfg)

        result = CapacityService(dataset, org_cfg=org_cfg, manage_cfg=manage_cfg,
                                 org_sheet_rows=org_sheet_rows).summarize(period=args.period)
        mask_result_names(result)
        text = json.dumps(result, ensure_ascii=False, indent=2)

        if args.out_path:
            atomic_write_text(args.out_path, text)
            LOGGER.info("结果已原子写盘：%s", args.out_path)

        print(text)
        return C.EXIT_OK

    except C.CapacityError as exc:
        payload = exc.to_payload()
        LOGGER.error("%s: %s (context=%s)", exc.code, exc.message, exc.context)
        print(json.dumps(payload, ensure_ascii=False))
        return C._ERROR_EXIT.get(exc.code, C.EXIT_UNEXPECTED)
    except Exception as exc:  # noqa: BLE001 最后一层兜底，避免裸 traceback
        LOGGER.exception("未预期异常")
        print(json.dumps({
            "ok": False, "errorCode": C.ERR_UNEXPECTED,
            "message": f"未预期异常：{exc}",
        }, ensure_ascii=False))
        return C.EXIT_UNEXPECTED


if __name__ == "__main__":
    sys.exit(main())
