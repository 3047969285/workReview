# -*- coding: utf-8 -*-
"""
输出与演示数据（output：姓名脱敏 / 原子写盘 / 内置 demo 样例）
==============================================================

- `mask_result_names()`：结果 JSON 输出层的姓名掩码（姓+*），由配置「输出姓名脱敏」
  控制（现网口径名字不脱敏，默认关闭）；
- `atomic_write_text()`：先写同目录临时文件再 os.replace 原子替换（R2 写文件不裸改）；
- `read_demo()`：内置样例数据，用于演示 / 自测，可对照字段结构。
"""
from __future__ import annotations

import os
from typing import Any, Dict

from . import constants as C
from .constants import OutputWriteError


def mask_name(name: Any) -> str:
    """姓名脱敏：姓 + *（如“张**”）；空值原样返回。"""
    s = str(name or "")
    if not s:
        return s
    return s[0] + "*" * (len(s) - 1)


def mask_result_names(result: Dict[str, Any]) -> None:
    """结果 JSON 输出层的姓名掩码；计算内部不掩码（供未来 β 匹配）。"""
    if not C.MASK_NAMES:
        return
    pb = result.get("persons_beta") or {}
    for team in list(pb):
        pb[team] = {mask_name(n): v for n, v in (pb[team] or {}).items()}


def atomic_write_text(path: str, content: str, encoding: str = "utf-8") -> None:
    """原子写：写同目录临时文件后 os.replace 替换；失败原目标不动并抛 OutputWriteError。"""
    target = os.path.abspath(path)
    parent = os.path.dirname(target)
    tmp_path = os.path.join(parent, f".{os.path.basename(target)}.tmp")
    try:
        with open(tmp_path, "w", encoding=encoding, newline="") as fh:
            fh.write(content)
            fh.flush()
            os.fsync(fh.fileno())
        os.replace(tmp_path, target)
    except OSError as exc:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except OSError:
            pass
        raise OutputWriteError(f"结果写盘失败：{exc}", target, str(exc)) from exc


def read_demo() -> Dict[str, Any]:
    """内置样例数据（虚构姓名，供演示 / 自测，勿填真实个人信息）。"""
    persons = [
        {"name": "张伟", "team": "变电检修一班", "beta": 1.0, "is_leader": True},
        {"name": "李娜", "team": "变电检修一班", "beta": 1.0, "is_leader": False},
        {"name": "王建军", "team": "输电运检二班", "beta": 1.0, "is_leader": True},
        {"name": "赵敏", "team": "输电运检二班", "beta": 1.0, "is_leader": False},
    ]
    plans = [
        {"name": "220kV 示例东线立塔", "team": "输电运检二班", "leader": "王建军",
         "members": 4, "risk": "三级", "hours": 6,
         "start": "2026-08-24", "end": "2026-08-27", "type": "检修施工"},
        {"name": "#1 主变例行试验", "team": "变电检修一班", "leader": "张伟",
         "members": 3, "risk": "四级", "hours": 7,
         "start": "2026-08-25", "end": "2026-08-26", "type": "检修施工"},
        {"name": "10kV 江南线巡视", "team": "变电检修一班", "leader": "李娜",
         "members": 2, "risk": "五级", "hours": 4,
         "start": "2026-08-26", "end": "2026-08-26", "type": "巡视"},
    ]
    absence = [
        {"name": "王建军", "start": "2026-08-25", "end": "2026-08-26"},  # 出差
    ]
    return {"persons": persons, "plans": plans, "absence": absence}
