# -*- coding: utf-8 -*-
"""
承载力量化报告 / 风险函件填充引擎（技能「安全作业承载力分析」内置工具）
====================================================================

依据技能 templates/ 目录下的 6 份官方公文模板（.doc），将结构化的填充参数
逐项替换进模板正文与附表，生成正式成品报告 / 函件。模板文件只读，不做任何改动。

支持的模板（--template 取值见 TEMPLATES）：
  - 警示函-满载      安全风险警示函（承载力满载 >90%）
  - 提示函-重载      安全风险提示函（承载力重载 >75%）
  - 提示函-长期作业  安全风险提示函（工作负责人连续作业超限）
  - 报告-日          国网青岛供电公司安全承载力分析报告（日）
  - 报告-周          国网青岛供电公司安全承载力分析报告（周）
  - 报告-月          国网青岛供电公司安全承载力分析报告（月）

------------------------------------------------------------
一、输入 / 输出
------------------------------------------------------------
  python fillReport.py --template 报告-周 --params params.json --out 成品.doc

  --params JSON 结构（replace + tables）：
  {
    "replace": {                       // 全文替换映射（标题、正文、占位符均可）
      "2026年X月X日": "2026年9月3日",
      "X1": "变电检修中心",             // 注意先长后短、具体先于通配，避免子串误替换
      "X2": "输电中心",
      "XX1": "市南中心"
    },
    "paragraph_replace": [             // 段落级替换（anchor 命中第 N 处）
      {
        "anchor": "管理承载力分析XXXXXXXXXX",   // 定位锚（须为文本子串，全文按第几处命中定位）
        "occurrence": 2,                        // 第几处命中（从 1 起；默认 1）
        "new_text": "9月份第2周管理承载力为86%"  // 默认：仅替换锚定子串，段尾句号/段落结束符保留；
                                                // 置 "next_paragraph": true 时整段替换锚点段落的下一段
      }
    ],
    "tables": [                        // 附表填充（按表头列名定位，0-N 张均可）
      {
        "columns": ["序号", "单位", "作业内容", "风险等级", "高风险作业时间"],
        "rows": [
          ["1", "变电检修中心", "#1主变综合检修", "三级", "9月3日-5日（3天）"],
        ],
        // "anchor": "各单位日计划如表所示",            // 可选：锚点句定位目标表
        // "max_units_per_table": 8                     // 可选：单位列数（总列数-1）
                                                      //   超阈值时自动纵切多张上下
                                                      //   堆叠子表（表间「（续表）」段）
      }
    ],
    "trim_month_weeks": 5              // 可选、仅月报：实际周数；<5 时删多余周节
  }

  --demo 演示：内置示例参数，快速验证模板链路。

  退出码：0=成功  2=用法错误  3=输入数据错误  4=文件/IO 错误  1=未知异常
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
import time
from typing import Any, Dict, List, Optional

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, "..", "templates"))

# 单实例 WPS：自动化挂到与用户会话共用的同一后端上，迭代多次后偶发「文档保存
# 失败」(3011)。SaveAs 短暂恢复间隙重试，属外部 GUI 自动化的常规稳健兜底。
_SAVE_MAX_TRIES = 3
_SAVE_RETRY_GAP = 1.5  # 秒

# Windows 非法文件名字符。出口文件名含 ASCII `*`（如姓名脱敏「姓+*」未全角化时）
# 会让 WPS SaveAs 稳定报 3011「文档保存失败」，且与并发无关、极难排查；此处兜底
# 在填充前给出明确报错，非法字符的全角化应由上游（genLetterParams._safe_filename）
# 完成。
_ILLEGAL_FN = '\\/:*?"<>|'

# 模板名（--template 取值）-> 模板文件名
TEMPLATES: Dict[str, str] = {
    "警示函-满载": "安全风险警示函-（XX单位XX时间承载力满载）.doc",
    "提示函-重载": "安全风险提示函（XX单位XX时间承载力重载）.doc",
    "提示函-长期作业": "安全风险提示函（XX单位XXX长期作业）.doc",
    "报告-日": "国网青岛供电公司2026年X月X日安全承载力分析报告.doc",
    "报告-周": "国网青岛供电公司2026年X月份第X周安全承载力分析报告.doc",
    "报告-月": "国网青岛供电公司2026年X月份安全承载力分析报告.doc",
}

# 统一日志器：写入 stderr，stdout 仅输出结果 JSON
LOGGER = logging.getLogger("fillReport")

# 统一错误码与退出码（与 calcCapacity.py 契约一致）
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
_ERROR_EXIT = {
    ERR_USAGE: EXIT_USAGE,
    ERR_INPUT_INVALID: EXIT_INPUT,
    ERR_FILE_NOT_FOUND: EXIT_IO,
    ERR_FILE_INVALID: EXIT_IO,
    ERR_OUTPUT_WRITE_FAILED: EXIT_IO,
    ERR_UNEXPECTED: EXIT_UNEXPECTED,
}


class FillError(Exception):
    """填充引擎统一异常：携带错误码与上下文。"""

    code = ERR_UNEXPECTED

    def __init__(self, message: str, context: Optional[Dict[str, Any]] = None) -> None:
        super().__init__(message)
        self.message = message
        self.context = context or {}

    def to_payload(self) -> Dict[str, Any]:
        return {"ok": False, "errorCode": self.code, "message": self.message}


class UsageError(FillError):
    code = ERR_USAGE


class InputError(FillError):
    code = ERR_INPUT_INVALID


class FileError(FillError):
    code = ERR_FILE_NOT_FOUND


class WordError(FillError):
    code = ERR_UNEXPECTED


# =============================================================================
# Word 文档操作（win32com，Windows + Office/WPS 环境）
# =============================================================================

# Word 文件格式常量
_WD_FORMAT_DOC = 0      # .doc
_WD_FORMAT_DOCX = 12    # .docx

_WD_CHARACTER = 1
_WD_PARAGRAPH = 4  # wdParagraph（勿改成 7=wdScreen）

# runReports 整批可挂入同一隔离 Word/WPS 实例，避免每步 DispatchEx + Quit。
# 未调用 begin_shared_word 时，各脚本仍各自起实例（直接 CLI 行为不变）。
_SHARED_APP = None


def current_word_app():
    """返回当前共用 Word 实例；无共用时为 None。"""
    return _SHARED_APP


def _launch_word_app():
    """DispatchEx 隔离新实例：先 Word.Application，失败再试 WPS ProgID。"""
    import win32com.client as win32
    last_exc = None
    for prog in ("Word.Application", "Kwps.Application", "KWPS.Application"):
        app = None
        try:
            app = win32.DispatchEx(prog)
            app.Visible = False
            try:
                app.DisplayAlerts = False
            except Exception:
                pass
            LOGGER.info("已启动隔离 Word/WPS：%s", prog)
            return app
        except Exception as exc:  # noqa: BLE001 下一 ProgID 再试
            last_exc = exc
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
    raise WordError(f"无法启动 Word COM：{last_exc}") from last_exc


def begin_shared_word() -> None:
    """为整条流水线挂入一个隔离 Word 实例（可重复调用，已有则复用）。"""
    global _SHARED_APP
    if _SHARED_APP is not None:
        return
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass
    _SHARED_APP = _launch_word_app()


def end_shared_word() -> None:
    """退出共用 Word 实例；无共用时为 no-op。"""
    global _SHARED_APP
    app = _SHARED_APP
    _SHARED_APP = None
    if app is None:
        return
    try:
        # 先关干净再 Quit，避免残留文档锁下一份模板
        while True:
            try:
                n = int(app.Documents.Count)
            except Exception:
                break
            if n < 1:
                break
            try:
                app.Documents(1).Close(SaveChanges=False)
            except Exception:
                break
    except Exception:
        pass
    try:
        app.Quit()
    except Exception:
        pass


def recycle_shared_word() -> None:
    """关掉当前共用实例再拉起新的（--all 日→周→月之间调用，清 COM 脏状态）。"""
    end_shared_word()
    begin_shared_word()


def _close_open_documents(app) -> None:
    """关闭 app 上全部未保存文档，避免下一 Open 撞锁。"""
    if app is None:
        return
    for _ in range(32):
        try:
            if int(app.Documents.Count) < 1:
                return
            app.Documents(1).Close(SaveChanges=False)
        except Exception:
            return


def _replace_next_paragraph(doc, anchor_rng, new_text: str) -> None:
    """整段替换「anchor_rng 所在段落」的下一段正文（保留段落结束符）。

    用于覆盖模板中紧跟在稳定标题行之后的硬编码示例段落。实现：
    取 anchor 段落的 Range.End（含段尾 \\r）作为下一段起点，用
    MoveEnd(wdParagraph) 展开为下一整段，再 MoveEnd(wdCharacter,-1)
    回退掉段尾 \\r，仅换段内正文 —— 段落标记与段落计数保持不变，
    不会与再下一段合并。失败时静默跳过（WPS 兼容性兜底）。
    """
    try:
        para_rng = anchor_rng.Paragraphs(1).Range
        nxt = doc.Range(para_rng.End, para_rng.End)
        nxt.MoveEnd(_WD_PARAGRAPH, 1)
        nxt.MoveEnd(_WD_CHARACTER, -1)
        if nxt.Text.strip():
            nxt.Text = new_text
    except Exception as exc:  # noqa: BLE001 WPS 下某些结构无法导航
        LOGGER.warning("next_paragraph 替换跳过：%s", exc)


class WordSession:
    """封装 Word COM 会话：打开模板副本 -> 替换/填表 -> 另存为成品。

    模板始终只读打开；成品的占位符替换与表格填充均在内存中进行，
    关闭时对模板不保存（SaveChanges=False）。
    """

    def __init__(self) -> None:
        try:
            import win32com.client as win32  # noqa: F401 仅探测依赖
        except ImportError as exc:
            raise WordError("未安装 pywin32（需 pip install pywin32），无法驱动 Word") \
                from exc
        shared = current_word_app()
        if shared is not None:
            self.app = shared
            self._owns_app = False
            return
        try:
            # DispatchEx 起隔离新实例，避免附着用户/残留 WPS 会话引发 3011
            self.app = _launch_word_app()
            self._owns_app = True
        except WordError:
            raise
        except Exception as exc:
            raise WordError(f"无法启动 Word COM：{exc}") from exc

    def render(self, template_path: str, params: Dict[str, Any],
               out_path: str, out_as_doc: bool) -> None:
        """将 params 渲染进模板并另存为成品文件。"""
        import win32com.client as win32
        _close_open_documents(self.app)
        doc = None
        open_err: Optional[Exception] = None
        for attempt in range(2):
            try:
                doc = self.app.Documents.Open(
                    template_path, ReadOnly=True, AddToRecentFiles=False)
                open_err = None
                break
            except Exception as exc:  # noqa: BLE001
                open_err = exc
                # 共用实例偶发脏状态：回收后重试一次
                if attempt == 0 and not self._owns_app:
                    LOGGER.warning("模板打开失败，回收共用 Word 后重试：%s", exc)
                    recycle_shared_word()
                    self.app = current_word_app()
                    _close_open_documents(self.app)
                else:
                    break
        if doc is None:
            raise FileError(f"模板无法打开：{template_path}", {"path": template_path}) \
                from open_err
        try:
            # 顺序：先普通全文替换（含月度总管理句），再段落级替换（逐周管理句）。
            # 月度总句通过较长的普通 replace key 先行替换，吞掉其 anchor 子串，
            # 确保后续 Find 仅命中剩余的同文周段落，occurrence 从 1 开始稳定计数。
            self._apply_replace(doc, params.get("replace") or {})
            self._apply_para_replace(doc, params.get("paragraph_replace") or [])
            for spec in params.get("tables") or []:
                self._fill_table(doc, spec)
            keep_weeks = params.get("trim_month_weeks")
            if keep_weeks is not None:
                self._trim_extra_month_weeks(doc, int(keep_weeks))
            # native 路径不经 applyCharts：此处统一把超版心表（周风险明细等）收进版心
            try:
                from pngCharts import fit_tables_to_text
                fit_tables_to_text(self.app, doc)
            except Exception as exc:  # noqa: BLE001
                LOGGER.warning("宽表收窄跳过：%s", exc)
            # 表内段落若带 KeepWithNext，会把表后正文/图整组顶到下一页，
            # 留下「表后大半页空白、图独霸下页」——明细表必须关掉。
            self._relax_table_keeps(doc)
            # 表居中、明细列宽统一、行距略压——给表后数量图腾位置
            try:
                self._normalize_report_tables(doc)
            except Exception as exc:  # noqa: BLE001
                LOGGER.warning("表格归一跳过：%s", exc)
            file_format = _WD_FORMAT_DOC if out_as_doc else _WD_FORMAT_DOCX
            last_save_err: Optional[Exception] = None
            for attempt in range(1, _SAVE_MAX_TRIES + 1):
                try:
                    doc.SaveAs(out_path, FileFormat=file_format)
                    last_save_err = None
                    break
                except Exception as exc:  # noqa: BLE001 WPS 单实例偶发 3011
                    last_save_err = exc
                    if attempt < _SAVE_MAX_TRIES:
                        time.sleep(_SAVE_RETRY_GAP * attempt)
            if last_save_err is not None:
                raise WordError(f"渲染保存失败：{last_save_err}") from last_save_err
            LOGGER.info("已生成成品：%s", out_path)
        except FillError:
            raise
        except Exception as exc:
            raise WordError(f"渲染失败：{exc}") from exc
        finally:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass

    @staticmethod
    def _relax_table_keeps(doc) -> None:
        """关掉各表单元格的 KeepWithNext/KeepTogether；明细表禁止行跨页拆开。

        模板或粘贴残留的 KeepWithNext 会把「表尾行 + 后续图组」绑死；
        图组稍高就整组翻页，表页底部留下大片空白。
        明细表（首列序号）若 AllowBreakAcrossPages=True，跨页续行会只剩
        「高风险作业时间」里 \\v 后的「（N天）」，前几列呈空白——必须禁拆行。
        """
        for ti in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(ti)
            try:
                tbl.AllowAutoFit = False
            except Exception:
                pass
            is_detail = False
            try:
                h0 = tbl.Cell(1, 1).Range.Text.replace("\r", " ").replace("\x07", " ").strip()
                is_detail = "序号" in h0
            except Exception:
                is_detail = False
            try:
                for r in range(1, tbl.Rows.Count + 1):
                    try:
                        # 明细整行同页；其它表仍允许跨页以免图组被粘死
                        tbl.Rows(r).AllowBreakAcrossPages = (not is_detail)
                    except Exception:
                        pass
                    for c in range(1, tbl.Columns.Count + 1):
                        try:
                            pf = tbl.Cell(r, c).Range.ParagraphFormat
                            pf.KeepWithNext = False
                            pf.KeepTogether = False
                        except Exception:
                            pass
            except Exception as exc:  # noqa: BLE001
                LOGGER.warning("表%d KeepWithNext 清理跳过：%s", ti, exc)

    @staticmethod
    def _detail_col_prefs(columns: List[str], text_w: float) -> List[float]:
        """明细表列宽：委托 capacity.office_chart 版面策略（唯一配方源）。"""
        from capacity.office_chart import detail_col_widths
        return detail_col_widths(columns, text_w)

    @staticmethod
    def _normalize_report_tables(doc) -> None:
        """报告表统一观感：整表居中；明细表列宽/字号/行距读 report_layout。

        不改正文 29pt。明细略紧便于表后接数量图；跨页续表与首表同配方。
        改观感只改 capacity/office_chart.py 版面策略区，勿在此写死 pt。
        """
        from capacity.office_chart import (
            DETAIL_FONT_PT, DETAIL_LINE_PT, TEXT_WIDTH_FALLBACK_PT)

        try:
            ps = doc.PageSetup
            text_w = float(ps.PageWidth - ps.LeftMargin - ps.RightMargin)
        except Exception:
            text_w = TEXT_WIDTH_FALLBACK_PT
        for ti in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(ti)
            try:
                # wdAlignRowCenter=1：整表相对版心居中
                tbl.Rows.Alignment = 1
            except Exception:
                pass
            try:
                tbl.AllowAutoFit = False
            except Exception:
                pass
            is_detail = False
            headers: List[str] = []
            try:
                for c in range(1, tbl.Columns.Count + 1):
                    h = tbl.Cell(1, c).Range.Text.replace("\r", " ").replace("\x07", " ").strip()
                    headers.append(h)
                is_detail = bool(headers) and "序号" in headers[0]
            except Exception:
                is_detail = False
            if not is_detail:
                # 非明细：仍锁总宽≈版心，观感齐
                try:
                    tbl.PreferredWidthType = 3  # wdPreferredWidthPoints
                    tbl.PreferredWidth = text_w
                except Exception:
                    pass
                continue
            prefs = WordSession._detail_col_prefs(headers, text_w)
            for c, width in enumerate(prefs, start=1):
                try:
                    tbl.Columns(c).PreferredWidthType = 3
                    tbl.Columns(c).PreferredWidth = width
                    tbl.Columns(c).Width = width
                except Exception:
                    try:
                        tbl.Columns(c).Width = width
                    except Exception:
                        pass
            try:
                tbl.PreferredWidthType = 3
                tbl.PreferredWidth = min(text_w, sum(prefs))
            except Exception:
                pass
            for r in range(1, tbl.Rows.Count + 1):
                try:
                    # 明细行禁止跨页拆开：否则下页只剩「（N天）」等续段，前几列呈空白
                    tbl.Rows(r).HeightRule = 0  # auto
                    tbl.Rows(r).AllowBreakAcrossPages = False
                except Exception:
                    pass
                for c in range(1, tbl.Columns.Count + 1):
                    try:
                        cell = tbl.Cell(r, c)
                        cell.VerticalAlignment = 1
                        rng = cell.Range
                        try:
                            rng.Font.NameFarEast = "宋体"
                            rng.Font.Name = "宋体"
                            rng.Font.Size = DETAIL_FONT_PT
                            rng.Font.Bold = True if r == 1 else False
                        except Exception:
                            pass
                        pf = rng.ParagraphFormat
                        pf.SpaceBefore = 0
                        pf.SpaceAfter = 0
                        pf.LineSpacingRule = 4  # wdLineSpaceExactly
                        pf.LineSpacing = DETAIL_LINE_PT
                        pf.KeepWithNext = False
                        pf.KeepTogether = False
                    except Exception:
                        pass

    @staticmethod
    def _trim_extra_month_weeks(doc, keep_weeks: int) -> None:
        """月报不足 5 周时删除多余周节（含 5 月示例正文与图位），止于「三、」章。

        触发：params.trim_month_weeks < 5。按「N.…第N周」标题定位，从第一处
        超出 keep_weeks 的周标题删到「三、辅助建议」前一段。keep≥5 为 no-op。
        """
        if keep_weeks >= 5:
            return
        week_re = re.compile(r"^([1-5])\..*第([1-5])周")
        start_idx = None
        end_idx = None
        n = doc.Paragraphs.Count
        for i in range(1, n + 1):
            t = doc.Paragraphs(i).Range.Text.rstrip("\r\n\x07").strip()
            m = week_re.match(t)
            if m and int(m.group(2)) > keep_weeks and start_idx is None:
                start_idx = i
            if t.startswith("三、") and start_idx is not None:
                end_idx = i
                break
        if start_idx is None or end_idx is None or end_idx <= start_idx:
            return
        for i in range(end_idx - 1, start_idx - 1, -1):
            try:
                doc.Paragraphs(i).Range.Delete()
            except Exception:
                pass

    @staticmethod
    def _apply_replace(doc, mapping: Dict[str, str]) -> None:
        """全文子串替换：主体段落 + 表格单元格。

        Range.Text 整体赋值可规避 Word 查找/替换 255 字符上限，
        段落与单元格样式由段落标记/表格样式继承，保持不变。
        仅做纯子串替换（键匹配文本中段，不含段落结束符 \r / 单元格结束符 \x07），
        因此段落计数不变、不会合并段落，动态 Paragraphs(i) 索引保持稳定。
        """
        if not mapping:
            return

        def _sub(each: str) -> str:
            # 先长后短替换：避免短键（如 X1）先命中破坏长键（如 XX1）
            for key in sorted(mapping.keys(), key=len, reverse=True):
                if key in each:
                    each = each.replace(key, mapping[key])
            return each

        # 主体段落
        for i in range(1, doc.Paragraphs.Count + 1):
            rng = doc.Paragraphs(i).Range
            new = _sub(rng.Text)
            if new == rng.Text:
                continue
            # 只重写段内文本、保留段落标记 \r：整段 Paragraph.Range 重写会让 WPS
            # 重算边界、把紧随其后的表格边界上吞，致正文段被并入下一表格首格
            # （锚点句吸进矩阵表首格，anchor 定位随之失灵）。段内子 range 不含
            # 段落标记，赋值不触碰表格边界，天然规避该问题。
            if new.endswith("\r"):
                new = new[:-1]
            body = doc.Range(rng.Start, rng.End - 1)  # 不含段落标记 \r
            body.Text = new
        # 表格单元格
        for ti in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(ti)
            for ri in range(1, tbl.Rows.Count + 1):
                for ci in range(1, tbl.Columns.Count + 1):
                    try:
                        cell = tbl.Cell(ri, ci)
                    except Exception:
                        continue
                    rng = cell.Range
                    new = _sub(rng.Text)
                    if new != rng.Text:
                        rng.Text = new

    @staticmethod
    def _apply_para_replace(doc, specs: List[Dict[str, Any]]) -> None:
        """段落级替换：anchor + occurrence 定位，仅替换 anchor 命中的子串区域。

        语义：把「anchor 在全文中的第 N 处命中」替换为 new_text。anchor 只须
        是段落文本的某个子串；替换只落在该子串上，段落的其余部分（句尾标点、
        段落结束符 \\r）原样保留——因此不会合并段落、段落序号保持稳定，可精确
        定位「多段同文」占位（如周报里每个周的管理承载力占位均为
        '管理承载力分析XXXXXXXXXX'，用 occurrence 1..5 逐周写入）。

        扩展模式：spec 置 next_paragraph=true 时，命中第 N 处后不再做子串替换，
        而是整段替换「锚点所在段落的下一段」正文（保留段落结束符）。用于覆盖
        模板中紧跟在某个稳定标题（如月报每周的'（1）作业承载力分析'）之后的
        硬编码示例段落，避免示例内容泄漏进成品。

        实现采用原生 Find.Execute 逐次命中（Forward=True / Wrap=wdFindStop）
        在 doc.Content 上从前往后扫描：命中一次计数 +1，若该 occurrence 有对应
        new_text 则按模式替换，随后向区间末尾折叠并重新下发 Find，确保
        下一轮从替换点之后继续向后找（不会死循环）。整个流程不依赖
        Paragraphs(i) 动态索引，规避 WPS 下段落漂移 / E_FAIL。

        注意：更长的普通替换键（如月度总句 'X月份，管理承载力分析XXXXXXXXXX'）
        应放进 replace（先执行），它会吞掉自身 anchor 子串，保证此处 Find
        只命中剩余的同文周段落，occ 从 1 稳定计数。
        """
        if not specs:
            return
        # 按 anchor 分组：anchor -> {occurrence: 完整 spec}（支持多条同 anchor 不同 occurrence）
        groups: Dict[str, Dict[int, Dict[str, Any]]] = {}
        for sp in specs:
            anchor = sp.get("anchor")
            if not anchor:
                continue
            occ = int(sp.get("occurrence", 1))
            text = sp.get("new_text")
            if text is None:
                continue
            groups.setdefault(anchor, {})[occ] = dict(sp)

        for anchor, occ_map in groups.items():
            rng = doc.Content
            f = rng.Find
            f.ClearFormatting()
            f.Text = anchor
            f.Forward = True
            f.Wrap = 0  # wdFindStop：查到文档末尾即停，不折回开头
            occ = 0
            while f.Execute():
                occ += 1
                spec = occ_map.get(occ)
                if spec is not None:
                    if spec.get("next_paragraph"):
                        _replace_next_paragraph(doc, rng, str(spec["new_text"]))
                    else:
                        rng.Text = str(spec["new_text"])
                # 无论是否替换均折叠到命中区间末尾，从其后继续查找
                rng.Collapse(0)  # wdCollapseEnd
                f = rng.Find
                f.ClearFormatting()
                f.Text = anchor
                f.Forward = True
                f.Wrap = 0

    @staticmethod
    def _cell_text(cell) -> str:
        """取单元格自身文本(不含结束符)。

        WPS COM 的 Cell(1,1).Range 会把紧邻表格上方的正文段落(如表前的
        「共安排日计划X项」概况句)一并并入文本,表现为「段文本 + \\r + 单元格
        首行」,且单元格自身文本常以段落标记 \\r 结尾。此函数先剥掉末尾的
        \\r、切除 \\x07 结束符,再取最后一个 \\r 之后的部分,只留下单元格真正
        的内容,避免表头比较/写入时误伤表前段落。
        """
        try:
            txt = cell.Range.Text
        except Exception:
            return ""
        x07 = txt.rfind("\x07")
        if x07 >= 0:
            txt = txt[:x07]
        while txt.endswith("\r"):
            txt = txt[:-1]
        if "\r" in txt:
            txt = txt.split("\r")[-1]
        return txt.strip()

    @staticmethod
    def _table_headers(tbl) -> List[str]:
        """读取表格首行表头文本列表(仅单元格自身内容)。"""
        heads: List[str] = []
        for ci in range(1, tbl.Columns.Count + 1):
            heads.append(WordSession._cell_text(tbl.Cell(1, ci)))
        return heads

    @staticmethod
    def _match_score(heads: List[str], columns: List[str]) -> int:
        """表头与目标列名的顺序匹配分数（用于在多个附图中定位目标表）。"""
        cset = set(columns)
        score = 0
        max_len = min(len(heads), len(columns))
        for i in range(max_len):
            if heads[i] and heads[i] in cset:
                score += 2
        # 顺序连续加分：目的性更强
        run = 0
        prev = -10
        for i, h in enumerate(heads):
            if h in cset:
                if i == prev + 1:
                    run += 1
                    score += run
                else:
                    run = 1
                prev = i
        # 列数完全一致加权（周报矩阵表列头是动态日期，同月周「7日」等与模板
        # 跨月占位「8月31日」不重叠，仅靠表头重叠会与明细表打成平手误选目标表）
        if len(heads) == len(columns):
            score += 20
        return score

    def _fill_table(self, doc, spec: Dict[str, Any]) -> None:
        """按列定位附表并填充数据行。

        规则：保留表头行，删除其余全部行（模板中的示例/占位行），
        再按数据行数追加并逐格填入；行序与 spec.columns 一致。
        """
        columns = [str(c) for c in spec.get("columns") or []]
        rows = spec.get("rows") or []
        if not columns:
            return
        target = None
        anchor = spec.get("anchor")
        if anchor:
            # 转置矩阵表列头为分类名(单位)，与模板列头零重叠、表头匹配会误选
            # 其他表，改按锚点句段定位目标表。
            # 注意：锚点段（紧跟表格上方的正文）经 _apply_replace 整段 Range.Text
            # 改写后，WPS/Word 会把该段并入紧邻表格首格（表格 Range 前移吞掉段落），
            # 此时 Find 命中区间落在表格内部——该表正是锚点所述的目标表；否则
            # 按「命中区间之后的第一个表格」定位。
            fnd = doc.Content.Find
            fnd.ClearFormatting()
            fnd.Text = str(anchor)
            fnd.Forward = True
            fnd.Wrap = 0  # wdFindStop
            if fnd.Execute():
                hs, he = fnd.Parent.Start, fnd.Parent.End
                for ti in range(1, doc.Tables.Count + 1):
                    tbl = doc.Tables(ti)
                    if tbl.Range.Start <= hs and he <= tbl.Range.End:
                        target = tbl
                        break
                if target is None:
                    for ti in range(1, doc.Tables.Count + 1):
                        tbl = doc.Tables(ti)
                        if tbl.Range.Start >= he and (
                                target is None or tbl.Range.Start < target.Range.Start):
                            target = tbl
        if target is None:
            best = -1
            for ti in range(1, doc.Tables.Count + 1):
                tbl = doc.Tables(ti)
                heads = self._table_headers(tbl)
                score = self._match_score(heads, columns)
                if score > best:
                    best = score
                    target = tbl
        if target is None:
            raise InputError(
                f"未找到匹配附表（列：{columns}），请核对 --params tables.columns")

        if not rows:
            # 无数据行（如该周期无三级及以上风险作业）：整表删除（含表头），
            # 避免空明细表仅剩表头行占用版面。概况句「…明细表如下：」收尾的
            # 改写由 genReportParams 的 replace 负责，二者配套整体移除该区块。
            # 概况句文本先读出备查；正文段改用段内文本重写（保留段落标记，不再
            # 把概况句并入紧邻表格）后，删表不再吞掉概况句，故删表后按「无需列示
            # 明细。」是否仍命中判断：还在则不重建（避免重复段），否则按「各作业
            # 类型日计划数量如下图所示」锚点重建概况句段。
            try:
                overview = None
                fnd = doc.Content.Find
                fnd.ClearFormatting()
                fnd.Text = "无需列示明细。"
                fnd.Forward = True
                fnd.Wrap = 0  # wdFindStop
                if fnd.Execute():
                    tail = fnd.Parent.End
                    for pi in range(1, doc.Paragraphs.Count + 1):
                        rng = doc.Paragraphs(pi).Range
                        if rng.Start <= tail <= rng.End:
                            overview = doc.Range(rng.Start, tail).Text
                            break
                target.Range.Select()
                self.app.Selection.Rows.Delete()
                if overview:
                    still = doc.Content.Find
                    still.ClearFormatting()
                    still.Text = "无需列示明细。"
                    still.Forward = True
                    still.Wrap = 0  # wdFindStop
                    if not still.Execute():
                        res = doc.Content.Find
                        res.ClearFormatting()
                        res.Text = "各作业类型日计划数量如下图所示"
                        res.Forward = True
                        res.Wrap = 0
                        if res.Execute():
                            rng = res.Parent
                            rng.Collapse(1)  # wdCollapseStart：锚点段首
                            rng.InsertBefore(overview.replace("\r", "") + "\r")
            except Exception:
                pass
            return

        # v2.6：矩阵型表（spec 带 max_units_per_table 标记）单位列数超阈值时
        # 自动纵切多张上下堆叠子表（每块 ≤ 阈值单位列、首列行键重复），防列数
        # 过多撑爆版心；明细表等无标记 spec 走单表填充。
        max_units = spec.get("max_units_per_table")
        if max_units is not None and len(columns) - 1 > int(max_units):
            self._fill_matrix_chunked(doc, target, columns, rows, int(max_units))
        else:
            self._fill_table_data(doc, target, columns, rows)

    def _fill_table_data(self, doc, target, columns: List[str],
                         rows: List[List[Any]]) -> None:
        """填充目标表：表头覆盖写 → 列数增减对齐 → 删示例行 → 加数据行 → 逐格填。

        单表与矩阵分块（_fill_matrix_chunked 的每块）共用；列数由调用方决定
        （spec.columns 或本块列）。
        """
        # 表头按 spec.columns 覆盖写（仅周报矩阵表列头是动态日期；明细表列名与
        # 模板一致时跳过，零写入）。
        # 注意：给表格首行单元格 Range.Text 赋新值时,Cell(1,1).Range 会把紧邻
        # 表格上方的正文段落（如「共安排日计划X项」概况句）一并并入而被覆盖删除,
        # 故仅当目标表头确与 spec 不一致才写（此时目标必为矩阵表非首列、未并入
        # 段落，安全）；范围收缩到 r.End-1 排除单元格结束符 \x07，避免结构漂移。
        # 表头判等依赖 _cell_text 剥离并入段落与末尾 \r，否则模板自带表头会被
        # 误判为不匹配而触发写入、覆盖表前概况句。
        existing = self._table_headers(target)
        # 参数列比模板表格少（如周报矩阵表去掉周六/周日列后按工作日列渲染）时，
        # 删除末尾多余列，避免残留空占位列（周六/周日列头+空单元格）。
        while target.Columns.Count > len(columns):
            try:
                target.Columns(target.Columns.Count).Delete()
                existing.pop()
            except Exception:
                break
        # 参数列比模板表格多（转置矩阵表列头为分类名、单位数可能超过模板列数）
        # 时补足列，否则表头忠实写入被截断、数据列缺失。
        while target.Columns.Count < len(columns):
            try:
                target.Columns.Add()
                existing.append("")
            except Exception:
                break
        for c, want in enumerate(columns, start=1):
            got = existing[c - 1] if c - 1 < len(existing) else ""
            if got == want:
                continue
            try:
                cell = target.Cell(1, c)
                rng = cell.Range
                # 首格顶部可能被并入紧邻上段的正文（锚点/概况句，以 \r 分隔），
                # 仅覆盖「最后一个 \r 之后」的真实单元格文本、保留并入前缀，
                # 避免改写时把表前段落删除。
                raw = rng.Text
                if "\x07" in raw:
                    raw = raw[: raw.rfind("\x07")]
                raw = raw.rstrip("\r")
                seg = raw.rfind("\r")
                real_len = len(raw) - seg - 1 if seg >= 0 else len(raw)
                if raw[-real_len:] == got and real_len < len(raw):
                    rng.Text = raw[: len(raw) - real_len] + str(want)
                else:
                    rng.End = rng.End - 1  # 排除单元格结束符 \x07
                    rng.Text = str(want)
            except Exception:
                pass

        # 删除表头以下所有示例行
        total_rows = target.Rows.Count
        for r in range(total_rows, 1, -1):
            target.Rows(r).Delete()

        header_cols = target.Columns.Count
        for _ in range(len(rows)):
            target.Rows.Add()
        for offset, data in enumerate(rows, start=2):
            for c, value in enumerate(data, start=1):
                if c > header_cols:
                    break
                try:
                    target.Cell(offset, c).Range.Text = str(value)
                except Exception:
                    pass
        # v2.9 明细表排版修复（仅首列「序号」的明细表生效，矩阵/汇总表不动）：
        # ① 首行跨页重复表头；② 单元格垂直居中；
        # ③ 列宽按版心分配（周报 7 列原先合计 556pt 超 A4 版心≈442pt，必须按比例收进版心）。
        if columns and str(columns[0]).strip() == "序号":
            try:
                target.Rows(1).HeadingFormat = True
            except Exception:
                pass
            try:
                for r in range(1, target.Rows.Count + 1):
                    for c in range(1, target.Columns.Count + 1):
                        try:
                            target.Cell(r, c).VerticalAlignment = 1  # wdCellAlignVerticalCenter
                        except Exception:
                            pass
            except Exception:
                pass
            try:
                ps = target.Range.Document.PageSetup
                text_w = float(ps.PageWidth - ps.LeftMargin - ps.RightMargin)
            except Exception:
                text_w = 442.0
            prefs = WordSession._detail_col_prefs(columns, text_w)
            for c, width in enumerate(prefs, start=1):
                try:
                    target.Columns(c).PreferredWidthType = 3  # wdPreferredWidthPoints
                    target.Columns(c).PreferredWidth = width
                    target.Columns(c).Width = width
                except Exception:
                    try:
                        target.Columns(c).Width = width
                    except Exception:
                        pass
            try:
                target.AllowAutoFit = False
            except Exception:
                pass
            try:
                target.PreferredWidthType = 3
                target.PreferredWidth = min(text_w, sum(prefs))
            except Exception:
                pass
            try:
                target.Rows.Alignment = 1  # 居中
            except Exception:
                pass

    def _fill_matrix_chunked(self, doc, target, columns: List[str],
                             rows: List[List[Any]], max_units: int) -> None:
        """矩阵表纵切多块上下堆叠（每块 ≤ max_units 个单位列，首列行键重复）。

        块 0 复用模板既有表；块 1..n 用 doc.Tables.Add 接在前一张表之后插入新表
        （表间以「（续表）」正文段分隔，防 WPS 保存时合并相邻表），复制模板表
        边框/字体/列宽保持同款样式。纯数据驱动（由 len(columns) 决定），
        日/周/月任何模板的矩阵表都自动适用；收窄由 applyCharts 的
        fit_tables_to_text 统一兜底。
        """
        key_col = columns[0]
        units = columns[1:]
        n = len(units)
        k = -(-n // max_units)  # 最少块数（每块 ≤ max_units 单位列）
        base, rem = divmod(n, k)  # 平衡分块：如 17 单位 → 6/6/5
        blocks = []
        idx = 0
        for i in range(k):
            s = base + (1 if i < rem else 0)
            blocks.append(units[idx:idx + s])
            idx += s

        start = 0
        prev = target
        ref_tbl = None
        widest = max(len(b) for b in blocks)  # 列宽重排按最宽块单位数，各块单位列等宽
        for bi, blk in enumerate(blocks):
            blk_cols = [key_col] + blk
            blk_rows = [[r[0]] + r[1 + start:1 + start + len(blk)] for r in rows]
            if bi == 0:
                tbl = target
                self._fill_table_data(doc, tbl, blk_cols, blk_rows)
                self._compact_matrix_table(tbl)
                ref_tbl = tbl  # 填好的模板表 = 新表字体归一参照源
                self._rebalance_matrix_widths(doc, tbl, widest)
            else:
                tbl = self._append_matrix_table(doc, prev, target,
                                                nrows=len(rows) + 1,
                                                ncols=len(blk_cols))
                self._fill_table_data(doc, tbl, blk_cols, blk_rows)
                self._compact_matrix_table(tbl, ref_tbl)
                self._rebalance_matrix_widths(doc, tbl, widest)
            LOGGER.info("矩阵表分块 %d/%d：单位列 %d 个（%s）",
                        bi + 1, k, len(blk), "、".join(blk))
            prev = tbl
            start += len(blk)

    def _compact_matrix_table(self, tbl, ref=None, font_size: float = 12.0) -> None:
        """矩阵块行高统一 + 新表单元格文本格式归一。

        - 行自动高度 + 单倍行距：块 0 复用模板既有表（带模板固定高行
          57.8pt/行）；Tables.Add 新表继承正文固定 29pt 行距（行高约 31pt）。
          三块上下堆叠时行高不一致、总高膨胀。统一后行高随字号自动、观感一致。
        - 新表单元格写值继承正文三号（16pt）字体与 2 字符首行缩进——WPS 按
          **字符单位**存储（CharacterUnitFirstLineIndent=2），按点赋
          FirstLineIndent=0 不生效，须按字符单位清。49pt 宽单元格里一行放
          不下两个汉字，值被逐字竖排换行（「16」拆「1」「6」、「市北中心」
          拆 3 行），行高撑到约 63pt，周报 4 页被撑到 6 页。归一：清缩进
          （字符+点两种单位）、居中，字体逐格赋值——新表表级 Range.Font
          读回返回 9999999 哨兵且表级赋值无效，只有逐格 Font 赋值生效；
          参照取模板表第 2 列同位格（首列可能并入正文段、格式混杂）。
        - v2.6.1：字号统一 12pt（模板 10.5pt 太挤看不清）。12pt 下 4 字单位名
          48pt+边距 10.8≈58.8pt < 重排后单位列宽 63.7pt，表头单行不折行；
          字号对全部块逐格无条件写入（含块 0，模板表同样放大）。
        """
        try:
            for r in range(1, tbl.Rows.Count + 1):
                tbl.Rows(r).HeightRule = 0  # wdRowHeightAuto
        except Exception:
            pass
        try:
            for p in tbl.Range.Paragraphs:
                pf = p.Format
                pf.LineSpacingRule = 0  # 单倍行距
                try:
                    pf.CharacterUnitFirstLineIndent = 0  # 字符单位首行缩进
                except Exception:
                    pass
                pf.FirstLineIndent = 0
                try:
                    pf.CharacterUnitLeftIndent = 0
                except Exception:
                    pass
                pf.LeftIndent = 0
                pf.RightIndent = 0
                pf.Alignment = 1  # 居中，与模板表一致
        except Exception:
            pass
        for r in range(1, tbl.Rows.Count + 1):
            sf = None
            if ref is not None and r <= ref.Rows.Count:
                try:
                    sf = ref.Cell(r, 2).Range.Font
                except Exception:
                    sf = None
            for c in range(1, tbl.Columns.Count + 1):
                try:
                    df = tbl.Cell(r, c).Range.Font
                    df.Size = font_size  # v2.6.1：无条件 12pt
                    if sf is not None:
                        if sf.Name:
                            df.Name = sf.Name
                        if sf.NameFarEast:
                            df.NameFarEast = sf.NameFarEast
                        if sf.Bold in (0, -1, True, False):
                            df.Bold = sf.Bold
                except Exception:
                    pass

    def _rebalance_matrix_widths(self, doc, tbl, max_units: int) -> None:
        """v2.6.1：矩阵块列宽按版心重排——首列（日期）60pt，单位列均分余量。

        模板表原列宽（首列 100pt/单位列 49pt）是为旧 10.5pt 字设计的；12pt
        下 4 字单位名需 ~58.8pt 才能单行显示，49pt 会折行。按版心重算：
        首列固定 60pt（12pt「8月31日」≈42pt+边距，放得下），单位列宽 =
        (版心宽-60)/最宽块单位数（17 单位→3 块 6/6/5，(442.2-60)/6≈63.7pt），
        各块单位列等宽观感一致。AllowAutoFit=False 锁宽防保存重置。
        applyCharts.fit_tables_to_text 幂等兜底：重排后总宽≤版心，不再动。
        """
        try:
            ps = doc.PageSetup
            text_w = ps.PageWidth - ps.LeftMargin - ps.RightMargin
            w_first = 60.0
            w_unit = (text_w - w_first) / max(1, max_units)
            tbl.Columns(1).Width = w_first
            for c in range(2, tbl.Columns.Count + 1):
                tbl.Columns(c).Width = w_unit
            tbl.AllowAutoFit = False
        except Exception:
            pass

    def _append_matrix_table(self, doc, prev, tbl0, nrows: int,
                             ncols: int):
        """在 prev 表后插入新子表，两表间以「（续表）」正文段分隔。

        WPS 干跑验证过的坑：
        - 在 Table.Range.Start 处 InsertBefore 实际会写进新表**首格**，且两张
          裸相邻表在保存时会被合并 —— 「（续表）」段必须用 prev 表尾
          InsertAfter("（续表）\\r") 生成独立正文段，再在其后 Tables.Add；
        - 新表列宽读取返回 9999999 哨兵值，不可链式复制 —— 列宽快照取自
          原始模板表 tbl0（首列宽 + 单位列宽）；
        - 设定列宽后须 AllowAutoFit=False 锁定，否则保存后宽度被重置。
        """
        mark = "（续表）"
        P = prev.Range.End
        doc.Range(P, P).InsertAfter(mark + "\r")
        pos = P + len(mark) + 1
        new_tbl = doc.Tables.Add(doc.Range(pos, pos), nrows, ncols)
        # 边框 8 条 + 字体：从模板表复制，保持同款样式
        for bi in range(1, 9):
            try:
                s, d = tbl0.Borders(bi), new_tbl.Borders(bi)
                d.LineStyle = s.LineStyle
                d.LineWidth = s.LineWidth
                d.Color = s.Color
            except Exception:
                pass
        try:
            fs, fd = tbl0.Range.Font, new_tbl.Range.Font
            fd.Name = fs.Name
            fd.NameFarEast = fs.NameFarEast
            fd.Size = fs.Size
        except Exception:
            fs = None
        # 列宽：首列=模板表首列宽，其余=模板表第 2 列宽（快照自原表，不读新表）
        try:
            w_first = tbl0.Columns(1).Width
            w_unit = tbl0.Columns(2).Width if tbl0.Columns.Count >= 2 else w_first
            new_tbl.Columns(1).Width = w_first
            for c in range(2, ncols + 1):
                new_tbl.Columns(c).Width = w_unit
            new_tbl.AllowAutoFit = False
        except Exception:
            pass
        # 「（续表）」段：居中 + 与表同款字体
        try:
            p = doc.Range(P, P + len(mark)).Paragraphs(1)
            p.Alignment = 1
            if fs is not None:
                p.Range.Font.Name = fs.Name
                p.Range.Font.NameFarEast = fs.NameFarEast
                p.Range.Font.Size = fs.Size
        except Exception:
            pass
        return new_tbl

    def close(self) -> None:
        if not getattr(self, "_owns_app", True):
            return
        try:
            self.app.Quit()
        except Exception:
            pass


# =============================================================================
# 控制器（CLI 入口）
# =============================================================================

def _demo_params() -> Dict[str, Any]:
    """内置示例参数（用于 --demo 快速验证模板链路）。"""
    return {
        "replace": {
            "20××年××月××日至××月××日/20××年第×周（20××年××月××日至××月××日）": "2026年9月1日至9月5日（2026年第36周）",
            "20××年××月××日": "2026年9月3日",
            "20××年第××号": "2026年第36号",
            "××公司××工区工作负责人×××连续工作超××天": "青岛公司变电检修工区工作负责人王**连续工作超5天",
            "××公司××工区": "青岛公司变电检修工区",
            "××公司××中心": "青岛公司配电运检中心",
            "××月××日": "9月5日",
            "××单位": "变电检修中心",
            "2026年X月X日": "2026年9月3日",
            "X月份第X周（X月X日至X日）": "9月份第3周（9月1日至6日）",
            "X月X日至X日（X天）": "9月1日至6日（6天）",
            "X月份": "9月份",
            "X月X日": "9月3日",
            "共安排日计划X项": "共安排日计划12项",
            "二级风险作业X项": "二级风险作业1项",
            "三级风险作业X项": "三级风险作业3项",
            "X1": "变电检修中心",
            "X2": "输电中心",
            "XX1": "市南中心",
            "XX2": "电缆中心",
            "XXXXX等工作": "日常工作",
        },
        "tables": [
            {
                "columns": ["序号", "单位", "作业内容", "风险等级", "高风险作业时间"],
                "rows": [
                    ["1", "变电检修中心", "#1主变综合检修", "三级", "9月3日-5日（3天）"],
                    ["2", "输电中心", "220kV线路消缺", "二级", "9月3日-4日（2天）"],
                ],
            },
            {
                "columns": ["日期", "变电运维", "变电检修"],
                "rows": [
                    ["1日", 40, 55],
                    ["2日", 30, 120],
                    ["3日", 45, 35],
                ],
            },
        ],
    }


def build_main_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="fillReport.py",
        description="承载力量化报告 / 风险函件填充引擎（技能「安全作业承载力分析」）",
        add_help=True,
    )
    parser.add_argument("--template", dest="template", default=None,
                        help=f"模板名：{' / '.join(TEMPLATES)}")
    parser.add_argument("--params", dest="params_path", default=None, metavar="PATH",
                        help="填充参数 JSON 文件（replace + tables）")
    parser.add_argument("--out", dest="out_path", default=None, metavar="PATH",
                        help="成品输出路径（.doc 或 .docx）")
    parser.add_argument("--list", action="store_true",
                        help="列出全部可用模板")
    parser.add_argument("--demo", action="store_true",
                        help="用内置示例参数验证模板链路")
    return parser


def _setup_logging() -> None:
    logging.basicConfig(level=logging.INFO, stream=sys.stderr,
                        format="[%(levelname)s] %(message)s")


def main(argv: Optional[List[str]] = None) -> int:
    args = build_main_parser().parse_args(argv)
    _setup_logging()

    if args.list:
        print(json.dumps(TEMPLATES, ensure_ascii=False, indent=2))
        return EXIT_OK

    if not args.template:
        print(__doc__)
        return EXIT_USAGE
    if not args.out_path:
        raise SystemExit(f"--out 必填；模板：{' / '.join(TEMPLATES)}")

    template_name = TEMPLATES.get(args.template)
    if template_name is None:
        print(f"未知模板：{args.template}，可用：{' / '.join(TEMPLATES)}",
              file=sys.stderr)
        return EXIT_USAGE

    if args.demo:
        params: Dict[str, Any] = _demo_params()
    else:
        if not args.params_path:
            print(__doc__)
            return EXIT_USAGE
        if not os.path.exists(args.params_path):
            LOGGER.error("参数文件不存在：%s", args.params_path)
            print(json.dumps({"ok": False, "errorCode": ERR_FILE_NOT_FOUND,
                              "message": f"参数文件不存在：{args.params_path}"},
                             ensure_ascii=False))
            return EXIT_IO
        try:
            with open(args.params_path, "r", encoding="utf-8") as fh:
                params = json.load(fh)
        except (json.JSONDecodeError, OSError) as exc:
            LOGGER.error("参数文件解析失败：%s（%s）", args.params_path, exc)
            print(json.dumps({"ok": False, "errorCode": ERR_FILE_INVALID,
                              "message": f"参数文件解析失败：{exc}"},
                             ensure_ascii=False))
            return EXIT_IO
        if not isinstance(params, dict):
            LOGGER.error("参数 JSON 顶层必须是对象")
            return EXIT_INPUT

    template_path = os.path.join(TEMPLATES_DIR, template_name)
    if not os.path.exists(template_path):
        LOGGER.error("模板文件缺失：%s", template_path)
        print(json.dumps({"ok": False, "errorCode": ERR_FILE_NOT_FOUND,
                          "message": f"模板文件缺失：{template_path}"},
                         ensure_ascii=False))
        return EXIT_IO

    out_path = os.path.abspath(args.out_path)
    low = out_path.lower()
    out_as_doc = low.endswith(".doc") and not low.endswith(".docx")

    # 出口校验：WPS SaveAs 对非法文件名字符报 3011（保存失败），重试也救不回，
    # 且难排查。非法字符全角化应由上游完成，此处只负责把问题点提前暴露成明确报错。
    bad_fn = [ch for ch in os.path.basename(out_path) if ch in _ILLEGAL_FN]
    if bad_fn:
        LOGGER.error("输出文件名含 Windows 非法字符 %s（%s）：%s",
                     bad_fn, args.out_path,
                     "请让上游 genLetterParams._safe_filename 做全角化转换")
        print(json.dumps({"ok": False, "errorCode": ERR_OUTPUT_WRITE_FAILED,
                          "message": f"输出文件名含非法字符 {bad_fn}：{args.out_path}"},
                         ensure_ascii=False))
        return EXIT_IO

    session = WordSession()
    try:
        session.render(template_path, params, out_path, out_as_doc)
    except FillError as exc:
        LOGGER.error("%s: %s (context=%s)", exc.code, exc.message, exc.context)
        print(json.dumps(exc.to_payload(), ensure_ascii=False))
        return _ERROR_EXIT.get(exc.code, EXIT_UNEXPECTED)
    except Exception as exc:  # noqa: BLE001 最后一层兜底
        LOGGER.exception("未预期异常")
        print(json.dumps({"ok": False, "errorCode": ERR_UNEXPECTED,
                          "message": f"未预期异常：{exc}"}, ensure_ascii=False))
        return EXIT_UNEXPECTED
    finally:
        try:
            session.close()
        except Exception:
            pass

    print(json.dumps({"ok": True, "template": args.template, "out": out_path},
                     ensure_ascii=False))
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
