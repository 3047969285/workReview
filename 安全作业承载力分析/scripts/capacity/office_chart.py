# -*- coding: utf-8 -*-
"""
把柱状图以 Word DrawingML 写入 .docx（可双击编辑，不经 PNG、不经 WPS AddChart）。

成品是 OOXML 标准图表：内嵌一份 xlsx 数据簿 + chart.xml。Word 与 WPS 在内网
均可打开并改数据；生成端只依赖已安装的 openpyxl 与标准库 zipfile。

统一竖柱（国网汇报风）：横轴=类别（单位/类型），纵轴数值（刻度自带 %，无「项/%」轴标题）。
图内宋体小四；数据标签整数；分类轴横排不倾斜；无外框无图例。
正文 29pt 固定行距与模板字体不在本模块改写。
月报弹性分页：图1跟计划表后空白、图2+图3续排、（二）+第1周同页、第2～5周各一页。
引导句与图：空段压扁、段间距清零，贴紧。
"""
from __future__ import annotations

import io
import os
import re
import tempfile
import zipfile
from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence, Tuple
from xml.sax.saxutils import escape

from openpyxl import Workbook

EMU_PER_PT = 12700
CHART_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart")
PKG_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package")
CHART_CT = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
XLSX_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# 图内字体：宋体小四（DrawingML sz 为百分之一磅，12pt = 1200）
CHART_FONT = "宋体"
CHART_SZ = "1200"

_PARA_RE = re.compile(r"<w:p(?:\s[^>]*)?>[\s\S]*?</w:p>")
_PPR_RE = re.compile(r"<w:pPr(?:\s[^>]*)?>[\s\S]*?</w:pPr>")
_WT_RE = re.compile(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>")
_DRAW_RE = re.compile(
    r"(?:<mc:AlternateContent[\s\S]*?</mc:AlternateContent>"
    r"|<w:drawing[\s\S]*?</w:drawing>"
    r"|<w:pict[\s\S]*?</w:pict>)")
_RID_RE = re.compile(r'Id="rId(\d+)"')
_DOCPR_RE = re.compile(r'<wp:docPr[^>]*\bid="(\d+)"')
_GUIDE_RE = re.compile(r"(?:如图|如下图|如下表|如表)所示")
_TITLE_RE = re.compile(
    r"^(?:[一二三四五六七八九十]+、"
    r"|（[^）]*[0-9一二三四五六七八九十][^）]*）"
    r"|[0-9]+[.、．])")


@dataclass
class OfficeChartSpec:
    """一份待写入报告的可编辑柱状图。"""

    caption: str
    labels: List[str]
    values: List[float]
    colors: List[str]
    y_title: str
    pct: bool
    horizontal: bool
    width_pt: float
    height_pt: float
    value_max: Optional[float] = None


def rgb_int_to_hex(n: int) -> str:
    """把 Excel ForeColor.RGB 长整型转成 RRGGBB。"""
    red = n & 0xFF
    green = (n >> 8) & 0xFF
    blue = (n >> 16) & 0xFF
    return f"{red:02X}{green:02X}{blue:02X}"


def list_captions(docx_path: str) -> List[str]:
    """按正文段落顺序取出「图 N …」图题（与月报模板图位一一对应）。"""
    with zipfile.ZipFile(docx_path, "r") as zf:
        xml = zf.read("word/document.xml").decode("utf-8")
    out: List[str] = []
    for para in _PARA_RE.findall(xml):
        text = _para_text(para)
        if _is_caption_text(text):
            out.append(text)
    return out


def inject_office_charts(docx_path: str, specs: Sequence[OfficeChartSpec],
                         out_path: str) -> Dict[str, int]:
    """把 specs 写入 docx 图位段，并绑定引导句+图+图题，返回 inserted/skipped。"""
    src = os.path.abspath(docx_path)
    dst = os.path.abspath(out_path)
    with zipfile.ZipFile(src, "r") as zin:
        parts = {name: zin.read(name) for name in zin.namelist()}
    doc_xml = parts["word/document.xml"].decode("utf-8")
    rels_name = "word/_rels/document.xml.rels"
    rels_xml = parts[rels_name].decode("utf-8")
    ct_xml = parts["[Content_Types].xml"].decode("utf-8")
    doc_xml = _ensure_drawing_ns(doc_xml)
    paras = _PARA_RE.findall(doc_xml)
    rid_n = _max_int(_RID_RE.findall(rels_xml))
    docpr_n = _max_int(_DOCPR_RE.findall(doc_xml))
    chart_n = _next_chart_index(parts)
    inserted = 0
    skipped = 0
    new_parts: Dict[str, bytes] = {}
    rel_adds: List[str] = []
    ct_adds: List[str] = []
    slots = _assign_slots(paras, specs)
    used_figs = set()
    used_specs = set()

    def _add_chart_parts(spec: OfficeChartSpec) -> Tuple[str, int]:
        nonlocal chart_n, rid_n, docpr_n
        chart_n += 1
        rid_n += 1
        docpr_n += 1
        rid = f"rId{rid_n}"
        chart_part = f"word/charts/chart{chart_n}.xml"
        embed_part = f"word/embeddings/Microsoft_Excel_Worksheet{chart_n}.xlsx"
        rels_part = f"word/charts/_rels/chart{chart_n}.xml.rels"
        new_parts[chart_part] = _chart_xml(spec).encode("utf-8")
        new_parts[embed_part] = _embedding_xlsx(spec)
        new_parts[rels_part] = _chart_rels_xml(chart_n).encode("utf-8")
        rel_adds.append(
            f'<Relationship Id="{rid}" Type="{CHART_REL_TYPE}" '
            f'Target="charts/chart{chart_n}.xml"/>')
        ct_adds.append(
            f'<Override PartName="/{chart_part}" ContentType="{CHART_CT}"/>')
        ct_adds.append(
            f'<Override PartName="/{embed_part}" ContentType="{XLSX_CT}"/>')
        return rid, docpr_n

    for fig_idx, cap_idx, spec in slots:
        if fig_idx in used_figs or not spec.labels or not spec.values:
            skipped += 1
            continue
        used_figs.add(fig_idx)
        used_specs.add(id(spec))
        rid, docpr_id = _add_chart_parts(spec)
        drawing = _inline_drawing(rid, spec, docpr_id)
        paras[fig_idx] = _replace_figure(paras[fig_idx], drawing)
        _bind_chart_block(paras, fig_idx, cap_idx)
        inserted += 1
    new_doc = _PARA_RE.sub(lambda m, it=iter(paras): next(it), doc_xml, count=len(paras))
    for spec in specs:
        if id(spec) in used_specs:
            continue
        if not spec.labels or not spec.values:
            skipped += 1
            continue
        guide = _guide_for_spec(spec)
        if not any(guide in _para_text(p) for p in _PARA_RE.findall(new_doc)):
            skipped += 1
            continue
        rid, docpr_id = _add_chart_parts(spec)
        drawing = _inline_drawing(rid, spec, docpr_id)
        new_para = _new_figure_para(drawing)
        new_doc, ok = _splice_after_text(new_doc, guide, new_para)
        if ok:
            inserted += 1
            used_specs.add(id(spec))
        else:
            skipped += 1
    rels_xml = _insert_before(rels_xml, "</Relationships>", "".join(rel_adds))
    ct_xml = _ensure_xlsx_default(ct_xml)
    ct_xml = _insert_before(ct_xml, "</Types>", "".join(ct_adds))
    parts["word/document.xml"] = new_doc.encode("utf-8")
    parts[rels_name] = rels_xml.encode("utf-8")
    parts["[Content_Types].xml"] = ct_xml.encode("utf-8")
    parts.update(new_parts)
    _write_zip(parts, dst if dst != src else src)
    return {"inserted": inserted, "skipped": skipped}


def reflow_month_type_chart(docx_path: str) -> Dict[str, float]:
    """月报版面回流入口：OOXML 清粘连后交给本模块 polish_month_docx。

    版面策略集中在本文件后半「版面策略」区；改观感只改该区常量/函数。
    """
    path = os.path.abspath(docx_path)
    with zipfile.ZipFile(path, "r") as zin:
        parts = {name: zin.read(name) for name in zin.namelist()}
    doc_xml = parts["word/document.xml"].decode("utf-8")
    paras = _PARA_RE.findall(doc_xml)
    guide_i = None
    for i, para in enumerate(paras):
        if "各作业类型日计划数量如下图所示" in _para_text(para):
            guide_i = i
            break
    if guide_i is None:
        return {"ok": 0.0}
    fig_i = None
    cap_i = None
    for j in range(guide_i + 1, min(guide_i + 8, len(paras))):
        p = paras[j]
        t = _para_text(p)
        if fig_i is None and (
                _DRAW_RE.search(p) or "w:drawing" in p or "w:object" in p
                or "OLEObject" in p):
            fig_i = j
        if t.startswith("图") and "作业类型" in t:
            cap_i = j
            break
    if fig_i is None or cap_i is None:
        return {"ok": 0.0, "guide": float(guide_i),
                "fig": float(fig_i or -1), "cap": float(cap_i or -1)}

    # 注入初值：取上限，COM 按剩余空间下调
    target_h = H_TYPE_MAX
    for i in range(guide_i, cap_i + 1):
        is_guide = (i == guide_i)
        is_fig = (i == fig_i)
        is_empty = (not _para_text(paras[i]) and not is_fig)
        paras[i] = _patch_layout_ppr(
            paras[i],
            keep_next=False,
            pbb_off=True,
            single_spacing=is_fig or is_empty,
            clear_first_indent=True,
            center=is_fig,
            keep_lines=is_fig,
            zero_gap=is_fig or is_guide or is_empty,
            crush_empty=is_empty,
            guide_tight=is_guide)

    cy = int(round(target_h * EMU_PER_PT))
    fig = paras[fig_i]
    fig2, n_ext = re.subn(
        r'(<wp:extent\b[^>]*\bcy=")(\d+)(")',
        lambda m: f"{m.group(1)}{cy}{m.group(3)}",
        fig, count=1)
    if n_ext:
        fig2 = re.sub(
            r'(<a:ext\b[^>]*\bcy=")(\d+)(")',
            lambda m: f"{m.group(1)}{cy}{m.group(3)}",
            fig2, count=1)
        paras[fig_i] = fig2

    new_doc = _PARA_RE.sub(lambda m, it=iter(paras): next(it), doc_xml, count=len(paras))
    parts["word/document.xml"] = new_doc.encode("utf-8")
    _write_zip(parts, path)

    try:
        com_info = polish_month_docx(path)
    except Exception as exc:  # noqa: BLE001
        com_info = {"com_ok": 0.0, "error": 0.0}
        # 保留可诊断字段但不抛死流水线
        _ = str(exc)
    return {
        "ok": 1.0,
        "guide": float(guide_i),
        "fig": float(fig_i),
        "cap": float(cap_i),
        "height_pt": float(com_info.get("type_h", target_h)),
        "extent_patched": float(n_ext),
        **{k: float(v) for k, v in com_info.items() if isinstance(v, (int, float))},
    }


def apply_month_fixed_pages(docx_path: str) -> Dict[str, int]:
    """月报弹性分页（按实页观感）：

    1) 图1（作业类型数量）紧接计划表后；「二、公司各单位」换页，
       图2+图3同页填满，避免图3独页留白
    2) 「（二）周…」换页 → 与第1周同页（第1周不再另起页）
    3) 第2～5周标题换页 → 一周两图一页
    4) 辅助建议换页
    不改正文 29pt/字体。
    """
    path = os.path.abspath(docx_path)
    with zipfile.ZipFile(path, "r") as zin:
        parts = {name: zin.read(name) for name in zin.namelist()}
    doc_xml = parts["word/document.xml"].decode("utf-8")
    paras = _PARA_RE.findall(doc_xml)
    week_re = re.compile(r"^([1-5])\..*第[1-5]周")
    overview_keys = ("各作业类型日计划数量如下图所示",)
    unit_section_re = re.compile(r"^二、公司各单位")
    week_section_re = re.compile(r"^(?:[（(]二[）)]|二、).{0,8}周")
    advice_keys = ("四、辅助建议", "三、辅助建议")
    n_overview_cleared = 0
    n_unit = 0
    n_week = 0
    n_week1_cleared = 0
    n_week_section = 0
    n_advice = 0
    overview_done = False
    unit_done = False
    week_section_done = False
    for i, para in enumerate(paras):
        text = _para_text(para)
        if not text:
            continue
        wm = week_re.match(text)
        if wm:
            # 第1周跟「（二）」同页，清硬分页；第2～5周换页
            if wm.group(1) == "1":
                paras[i] = _clear_page_break_before(para)
                n_week1_cleared += 1
            else:
                paras[i] = _force_page_break_before(para)
                n_week += 1
            continue
        # 图1引导句：清掉硬分页，并连带清掉后续图段/图注的段前分页
        if not overview_done and any(k in text for k in overview_keys):
            paras[i] = _clear_page_break_before(para)
            overview_done = True
            n_overview_cleared += 1
            for j in range(i + 1, min(i + 6, len(paras))):
                paras[j] = _clear_page_break_before(paras[j])
                jt = _para_text(paras[j])
                if jt.startswith("图") and "作业类型" in jt:
                    break
            continue
        # 「二、公司各单位」硬换页：图1跟风险表；图2+图3独占次页，避免图3独大留白
        if not unit_done and unit_section_re.match(text):
            paras[i] = _force_page_break_before(para)
            unit_done = True
            n_unit += 1
            continue
        force = False
        if not week_section_done and week_section_re.match(text):
            force = True
            week_section_done = True
            n_week_section += 1
        elif any(text.startswith(k) for k in advice_keys):
            force = True
            n_advice += 1
        if force:
            paras[i] = _force_page_break_before(para)
    new_doc = _PARA_RE.sub(lambda m, it=iter(paras): next(it), doc_xml, count=len(paras))
    parts["word/document.xml"] = new_doc.encode("utf-8")
    _write_zip(parts, path)
    return {
        "overview_cleared": n_overview_cleared,
        "unit_section": n_unit,
        "week_section": n_week_section,
        "weeks": n_week,
        "week1_cleared": n_week1_cleared,
        "advice": n_advice,
    }


def _force_page_break_before(para: str) -> str:
    """段前分页：去掉旧 pageBreakBefore 后写入开启值。"""
    para = _ensure_ppr(para)
    m = _PPR_RE.search(para)
    if not m:
        return para
    ppr = _strip_ppr_child(m.group(0), "pageBreakBefore")
    ppr = re.sub(
        r"<w:pPr(?:\s[^>]*)?>",
        lambda mm: mm.group(0) + "<w:pageBreakBefore/>",
        ppr, count=1)
    return para[:m.start()] + ppr + para[m.end():]


def _clear_page_break_before(para: str) -> str:
    """去掉段前分页（含空元素与 val=1），避免残留硬分页。"""
    if "pageBreakBefore" not in para:
        return para
    para = _ensure_ppr(para)
    m = _PPR_RE.search(para)
    if not m:
        return para
    ppr = _strip_ppr_child(m.group(0), "pageBreakBefore")
    return para[:m.start()] + ppr + para[m.end():]


def _max_int(found: List[str]) -> int:
    return max((int(x) for x in found), default=0)


def _next_chart_index(parts: Dict[str, bytes]) -> int:
    nums = []
    for name in parts:
        m = re.search(r"word/charts/chart(\d+)\.xml$", name)
        if m:
            nums.append(int(m.group(1)))
    return max(nums) if nums else 0


def _para_text(para: str) -> str:
    return "".join(_WT_RE.findall(para)).replace("\u00a0", " ").strip()


def _is_caption_text(text: str) -> bool:
    return text.startswith("图 ") and len(text) <= 40


def _is_guide_text(text: str) -> bool:
    return bool(text) and (
        bool(_GUIDE_RE.search(text)) or "管理承载力为" in text)


def _is_short_title(text: str) -> bool:
    return bool(text) and bool(_TITLE_RE.match(text)) and len(text) <= 24


def _caption_indexes(paras: List[str]) -> List[Tuple[int, str]]:
    out: List[Tuple[int, str]] = []
    for i, para in enumerate(paras):
        text = _para_text(para)
        if _is_caption_text(text):
            out.append((i, text))
    return out


def _assign_slots(
        paras: List[str],
        specs: Sequence[OfficeChartSpec]
        ) -> List[Tuple[int, Optional[int], OfficeChartSpec]]:
    """图位配对：有「图 N」图题时按图题；否则按文档中图框出现顺序。"""
    caption_at = _caption_indexes(paras)
    draw_at = [i for i, p in enumerate(paras) if _DRAW_RE.search(p)]
    by_cap = {s.caption: s for s in specs if _is_caption_text(s.caption or "")}
    if by_cap and caption_at:
        used = set()
        pairs: List[Tuple[int, Optional[int], OfficeChartSpec]] = []
        for cap_idx, caption in caption_at:
            spec = by_cap.get(caption)
            if spec is None or caption in used or cap_idx <= 0:
                continue
            used.add(caption)
            pairs.append((cap_idx - 1, cap_idx, spec))
        return pairs
    cap_after = {cap_idx - 1: cap_idx for cap_idx, _ in caption_at if cap_idx > 0}
    return [(fig, cap_after.get(fig), spec)
            for fig, spec in zip(draw_at, specs)]


def _xml_block_start(paras: List[str], fig_idx: int) -> int:
    """图组起点：最多回溯短标题 + 引导句，中间空段一并纳入。"""
    start = fig_idx
    taken = 0
    j = fig_idx - 1
    while j >= 0 and taken < 2:
        text = _para_text(paras[j])
        if not text and not _DRAW_RE.search(paras[j]):
            start = j
            j -= 1
            continue
        if _is_guide_text(text) or _is_short_title(text):
            start = j
            taken += 1
            j -= 1
            continue
        break
    return start


def _xml_block_end(paras: List[str], fig_idx: int,
                   cap_idx: Optional[int]) -> int:
    if cap_idx is not None:
        return cap_idx
    for k in range(fig_idx + 1, min(fig_idx + 4, len(paras))):
        if _is_caption_text(_para_text(paras[k])):
            return k
    return fig_idx


def _bind_chart_block(paras: List[str], fig_idx: int,
                      cap_idx: Optional[int]) -> None:
    """引导句+图+图题整组粘死：空段压到近零，段前段后距清零，禁止拆页。

    不改正文 29pt 字号；引导句保留 29pt 固定行距，只去掉段间距，与图贴紧。
    """
    start = _xml_block_start(paras, fig_idx)
    end = _xml_block_end(paras, fig_idx, cap_idx)
    for i in range(start, end + 1):
        text = _para_text(paras[i])
        is_fig = (i == fig_idx)
        is_cap = (i == end and end > fig_idx)
        is_empty = (not text and not is_fig)
        is_guide = (not is_fig and not is_cap and not is_empty)
        paras[i] = _patch_layout_ppr(
            paras[i],
            keep_next=(not is_cap),
            pbb_off=True,
            single_spacing=is_fig or is_empty,
            clear_first_indent=is_fig or is_empty or is_guide,
            center=is_fig,
            keep_lines=True,
            zero_gap=is_fig or is_guide or is_empty,
            crush_empty=is_empty,
            guide_tight=is_guide)


def _strip_ppr_child(ppr: str, tag: str) -> str:
    ppr = re.sub(rf"<w:{tag}\b[^>]*/>", "", ppr)
    ppr = re.sub(rf"<w:{tag}\b[^>]*>[\s\S]*?</w:{tag}>", "", ppr)
    return ppr


def _ensure_ppr(para: str) -> str:
    if _PPR_RE.search(para):
        return para
    return re.sub(
        r"<w:p(?:\s[^>]*)?>",
        lambda m: m.group(0) + "<w:pPr></w:pPr>",
        para, count=1)


def _patch_layout_ppr(para: str, *, keep_next: Optional[bool],
                      pbb_off: bool, single_spacing: bool,
                      clear_first_indent: bool, center: bool,
                      keep_lines: bool = True,
                      zero_gap: bool = False,
                      crush_empty: bool = False,
                      guide_tight: bool = False) -> str:
    """只改 pPr 布局子节点，保留原 rPr（模板字体/字号）。"""
    para = _ensure_ppr(para)
    m = _PPR_RE.search(para)
    if not m:
        return para
    ppr = m.group(0)
    strip = ["keepNext", "pageBreakBefore", "widowControl", "keepLines"]
    if single_spacing or zero_gap or crush_empty or guide_tight:
        strip.append("spacing")
    if clear_first_indent or crush_empty:
        strip.append("ind")
    if center:
        strip.append("jc")
    for tag in strip:
        ppr = _strip_ppr_child(ppr, tag)
    blob = ""
    if keep_next:
        blob += "<w:keepNext/>"
    elif keep_next is False:
        blob += '<w:keepNext w:val="0"/>'
    if keep_lines:
        blob += "<w:keepLines/>"
    blob += '<w:widowControl w:val="0"/>'
    if pbb_off:
        blob += '<w:pageBreakBefore w:val="0"/>'
    if crush_empty:
        # 近零行高：空段几乎不占版面
        blob += ('<w:spacing w:before="0" w:after="0" '
                 'w:line="1" w:lineRule="exact"/>')
    elif guide_tight:
        # 引导句：保留约 29pt 固定行距，段前段后距清零，贴住下方图
        blob += ('<w:spacing w:before="0" w:after="0" '
                 'w:line="580" w:lineRule="exact"/>')
    elif zero_gap:
        blob += ('<w:spacing w:before="0" w:after="0" '
                 'w:line="240" w:lineRule="auto"/>')
    elif single_spacing:
        blob += '<w:spacing w:line="240" w:lineRule="auto"/>'
    if clear_first_indent or crush_empty:
        blob += '<w:ind w:firstLine="0" w:firstLineChars="0"/>'
    if center:
        blob += '<w:jc w:val="center"/>'
    ppr = re.sub(
        r"<w:pPr(?:\s[^>]*)?>",
        lambda mm: mm.group(0) + blob,
        ppr, count=1)
    return para[:m.start()] + ppr + para[m.end():]


def _ensure_drawing_ns(xml: str) -> str:
    """图 drawing 需要 wp/a 命名空间；WPS 另存后根节点有时缺声明。"""
    m = re.search(r"<w:document\b[^>]*>", xml)
    if not m:
        return xml
    tag = m.group(0)
    extra = ""
    if "xmlns:wp=" not in tag:
        extra += ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
    if "xmlns:a=" not in tag:
        extra += ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    if not extra:
        return xml
    return xml[:m.end() - 1] + extra + xml[m.end() - 1:]


def _ensure_xlsx_default(ct_xml: str) -> str:
    if 'Extension="xlsx"' in ct_xml:
        return ct_xml
    return _insert_before(
        ct_xml, "</Types>",
        f'<Default Extension="xlsx" ContentType="{XLSX_CT}"/>')


def _insert_before(xml: str, mark: str, blob: str) -> str:
    if not blob:
        return xml
    idx = xml.rfind(mark)
    if idx < 0:
        raise RuntimeError(f"docx 部件缺少 {mark}")
    return xml[:idx] + blob + xml[idx:]


def _guide_for_spec(spec: OfficeChartSpec) -> str:
    t = (spec.y_title or "") + (spec.caption or "")
    if "项" in t or "作业类型" in t:
        return "各作业类型日计划数量如下图所示"
    if "管理" in t:
        return "管理承载力为"
    return "作业承载力如图所示"


def _new_figure_para(drawing: str) -> str:
    return (
        "<w:p><w:pPr><w:keepNext/><w:keepLines/>"
        '<w:pageBreakBefore w:val="0"/>'
        '<w:spacing w:line="240" w:lineRule="auto"/>'
        '<w:ind w:firstLine="0" w:firstLineChars="0"/>'
        '<w:jc w:val="center"/></w:pPr>'
        f"<w:r>{drawing}</w:r></w:p>"
    )


def _splice_after_text(doc_xml: str, needle: str, new_para: str) -> Tuple[str, bool]:
    if not needle:
        return doc_xml, False
    for para in _PARA_RE.findall(doc_xml):
        if needle not in _para_text(para):
            continue
        pos = doc_xml.find(para)
        if pos < 0:
            continue
        end = pos + len(para)
        return doc_xml[:end] + new_para + doc_xml[end:], True
    return doc_xml, False


def _replace_figure(para: str, drawing: str) -> str:
    if _DRAW_RE.search(para):
        return _DRAW_RE.sub(drawing, para, count=1)
    if para.endswith("</w:p>"):
        return para[:-6] + f"<w:r>{drawing}</w:r></w:p>"
    return para + f"<w:r>{drawing}</w:r>"


def _inline_drawing(rid: str, spec: OfficeChartSpec, docpr_id: int) -> str:
    cx = int(round(spec.width_pt * EMU_PER_PT))
    cy = int(round(spec.height_pt * EMU_PER_PT))
    return (
        f'<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{docpr_id}" name="Chart {docpr_id}"/>'
        f'<wp:cNvGraphicFramePr>'
        f'<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        f' noChangeAspect="0"/>'
        f'</wp:cNvGraphicFramePr>'
        f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        f'<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
        f' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        f' r:id="{rid}"/>'
        f'</a:graphicData></a:graphic></wp:inline></w:drawing>'
    )


def _chart_rels_xml(chart_n: int) -> str:
    target = f"../embeddings/Microsoft_Excel_Worksheet{chart_n}.xlsx"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{PKG_REL_TYPE}" Target="{target}"/>'
        "</Relationships>"
    )


def _embedding_xlsx(spec: OfficeChartSpec) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.cell(1, 1, "类别")
    ws.cell(1, 2, "%" if spec.pct else "项")
    labels = list(spec.labels)
    values = list(spec.values)
    # 横向条才反转；国网竖柱保持传入顺序（峰值在左）
    if spec.horizontal:
        labels = list(reversed(labels))
        values = list(reversed(values))
    for i, (lab, val) in enumerate(zip(labels, values), start=2):
        # 分类轴换行：单元格内真换行，供图绑定缓存使用
        ws.cell(i, 1, str(lab).replace("\\n", "\n"))
        ws.cell(i, 2, int(round(float(val))))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _ea_font(sz: str = CHART_SZ) -> str:
    return (
        f'<a:defRPr sz="{sz}">'
        f'<a:latin typeface="{CHART_FONT}"/>'
        f'<a:ea typeface="{CHART_FONT}"/>'
        f'<a:cs typeface="{CHART_FONT}"/>'
        f"</a:defRPr>"
    )


def _tx_pr(rot: int = 0, sz: str = CHART_SZ) -> str:
    """图内文字：默认不旋转；分类多时靠截断字号，避免 Word 自动倾斜撑高。"""
    rot_attr = f' rot="{int(rot)}"' if rot else ' rot="0"'
    return (
        f"<c:txPr><a:bodyPr{rot_attr} wrap=\"square\"/>"
        f"<a:lstStyle/><a:p><a:pPr>{_ea_font(sz)}</a:pPr>"
        f'<a:endParaRPr lang="zh-CN" altLang="en-US" sz="{sz}">'
        f'<a:latin typeface="{CHART_FONT}"/>'
        f'<a:ea typeface="{CHART_FONT}"/>'
        f"</a:endParaRPr></a:p></c:txPr>"
    )


def _wrap_cat_label(lab: str, n_cats: int) -> str:
    """分类轴标签：横排、不换行；类多时适度截断（与承载力图观感接近）。"""
    s = str(lab).replace("\n", "").replace("\\n", "").strip()
    if n_cats >= 16:
        return s[:3]
    if n_cats >= 12:
        return s[:4]
    return s


def _chart_xml(spec: OfficeChartSpec) -> str:
    """国网竖柱：横轴=类别，纵轴数值；无轴标题、无外框、无图例；标签整数。

    刻度格式自带 %（百分图）；不再写「项/%」轴标题——Word 会把左轴标题拧成倾斜竖字。
    """
    labels = list(spec.labels)
    values = [float(v) for v in spec.values]
    colors = list(spec.colors)
    if spec.horizontal:
        labels = list(reversed(labels))
        values = list(reversed(values))
        colors = list(reversed(colors))
    n = len(labels)
    last = n + 1
    bar_dir = "bar" if spec.horizontal else "col"
    vmax = spec.value_max if spec.value_max else (125.0 if spec.pct else None)
    lbl_fmt = "0"
    # 百分图刻度带 %；计数图只写整数——单位不进轴标题
    ax_fmt = '0&quot;%&quot;' if spec.pct else "0"
    gap = 80 if not spec.horizontal else 45
    if not spec.horizontal and n >= 14:
        gap = 120
    elif not spec.horizontal and n >= 8:
        gap = 100
    cat_sz = "800" if n >= 12 else ("900" if n >= 8 else CHART_SZ)
    # 数量图与承载力图一致：柱顶外置标签
    dlbl_pos = "outEnd"
    dlbl_sz = "900" if (not spec.pct and n >= 10) else CHART_SZ
    dpts = []
    pts_cat = []
    pts_val = []
    for i, (lab, val) in enumerate(zip(labels, values)):
        color = colors[i] if i < len(colors) else "1457A8"
        color = color.lstrip("#")
        iv = int(round(val))
        show_lab = lab if spec.horizontal else _wrap_cat_label(lab, n)
        dpts.append(
            f'<c:dPt><c:idx val="{i}"/><c:bubble3D val="0"/><c:spPr>'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:ln><a:noFill/></a:ln></c:spPr></c:dPt>')
        pts_cat.append(
            f'<c:pt idx="{i}"><c:v>{escape(show_lab)}</c:v></c:pt>')
        pts_val.append(f'<c:pt idx="{i}"><c:v>{iv}</c:v></c:pt>')
    scaling = '<c:scaling><c:orientation val="minMax"/>'
    if vmax is not None:
        scaling += f'<c:max val="{float(vmax)}"/><c:min val="0"/>'
    scaling += "</c:scaling>"
    major = '<c:majorUnit val="25"/>' if spec.pct else ""
    cat_pos = "l" if spec.horizontal else "b"
    val_pos = "b" if spec.horizontal else "l"
    no_box = (
        '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"{no_box}"
        '<c:chart><c:autoTitleDeleted val="1"/>'
        f'<c:plotArea>{no_box}<c:layout/>'
        f'<c:barChart><c:barDir val="{bar_dir}"/><c:grouping val="clustered"/>'
        '<c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/>'
        "<c:tx><c:v></c:v></c:tx>"
        f"{''.join(dpts)}"
        "<c:dLbls>"
        f"{_tx_pr(0, dlbl_sz)}"
        f'<c:dLblPos val="{dlbl_pos}"/>'
        "<c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/>"
        "<c:showCatName val=\"0\"/><c:showSerName val=\"0\"/>"
        "<c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/>"
        f'<c:numFmt formatCode="{lbl_fmt}" sourceLinked="0"/>'
        "</c:dLbls>"
        f'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A${last}</c:f>'
        f'<c:strCache><c:ptCount val="{n}"/>{"".join(pts_cat)}</c:strCache>'
        "</c:strRef></c:cat>"
        f'<c:val><c:numRef><c:f>Sheet1!$B$2:$B${last}</c:f>'
        f'<c:numCache><c:formatCode>0</c:formatCode><c:ptCount val="{n}"/>'
        f'{"".join(pts_val)}</c:numCache></c:numRef></c:val>'
        f'</c:ser><c:gapWidth val="{gap}"/><c:overlap val="0"/>'
        '<c:axId val="1"/><c:axId val="2"/>'
        "</c:barChart>"
        '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/>'
        "</c:scaling><c:delete val=\"0\"/>"
        f'<c:axPos val="{cat_pos}"/>'
        "<c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/>"
        "<c:tickLblPos val=\"nextTo\"/>"
        f"{_tx_pr(0, cat_sz)}"
        "<c:crossAx val=\"2\"/>"
        "<c:crosses val=\"autoZero\"/></c:catAx>"
        f'<c:valAx><c:axId val="2"/>{scaling}<c:delete val="0"/>'
        f'<c:axPos val="{val_pos}"/>'
        "<c:majorGridlines>"
        '<c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="D8DEE8"/>'
        "</a:solidFill></a:ln></c:spPr></c:majorGridlines>"
        # 无轴标题：避免「项/%」被 Word 拧成倾斜竖字
        f'<c:numFmt formatCode="{ax_fmt}" sourceLinked="0"/>'
        f"{_tx_pr(0)}"
        "<c:majorTickMark val=\"out\"/><c:minorTickMark val=\"none\"/>"
        f"{major}<c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"1\"/>"
        "<c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/>"
        "</c:valAx></c:plotArea><c:plotVisOnly val=\"1\"/></c:chart>"
        '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>'
        "</c:chartSpace>"
    )


def _write_zip(parts: Dict[str, bytes], dest: str) -> None:
    parent = os.path.dirname(os.path.abspath(dest))
    os.makedirs(parent, exist_ok=True)
    fd, tmp = tempfile.mkstemp(suffix=".docx", dir=parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, data in parts.items():
                zf.writestr(name, data)
        os.replace(tmp, dest)
    except OSError:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise

# =====================================================================
# 版面策略（原 report_layout.py；改观感只改本区）
# =====================================================================

# ---------- 版心与图宽（与公文模板一致）----------
CHART_W_PT = 435.0
TEXT_WIDTH_FALLBACK_PT = 442.0

# ---------- 明细表（序号表）----------
# 相对正文 29pt 略紧，给表后数量图腾位；禁止再写死 RowHeight Exact（会裁字/撑页）
DETAIL_FONT_PT = 10.5
DETAIL_LINE_PT = 16.0  # wdLineSpaceExactly
DETAIL_COL_PREF = {
    "序号": 28.0,
    "风险等级": 32.0,
    "作业内容": 150.0,
    "时间": 72.0,
    "日期": 72.0,
    "班组成员": 40.0,
    "人数": 40.0,
    "工作负责人": 48.0,
    "负责人": 48.0,
    "单位": 48.0,
}
DETAIL_COL_DEFAULT = 52.0

# ---------- 图高策略：只给区间，运行时按剩余空间取值 ----------
# 全月作业/管理（图2+图3 同页）
H_MONTH_MIN, H_MONTH_MAX = 110.0, 130.0
# 每周两图同页
H_WEEK_MIN, H_WEEK_MAX = 130.0, 148.0
# 作业类型数量图：优先跟表末同页，否则独页放大
H_TYPE_MIN, H_TYPE_MAX = 110.0, 200.0
# 引导句+图注预留
TYPE_TEXT_RESERVE_PT = 48.0
# 周页标题/评述预留（两图分母）
WEEK_TEXT_RESERVE_PT = 200.0

_WEEK_TITLE_RE = re.compile(r"^[1-5]\.\d+月份第[1-5]周")


def detail_col_widths(headers: List[str], text_w: float) -> List[float]:
    """明细表列宽：语义配方后等比拉满版心（居中后左右不露白）。"""
    prefs: List[float] = []
    for h in headers:
        width = DETAIL_COL_DEFAULT
        for key, w in DETAIL_COL_PREF.items():
            if key in h:
                width = w
                break
        prefs.append(width)
    total = sum(prefs) or 1.0
    if abs(total - text_w) > 0.5:
        prefs = [w * text_w / total for w in prefs]
    return prefs


def chart_height_pt(n_cats: int, *, kind: str) -> float:
    """注入阶段初值（DrawingML extent）；最终以 polish_month_docx 实测为准。

    kind: type | week | month
    """
    n = max(int(n_cats), 1)
    if kind == "type":
        return min(H_TYPE_MAX, max(H_TYPE_MIN, 150.0))
    if kind == "week":
        return min(H_WEEK_MAX, max(H_WEEK_MIN, 132.0 + 0.4 * n))
    return min(H_MONTH_MAX, max(H_MONTH_MIN, 112.0 + 0.4 * n))


def _usable_height_pt(doc) -> float:
    ps = doc.PageSetup
    return float(ps.PageHeight - ps.TopMargin - ps.BottomMargin)


def _clamp(v: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, v))


def _para_plain(p) -> str:
    return p.Range.Text.replace("\r", "").replace("\x07", "").strip()


def _find_type_block(doc) -> Tuple[Optional[int], Optional[int], Optional[int]]:
    g_i = f_i = c_i = None
    for i in range(1, doc.Paragraphs.Count + 1):
        txt = _para_plain(doc.Paragraphs(i))
        if g_i is None and "各作业类型日计划数量如下图所示" in txt:
            g_i = i
        if g_i is not None and f_i is None and doc.Paragraphs(i).Range.InlineShapes.Count:
            f_i = i
        if g_i is not None and txt.startswith("图") and "作业类型" in txt:
            c_i = i
            break
    return g_i, f_i, c_i


def _table_end_page(doc) -> int:
    try:
        tbl = doc.Tables(1)
        return int(tbl.Rows(tbl.Rows.Count).Range.Information(3))
    except Exception:
        return -1


def _fit_height(
    doc, sh, g_i: int, f_i: int, c_i: int, *,
    hi: float, lo: float, want_tbl_page: int = -1,
) -> Tuple[float, str]:
    """从高到低试高：优先三件同页（且可选=表末页）；否则引导句+图同页。"""
    best_h = lo
    best_mode = "fig"
    step = 5
    h = float(hi)
    while h >= lo - 0.1:
        sh.Height = h
        sh.Width = CHART_W_PT
        doc.Paragraphs(g_i).Range.ParagraphFormat.KeepWithNext = True
        doc.Paragraphs(f_i).Range.ParagraphFormat.KeepWithNext = True
        doc.Repaginate()
        gp = int(doc.Paragraphs(g_i).Range.Information(3))
        fp = int(doc.Paragraphs(f_i).Range.Information(3))
        cp = int(doc.Paragraphs(c_i).Range.Information(3))
        if gp == fp == cp and (want_tbl_page < 0 or gp == want_tbl_page):
            return h, "all"
        doc.Paragraphs(f_i).Range.ParagraphFormat.KeepWithNext = False
        doc.Repaginate()
        gp = int(doc.Paragraphs(g_i).Range.Information(3))
        fp = int(doc.Paragraphs(f_i).Range.Information(3))
        if gp == fp and (want_tbl_page < 0 or gp == want_tbl_page):
            best_h = h
            best_mode = "fig"
            return best_h, best_mode
        h -= step
    # 放宽：不再要求跟表同页
    if want_tbl_page >= 0:
        return _fit_height(doc, sh, g_i, f_i, c_i, hi=hi, lo=lo, want_tbl_page=-1)
    return best_h, best_mode


def _center_all_charts(doc) -> None:
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        try:
            if p.Range.InlineShapes.Count >= 1:
                p.Range.ParagraphFormat.Alignment = 1
                sh = p.Range.InlineShapes(1)
                sh.LockAspectRatio = 0
                sh.Width = CHART_W_PT
            if _para_plain(p).startswith("图"):
                p.Range.ParagraphFormat.Alignment = 1
        except Exception:
            continue


def _purge_blank_before_weeks(doc) -> None:
    """删掉「第N周」前空段 + 清空段 PageBreakBefore，消灭空白页。"""
    for i in range(doc.Paragraphs.Count, 1, -1):
        t = _para_plain(doc.Paragraphs(i))
        if not _WEEK_TITLE_RE.match(t):
            continue
        while i > 1:
            prev = doc.Paragraphs(i - 1)
            if _para_plain(prev) or prev.Range.InlineShapes.Count:
                break
            prev.Range.Delete()
    for i in range(doc.Paragraphs.Count, 1, -1):
        p = doc.Paragraphs(i)
        if _para_plain(p):
            continue
        try:
            p.Range.ParagraphFormat.PageBreakBefore = False
        except Exception:
            pass


def _size_week_charts(doc) -> float:
    """按版心动态算周图高，保证同页两张不拆页。"""
    usable = _usable_height_pt(doc)
    target = _clamp((usable - WEEK_TEXT_RESERVE_PT) / 2.0, H_WEEK_MIN, H_WEEK_MAX)
    n = 0.0
    for i in range(1, doc.Paragraphs.Count + 1):
        txt = _para_plain(doc.Paragraphs(i))
        if not (txt.startswith("图") and "第" in txt and "周" in txt):
            continue
        if i <= 1:
            continue
        prev = doc.Paragraphs(i - 1)
        if prev.Range.InlineShapes.Count < 1:
            continue
        try:
            wsh = prev.Range.InlineShapes(1)
            wsh.LockAspectRatio = 0
            wsh.Width = CHART_W_PT
            wsh.Height = target
            n += 1.0
        except Exception:
            continue
    doc.Repaginate()
    # 同页≥2 张且仍拆页时略降
    by_page: Dict[int, List] = {}
    for si in range(1, doc.InlineShapes.Count + 1):
        try:
            sh = doc.InlineShapes(si)
            if float(sh.Height) < H_WEEK_MIN - 1:
                continue
            pg = int(sh.Range.Information(3))
            by_page.setdefault(pg, []).append(sh)
        except Exception:
            continue
    for shapes in by_page.values():
        if len(shapes) == 1 and float(shapes[0].Height) > H_WEEK_MIN + 5:
            try:
                shapes[0].Height = H_WEEK_MIN
            except Exception:
                pass
    return n


def polish_month_docx(docx_path: str) -> Dict[str, float]:
    """月报 COM 抛光：表归一 → 数量图按剩余空间 → 周图动态高 → 消空白页。

    每次 runReports 月报 native 注入后自动调用；勿为单次观感改调用方。
    优先复用 fillReport 共用 Word（--all 时禁止再 DispatchEx 第二实例互抢）。
    """
    import os

    import pythoncom
    import win32com.client as win32

    path = os.path.abspath(docx_path)
    out: Dict[str, float] = {"com_ok": 0.0}
    owned = False
    app = None
    try:
        from fillReport import _close_open_documents, current_word_app
        shared = current_word_app()
    except Exception:
        shared = None
        _close_open_documents = None  # type: ignore
    if shared is not None:
        app = shared
        if _close_open_documents is not None:
            _close_open_documents(app)
    else:
        pythoncom.CoInitialize()
        app = win32.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        owned = True
    doc = app.Documents.Open(path, AddToRecentFiles=False)
    try:
        try:
            from fillReport import WordSession
            WordSession._normalize_report_tables(doc)
        except Exception:
            pass

        g_i, f_i, c_i = _find_type_block(doc)
        type_h = H_TYPE_MIN
        if g_i and f_i and c_i:
            for i in range(g_i, c_i + 1):
                pf = doc.Paragraphs(i).Range.ParagraphFormat
                pf.KeepWithNext = False
                pf.PageBreakBefore = False
                pf.SpaceBefore = 0
                pf.SpaceAfter = 0
            doc.Paragraphs(f_i).Range.ParagraphFormat.KeepTogether = True
            sh = doc.Paragraphs(f_i).Range.InlineShapes(1)
            sh.LockAspectRatio = 0
            tbl_page = _table_end_page(doc)
            usable = _usable_height_pt(doc)
            hi = _clamp(usable - TYPE_TEXT_RESERVE_PT, H_TYPE_MIN, H_TYPE_MAX)
            type_h, mode = _fit_height(
                doc, sh, g_i, f_i, c_i,
                hi=hi, lo=H_TYPE_MIN, want_tbl_page=tbl_page)
            sh.Height = type_h
            sh.Width = CHART_W_PT
            doc.Paragraphs(g_i).Range.ParagraphFormat.KeepWithNext = True
            doc.Paragraphs(f_i).Range.ParagraphFormat.KeepWithNext = (mode == "all")
            doc.Repaginate()
            gp = int(doc.Paragraphs(g_i).Range.Information(3))
            fp = int(doc.Paragraphs(f_i).Range.Information(3))
            if gp != fp:
                doc.Paragraphs(g_i).Range.ParagraphFormat.KeepWithNext = False
                doc.Paragraphs(f_i).Range.ParagraphFormat.KeepWithNext = False
                doc.Repaginate()
            out["page_guide"] = float(int(doc.Paragraphs(g_i).Range.Information(3)))
            out["page_fig"] = float(int(doc.Paragraphs(f_i).Range.Information(3)))
            out["page_cap"] = float(int(doc.Paragraphs(c_i).Range.Information(3)))
            out["tbl_page"] = float(tbl_page)
            out["type_mode"] = 1.0 if mode == "all" else 0.0

        week_n = _size_week_charts(doc)
        _purge_blank_before_weeks(doc)
        _center_all_charts(doc)

        g_i, f_i, c_i = _find_type_block(doc)
        if g_i and f_i and c_i:
            sh = doc.Paragraphs(f_i).Range.InlineShapes(1)
            sh.LockAspectRatio = 0
            tbl_page = _table_end_page(doc)
            usable = _usable_height_pt(doc)
            hi = _clamp(usable - TYPE_TEXT_RESERVE_PT, H_TYPE_MIN, H_TYPE_MAX)
            type_h, mode = _fit_height(
                doc, sh, g_i, f_i, c_i,
                hi=hi, lo=H_TYPE_MIN, want_tbl_page=tbl_page)
            sh.Height = type_h
            sh.Width = CHART_W_PT
            doc.Paragraphs(g_i).Range.ParagraphFormat.KeepWithNext = True
            doc.Paragraphs(f_i).Range.ParagraphFormat.KeepWithNext = (mode == "all")
            doc.Paragraphs(f_i).Range.ParagraphFormat.Alignment = 1
            doc.Repaginate()
            out["page_guide"] = float(int(doc.Paragraphs(g_i).Range.Information(3)))
            out["page_fig"] = float(int(doc.Paragraphs(f_i).Range.Information(3)))
            out["page_cap"] = float(int(doc.Paragraphs(c_i).Range.Information(3)))
            out["type_mode"] = 1.0 if mode == "all" else 0.0

        out["week_charts"] = week_n
        out["type_h"] = float(type_h)
        out["com_ok"] = 1.0
        doc.Save()
    finally:
        try:
            doc.Close(SaveChanges=True)
        except Exception:
            pass
        if owned and app is not None:
            try:
                app.Quit()
            except Exception:
                pass
    return out

# =====================================================================
# 日/周/月 native 编排入口（原 renderNativeCharts.py）
# =====================================================================


import argparse
import json
import logging
import math
import os
import shutil
import sys
from typing import Any, Dict, List, Optional, Tuple

LOGGER_NATIVE = logging.getLogger("office_chart.native")

EXIT_OK, EXIT_UNEXPECTED, EXIT_USAGE, EXIT_INPUT, EXIT_IO = 0, 1, 2, 3, 4

# 图宽/图高区间：唯一源 capacity.report_layout（注入初值；月报 COM 再按剩余空间定稿）

# 预警分级配色（与 renderReportCharts 完全一致：#1457A8 / #5A8FD4 / #E3483F / #B02418）
C_MAIN = "#1457A8"
C_HEAVY = "#5A8FD4"
C_FULL = "#E3483F"
C_OVER = "#B02418"

# Excel 图表常量
_XL_COLUMN_CLUSTERED = 51      # xlColumnClustered（纵向柱状，类型图类少时用）
_XL_BAR_CLUSTERED = 57         # xlBarClustered（横向条，单位/管理承载力）
_XL_VALUE = 2                  # xlValue（值轴）
_XL_AUTOMATIC = -4105          # 分类轴标签方向自动
_XL_LABEL_OUT_END = 2          # 数据标签在条/柱外端
# 数据提取复用（与 PNG 版同口径）
from pngCharts import load_matrix, unit_period_mean, plan_type_count  # noqa: E402

_MAX_CAPTION_LEN = 40          # 图题段长度上限（「图 N 公司…」含空格）
_NONE_UNIT = "其他"            # 无类型计划归类名（与 renderReportCharts 一致）


def wrap_axis_label(s: str, width: int = 6) -> str:
    """竖柱图长分类标签按字数折行（单元格内换行 \\n）。"""
    if len(s) <= width:
        return s
    return "\n".join(s[i:i + width] for i in range(0, len(s), width))


def _pct_axis_max(values: List[float]) -> float:
    """百分图值轴上限：默认 125，峰值超过则按 25 进位，作业/管理共用便于对比。"""
    peak = max(values) if values else 0.0
    if peak <= 125.0:
        return 125.0
    return math.ceil(peak / 25.0) * 25.0


def _vbar_height(n: int, is_pct: bool = True, week: bool = False) -> float:
    """竖柱高度（pt）：委托 report_layout.chart_height_pt；勿在此写死区间。"""
    if not is_pct:
        return chart_height_pt(n, kind="type")
    if week:
        return chart_height_pt(n, kind="week")
    return chart_height_pt(n, kind="month")


def _rgb(hex_color: str) -> int:
    """'#RRGGBB' → Excel ForeColor.RGB 长整型（注意字节序）。"""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r | (g << 8) | (b << 16)


def bar_color(v: float) -> int:
    """预警分级逐点配色（≤75 深蓝 · 75~90 浅蓝 · >90 红 · >100 深红）。"""
    if v > 100:
        return _rgb(C_OVER)
    if v > 90:
        return _rgb(C_FULL)
    if v > 75:
        return _rgb(C_HEAVY)
    return _rgb(C_MAIN)


def _ptype_colors(n: int) -> List[int]:
    """作业类型图配色：≤6 类统一深蓝，>6 类由深到浅蓝渐变（与 PNG 版同口径）。"""
    root = _rgb(C_MAIN)
    if n <= 6:
        return [root] * n
    base = (0x14, 0x57, 0xA8)   # C_MAIN
    light = (0x9D, 0xC3, 0xE6)  # 浅蓝端
    out = []
    for i in range(n):
        t = 0.90 - 0.55 * (i / (n - 1))
        r = int(round(base[0] + (light[0] - base[0]) * t))
        g = int(round(base[1] + (light[1] - base[1]) * t))
        b = int(round(base[2] + (light[2] - base[2]) * t))
        out.append(r | (g << 8) | (b << 16))
    return out


def manage_period_mean(data_path: str, lo: str, hi: str) -> List[Tuple[str, float]]:
    """管理承载力各单位周期值（与 renderReportCharts.render_manage_chart 同口径）。"""
    import datetime
    from capacity.calculator import period_manage_caps_from_days

    with open(data_path, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    days = (data.get("management") or {}).get("days") or {}
    caps = period_manage_caps_from_days(
        days, datetime.date.fromisoformat(lo), datetime.date.fromisoformat(hi))
    rows: List[Tuple[str, float]] = sorted(caps.items(), key=lambda kv: kv[1], reverse=True)
    return rows


def _month_caption(caption: str) -> bool:
    return "月份" in caption and "周" not in caption


def build_specs(caption: str, month_lo: str, month_hi: str,
                week_specs: List[Dict[str, str]]) -> Optional[Tuple[str, str]]:
    """图题 → 数据区间。返回 (lo, hi)；无法匹配返回 None。"""
    if _month_caption(caption):
        return (month_lo, month_hi)
    for k, w in ((1, "第1周"), (2, "第2周"), (3, "第3周"), (4, "第4周"), (5, "第5周")):
        if w in caption and k <= len(week_specs):
            sp = week_specs[k - 1]
            return (sp["lo"], sp["hi"])
    return None


def _make_spec(caption: str, labels: List[str], vals: List[float],
               colors: List[str], ytitle: str, is_pct: bool,
               week: bool = False
               ) -> OfficeChartSpec:
    """国网竖柱：横轴类别、纵轴数值；标签取整。"""
    round_vals = [float(int(round(v))) for v in vals]
    vmax = None
    if not is_pct and round_vals:
        peak = max(round_vals)
        vmax = float(math.ceil(peak * 1.2 / 500.0) * 500.0) if peak > 0 else None
    return OfficeChartSpec(
        caption=caption, labels=labels, values=round_vals, colors=colors,
        y_title=ytitle, pct=is_pct, horizontal=False,
        width_pt=CHART_W_PT, height_pt=_vbar_height(len(labels), is_pct, week),
        value_max=vmax)


def prepare_period_specs(captions: List[str], calc_json: str, source: str,
                         config_path: str, lo: str, hi: str
                         ) -> Tuple[List[OfficeChartSpec], int]:
    """日报/周报三图：作业承载力、作业类型、管理承载力（竖柱）。"""
    matrix = load_matrix(calc_json)
    skipped = 0
    job_rows = unit_period_mean(matrix, lo, hi)
    ptypes = plan_type_count(source, config_path, lo, hi)
    try:
        mgmt_rows = manage_period_mean(calc_json, lo, hi)
    except Exception as exc:  # noqa: BLE001
        LOGGER_NATIVE.warning("管理承载力取数失败，跳过该图：%s", exc)
        mgmt_rows = []
        skipped += 1

    def _job() -> OfficeChartSpec:
        labels = [str(u) for u, _ in job_rows]
        vals = [v for _, v in job_rows]
        colors = [rgb_int_to_hex(bar_color(v)) for v in vals]
        return _make_spec("", labels, vals, colors, "作业承载力", True)

    def _ptype() -> OfficeChartSpec:
        labels = [str(t) for t, _ in ptypes]
        vals = [float(v) for _, v in ptypes]
        colors = [rgb_int_to_hex(c) for c in _ptype_colors(len(labels))]
        return _make_spec("", labels, vals, colors, "计划项数", False)

    def _mgmt() -> Optional[OfficeChartSpec]:
        if not mgmt_rows:
            return None
        labels = [str(u) for u, _ in mgmt_rows]
        vals = [v for _, v in mgmt_rows]
        colors = [rgb_int_to_hex(bar_color(v)) for v in vals]
        return _make_spec("", labels, vals, colors, "管理承载力", True)

    ordered: List[OfficeChartSpec] = []
    if captions:
        for caption in captions:
            if "作业类型" in caption:
                spec = _ptype()
            elif "管理" in caption:
                spec = _mgmt()
                if spec is None:
                    skipped += 1
                    continue
            else:
                spec = _job()
            spec.caption = caption
            if spec.labels:
                ordered.append(spec)
            else:
                skipped += 1
    else:
        for spec in (_job(), _ptype(), _mgmt()):
            if spec is None or not spec.labels:
                skipped += 1
                continue
            ordered.append(spec)
    ymax = 125.0
    for spec in ordered:
        if spec.pct:
            ymax = max(ymax, _pct_axis_max(spec.values))
    for spec in ordered:
        if spec.pct:
            spec.value_max = ymax
    return ordered, skipped


def prepare_month_specs(captions: List[str], calc_json: str, source: str,
                        config_path: str, month_lo: str, month_hi: str,
                        week_specs: List[Dict[str, str]]
                        ) -> Tuple[List[OfficeChartSpec], int]:
    """按图题取数，组装可写入 .docx 的图表规格。"""
    matrix = load_matrix(calc_json)
    ymax = 125.0
    prepared: List[OfficeChartSpec] = []
    skipped = 0
    for caption in captions:
        try:
            if "作业类型" in caption:
                ptypes = plan_type_count(source, config_path, month_lo, month_hi)
                labels = [str(t) for t, _ in ptypes]
                vals = [float(v) for _, v in ptypes]
                colors = [rgb_int_to_hex(c) for c in _ptype_colors(len(labels))]
                ytitle = "计划项数"
                is_pct = False
            elif "管理承载力" in caption:
                ab = build_specs(caption, month_lo, month_hi, week_specs)
                if ab is None:
                    skipped += 1
                    continue
                rows = manage_period_mean(calc_json, ab[0], ab[1])
                labels = [str(u) for u, _ in rows]
                vals = [v for _, v in rows]
                colors = [rgb_int_to_hex(bar_color(v)) for v in vals]
                ytitle = "管理承载力"
                is_pct = True
            else:
                ab = build_specs(caption, month_lo, month_hi, week_specs)
                if ab is None:
                    skipped += 1
                    continue
                cap_rows = unit_period_mean(matrix, ab[0], ab[1])
                labels = [str(u) for u, _ in cap_rows]
                vals = [v for _, v in cap_rows]
                colors = [rgb_int_to_hex(bar_color(v)) for v in vals]
                ytitle = "作业承载力"
                is_pct = True
            if not labels:
                skipped += 1
                continue
            if is_pct:
                ymax = max(ymax, _pct_axis_max(vals))
            is_week = ("第" in caption and "周" in caption
                       and "作业类型" not in caption)
            prepared.append(_make_spec(
                caption, labels, vals, colors, ytitle, is_pct, week=is_week))
        except Exception as exc:  # noqa: BLE001
            LOGGER_NATIVE.warning("图位%r取数失败，跳过：%s", caption, exc)
            skipped += 1
    for spec in prepared:
        if spec.pct:
            spec.value_max = ymax
    return prepared, skipped


def _doc_to_docx(src: str, dest: str) -> None:
    """.doc 草稿另存为 .docx（仅转换，不插图）。"""
    import pythoncom
    import win32com.client as win32
    pythoncom.CoInitialize()
    app = None
    doc = None
    owned = False
    try:
        try:
            from fillReport import current_word_app
            shared = current_word_app()
        except Exception:
            shared = None
        if shared is not None:
            app = shared
            doc = app.Documents.Open(src, AddToRecentFiles=False)
        else:
            for prog in ("Word.Application", "Kwps.Application", "KWPS.Application"):
                cand = None
                try:
                    cand = win32.DispatchEx(prog)
                    cand.Visible = False
                    try:
                        cand.DisplayAlerts = 0
                    except Exception:
                        pass
                    doc = cand.Documents.Open(src)
                    app = cand
                    owned = True
                    break
                except Exception:
                    if cand is not None:
                        try:
                            cand.Quit()
                        except Exception:
                            pass
        if app is None or doc is None:
            raise RuntimeError("无法连接 Word/WPS，无法把 .doc 转为 .docx")
        doc.SaveAs2(dest, FileFormat=16, AddToRecentFiles=False)
        doc.Close(False)
        doc = None
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if owned and app is not None:
            try:
                app.Quit()
            except Exception:
                pass


def _ensure_docx(src: str, dest: str) -> None:
    src = os.path.abspath(src)
    dest = os.path.abspath(dest)
    if src.lower().endswith(".docx"):
        if src != dest:
            shutil.copy2(src, dest)
        return
    _doc_to_docx(src, dest)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="native_main",
        description="日/周/月 Word 原生可编辑横向条（纵轴类别、横轴数值，标签整数）")
    p.add_argument("--doc", required=True, help="fillReport 输出的正文草稿 .doc/.docx（含模板图位）")
    p.add_argument("--result", default=None, help="calcCapacity 输出 JSON（含 matrix/management）")
    p.add_argument("--source", default=None, help="作业计划源 .xls/.xlsx/.json")
    p.add_argument("--config", default=None, help="capacity_config.json 路径")
    p.add_argument("--lo", default=None, help="主周期起始日期（全月/范围起始）")
    p.add_argument("--hi", default=None, help="主周期结束日期")
    p.add_argument("--kind", choices=("day", "week", "month"), default="month",
                   help="报告形态：day/week 三图，month 按图题 13 图位")
    p.add_argument("--specs", default="[]",
                   help="周(spec) JSON：[{prefix,lo,hi},...]（月报第1~5周）")
    p.add_argument("--out", required=True, help="输出成品 .docx 路径（.doc 会把图栅格化）")
    return p


def _fail_exit(error_code: str, message: str, code: int) -> int:
    LOGGER_NATIVE.error(message)
    print(json.dumps({"ok": False, "errorCode": error_code, "message": message},
                     ensure_ascii=False))
    return code


def native_main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    args.doc = os.path.abspath(args.doc)
    args.out = os.path.abspath(args.out)
    for k in ("result", "source", "config"):
        if getattr(args, k):
            setattr(args, k, os.path.abspath(getattr(args, k)))
    if not os.path.exists(args.doc):
        return _fail_exit("IO", f"草稿文档不存在：{args.doc}", EXIT_IO)
    if not args.out.lower().endswith(".docx"):
        return _fail_exit("USAGE", "可编辑图表必须输出 .docx（.doc 会把图栅格化）", EXIT_USAGE)
    if not args.result or not args.source or not args.config or not args.lo or not args.hi:
        return _fail_exit("USAGE", "--result/--source/--config/--lo/--hi 必填", EXIT_USAGE)
    try:
        week_specs = json.loads(args.specs)
    except Exception as exc:
        return _fail_exit("USAGE", f"--specs 应为 JSON 数组：{exc}", EXIT_USAGE)
    if args.config:
        from capacity.config import apply_formula_config, load_config
        apply_formula_config(load_config(args.config))
    try:
        _ensure_docx(args.doc, args.out)
        captions = list_captions(args.out)
        if args.kind == "month":
            specs, skipped = prepare_month_specs(
                captions, args.result, args.source, args.config,
                args.lo, args.hi, week_specs)
        else:
            specs, skipped = prepare_period_specs(
                captions, args.result, args.source, args.config,
                args.lo, args.hi)
        stats = inject_office_charts(args.out, specs, args.out)
        stats["skipped"] = int(stats.get("skipped") or 0) + skipped
        page_info = {}
        if args.kind == "month":
            page_info = apply_month_fixed_pages(args.out)
            try:
                page_info["reflow_type"] = reflow_month_type_chart(args.out)
            except Exception as exc:  # noqa: BLE001
                LOGGER_NATIVE.warning("月报类型图回流跳过：%s", exc)
                page_info["reflow_type"] = {"ok": 0.0, "error": str(exc)}
        print(json.dumps({"ok": True, "kind": args.kind, "chart_mode": "native",
                          "inserted": stats["inserted"], "skipped": stats["skipped"],
                          "pages": page_info, "file": args.out},
                         ensure_ascii=False))
        return EXIT_OK
    except Exception as exc:  # noqa: BLE001
        LOGGER_NATIVE.exception("未预期异常")
        return _fail_exit("UNEXPECTED", f"未预期异常：{exc}", EXIT_UNEXPECTED)


