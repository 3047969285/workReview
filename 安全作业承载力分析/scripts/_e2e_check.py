# -*- coding: utf-8 -*-
"""E2E smoke checks for capacity report pipeline (no new product deps)."""
from __future__ import annotations

import json
import os
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

OUT_DIR = ROOT / "承载力报告"
SOURCE = ROOT / "data" / "工作计划-测试数据.xls"
ADVICE = OUT_DIR / "建议" / "advice_20260917.json"
DATE_FOLDER = OUT_DIR / "承载力报告20260917"


def _latest_date_folder() -> Path:
    folders = sorted(
        [p for p in OUT_DIR.iterdir() if p.is_dir() and p.name.startswith("承载力报告")],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not folders:
        raise SystemExit("no 承载力报告YYYYMMDD folder")
    return folders[0]


def check_docx_has_charts(path: Path, min_charts: int = 1) -> dict:
    """Count DrawingML chart parts inside docx."""
    n = 0
    with zipfile.ZipFile(path, "r") as z:
        for name in z.namelist():
            if name.startswith("word/charts/chart") and name.endswith(".xml"):
                n += 1
            # also drawing relationships
        if n == 0:
            for name in z.namelist():
                if "/charts/" in name:
                    n += 1
    return {"file": path.name, "charts": n, "ok": n >= min_charts, "size": path.stat().st_size}


def check_no_advice_placeholder(path: Path) -> dict:
    with zipfile.ZipFile(path, "r") as z:
        xml = z.read("word/document.xml").decode("utf-8", errors="ignore")
    bad = "ADVICE_BODY_2026" in xml or "建议占位符" in xml
    return {"file": path.name, "placeholder": bad, "ok": not bad}


def check_com_layout(path: Path) -> dict:
    """Light Word COM: pages, table center, inline shapes."""
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    app = win32.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    doc = app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False)
    try:
        doc.Repaginate()
        pages = int(doc.ComputeStatistics(2))  # wdStatisticPages
        shapes = int(doc.InlineShapes.Count)
        tbl_align = None
        if doc.Tables.Count >= 1:
            try:
                tbl_align = int(doc.Tables(1).Rows.Alignment)
            except Exception:
                tbl_align = -1
        return {
            "file": path.name,
            "pages": pages,
            "inline_shapes": shapes,
            "tbl_align": tbl_align,
            "ok": pages >= 1 and shapes >= 1,
        }
    finally:
        try:
            doc.Close(False)
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass


def main() -> int:
    folder = _latest_date_folder()
    print("FOLDER", folder)
    targets = {
        "日": (list(folder.glob("*9月1日*.docx")) or list(folder.glob("*日*安全*.docx"))),
        "周": (list(folder.glob("*第1周*.docx"))),
        "月": (list(folder.glob("*9月份安全*.docx")) or [
            p for p in folder.glob("*月份*.docx") if "第" not in p.name
        ]),
    }
    # refine month: exclude 第N周
    month = [p for p in folder.glob("*.docx") if "月份" in p.name and "第" not in p.name and "周" not in p.name]
    day = [p for p in folder.glob("*.docx") if "日" in p.name and "周" not in p.name and "月份" not in p.name]
    # prefer 9月1日
    day = [p for p in day if "9月1日" in p.name] or day
    week = [p for p in folder.glob("*.docx") if "第1周" in p.name]
    targets = {"日": day, "周": week, "月": month}

    report = {"folder": str(folder), "items": []}
    ok_all = True
    for kind, files in targets.items():
        if not files:
            report["items"].append({"kind": kind, "ok": False, "error": "missing docx"})
            ok_all = False
            continue
        path = max(files, key=lambda p: p.stat().st_mtime)
        min_c = 3 if kind == "月" else 1
        c1 = check_docx_has_charts(path, min_c)
        c2 = check_no_advice_placeholder(path)
        try:
            c3 = check_com_layout(path)
        except Exception as exc:
            c3 = {"ok": False, "error": str(exc)}
        item = {"kind": kind, "path": path.name, "zip": c1, "advice": c2, "com": c3}
        item["ok"] = bool(c1.get("ok") and c2.get("ok") and c3.get("ok"))
        ok_all = ok_all and item["ok"]
        report["items"].append(item)
        print(json.dumps(item, ensure_ascii=False))

    report["ok"] = ok_all
    print("SUMMARY", json.dumps({"ok": ok_all}, ensure_ascii=False))
    return 0 if ok_all else 1


if __name__ == "__main__":
    raise SystemExit(main())
