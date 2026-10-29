# -*- coding: utf-8 -*-
"""
test_office.py - Office 静默打印/版面引擎测试 (不需要真打印机, 不出纸)

覆盖:
  1. 版面参数规范化 (纯逻辑)
  2. 扩展名 -> 组件判定
  3. Word 导出 PDF: 改页边距生效, 原文件不被修改
  4. Excel 导出 PDF: "所有列缩到一页宽" 页数显著减少, 原文件不被修改
没有可用 Office/WPS 时, 第 3/4 部分自动跳过。
"""

import os
import sys
import hashlib
import tempfile
import shutil

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import office_core as oc

PASS = FAIL = 0


def check(ok, name, extra=""):
    global PASS, FAIL
    if ok:
        PASS += 1
        print(f"  PASS {name}" + (f"  ({extra})" if extra else ""))
    else:
        FAIL += 1
        print(f"  FAIL {name}" + (f"  ({extra})" if extra else ""))


def sha(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def section(t):
    print(f"\n[{t}]")


def page_count(pdf):
    import pypdfium2 as pdfium
    d = pdfium.PdfDocument(pdf)
    try:
        return len(d)
    finally:
        d.close()


# ---------------------------------------------------------------- 1. 参数规范化
section("1. 版面参数")
lay = oc.normalize_layout({"margins_mm": [20, 20, 15, 15], "orientation": "横向",
                           "excel_fit": "width"})
check(lay["margins_mm"] == (20.0, 20.0, 15.0, 15.0), "页边距解析")
check(lay["orientation"] == "landscape", "中文方向映射为 landscape")
check(lay["excel_fit"] == "width", "Excel 缩到一页宽")

lay = oc.normalize_layout({"margins_mm": [999, 0, 0, 0]})
check(lay["margins_mm"] is None, "异常页边距被丢弃")

lay = oc.normalize_layout({"excel_fit": "percent", "excel_scale": 5})
check(lay["excel_fit"] == "" and lay["excel_scale"] == 0, "非法百分比退回不缩放")

lay = oc.normalize_layout({"excel_fit": "percent", "excel_scale": 85})
check(lay["excel_scale"] == 85, "百分比缩放 85%")

check(oc.layout_summary({"margins_mm": [10, 10, 10, 10], "excel_fit": "width"})
      == "页边距 上10/下10/左10/右10mm, 列缩到一页宽", "摘要文案")
check(oc.layout_summary({}) == "", "空设置摘要为空")

# ---------------------------------------------------------------- 2. 扩展名判定
section("2. 扩展名 -> 组件")
check(oc.ext_kind("a.docx") == "word", "docx -> word")
check(oc.ext_kind("a.DOC") == "word", "大写 DOC -> word")
check(oc.ext_kind("a.xlsx") == "excel", "xlsx -> excel")
check(oc.ext_kind("a.xls") == "excel", "xls -> excel")
check(oc.ext_kind("a.pptx") == "ppt", "pptx -> ppt")
check(oc.ext_kind("a.pdf") is None, "pdf 不属于 Office")
check(oc.is_office_ext("wps") and not oc.is_office_ext(".pdf"), "is_office_ext 判断")

# ---------------------------------------------------------------- 3/4. 真导出
tmp = tempfile.mkdtemp(prefix="bp_test_office_")
try:
    section("3. Word 导出 PDF")
    docx = os.path.join(tmp, "t.docx")
    try:
        import win32com.client as w
        app = w.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Add()
        doc.Content.Text = "\r\n".join(f"第 {i} 行测试内容" for i in range(80))
        doc.SaveAs2(docx, FileFormat=16)
        doc.Close(0)
        app.Quit()
        have_word = True
    except Exception as e:
        have_word = False
        print(f"  跳过: 无可用 Word ({e})")

    if have_word:
        h0 = sha(docx)
        pdf1, d1 = oc.export_pdf(docx, {})
        n1 = page_count(pdf1)
        pdf2, d2 = oc.export_pdf(docx, {"margins_mm": [8, 8, 8, 8]})
        n2 = page_count(pdf2)
        check(n1 >= 1 and n2 >= 1, "两种版面都导出成功", f"页数 {n1} / {n2}")
        check(sha(docx) == h0, "原文件未被修改")
        shutil.rmtree(d1, ignore_errors=True)
        shutil.rmtree(d2, ignore_errors=True)

    section("4. Excel 导出 PDF (列缩到一页宽)")
    xlsx = os.path.join(tmp, "t.xlsx")
    try:
        import win32com.client as w
        et = w.DispatchEx("Excel.Application")
        et.Visible = False
        et.DisplayAlerts = False
        wb = et.Workbooks.Add()
        ws = wb.Worksheets(1)
        for c in range(1, 16):
            ws.Cells(1, c).Value = f"很长的列标题 Number {c}"
            ws.Columns(c).ColumnWidth = 20
        for r in range(2, 40):
            for c in range(1, 16):
                ws.Cells(r, c).Value = f"R{r}C{c}"
        wb.SaveAs(xlsx)
        wb.Close(False)
        et.Quit()
        have_excel = True
    except Exception as e:
        have_excel = False
        print(f"  跳过: 无可用 Excel ({e})")

    if have_excel:
        hx = sha(xlsx)
        pdf0, d0 = oc.export_pdf(xlsx, {})
        n0 = page_count(pdf0)
        pdf1, d1 = oc.export_pdf(xlsx, {"excel_fit": "width"})
        n1 = page_count(pdf1)
        check(n0 > 1, "默认版面确实分成多页", f"{n0} 页")
        check(n1 == 1, "缩到一页宽后只剩 1 页", f"{n0} -> {n1} 页")
        check(sha(xlsx) == hx, "原文件未被修改")
        for d in (d0, d1):
            shutil.rmtree(d, ignore_errors=True)

    section("5. 不支持的类型")
    bad = os.path.join(tmp, "x.pdf")
    open(bad, "wb").close()
    try:
        oc.export_pdf(bad, {})
        check(False, "非 Office 文件应报错")
    except oc.OfficeError:
        check(True, "非 Office 文件报 OfficeError")
finally:
    oc.close_session()
    shutil.rmtree(tmp, ignore_errors=True)

print(f"\n结果: {PASS} 通过, {FAIL} 失败")
sys.exit(1 if FAIL else 0)
