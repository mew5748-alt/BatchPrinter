# -*- coding: utf-8 -*-
"""
test_paper.py - 纸型 / 纸张来源 / 逐页打印 自测 (不出纸, 不给打印机发作业)

覆盖:
  1. 纸型与纸盒枚举 (DeviceCapabilities)
  2. DEVMODE 写入 dmPaperSize / dmDefaultSource
  3. 纸型真的改变了 DC 的可打印尺寸 (CreateDC 后量物理尺寸, 只建 DC 不打印)
  4. 逐页打印把每页拆成独立作业 (用假 DC 记录 StartDoc/EndDoc)
  5. print_file 把纸型/来源/逐页参数传到打印通道
"""

import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import printer_core as pc

PASS = FAIL = 0


def check(ok, name, extra=""):
    global PASS, FAIL
    if ok:
        PASS += 1
        print(f"  PASS {name}" + (f"  ({extra})" if extra else ""))
    else:
        FAIL += 1
        print(f"  FAIL {name}" + (f"  ({extra})" if extra else ""))


def section(t):
    print(f"\n[{t}]")


printers = pc.list_printers()
default_printer = pc.get_default_printer()

# ---------------------------------------------------------------- 1. 枚举
section("1. 纸型 / 纸盒枚举")
for p in printers[:3]:
    sizes = pc.list_paper_sizes(p)
    bins = pc.list_paper_bins(p)
    ids = [s[0] for s in sizes]
    print(f"    {p}: {len(sizes)} 种纸型, {len(bins)} 个来源")
    check(isinstance(sizes, list) and isinstance(bins, list), f"枚举不报错: {p}")
    check(len(ids) == len(set(ids)), f"纸型 ID 不重复: {p}")
check(pc.list_paper_sizes("") == [], "空打印机名 -> 空列表")
check(pc.list_paper_bins("不存在的打印机") == [], "不存在的打印机 -> 空列表")

a4 = [(pid, name) for pid, name, w, h in pc.list_paper_sizes(default_printer) if w == 210 and h == 297]
check(bool(a4), "默认打印机支持 A4 (210×297)", str(a4[:2]))
if a4:
    check(a4[0][0] == 9, "A4 的 DMPAPER 编号是 9", str(a4[0]))

bins0 = pc.list_paper_bins(default_printer)
check(all(isinstance(b[0], int) and b[1] for b in bins0), "纸盒条目都是 (编号, 名称)", str(bins0[:3]))

# ---------------------------------------------------------------- 2. DEVMODE
section("2. DEVMODE 写入纸型 / 来源")
if printers:
    h = pc.win32print.OpenPrinter(default_printer)
    dm = pc.win32print.GetPrinter(h, 2)["pDevMode"]
    pc.win32print.ClosePrinter(h)
    pc._modify_devmode(dm, paper_size=8, paper_source=1)
    check(dm.PaperSize == 8, "PaperSize 写入", str(dm.PaperSize))
    check(dm.DefaultSource == 1, "DefaultSource 写入", str(dm.DefaultSource))
    check(bool(dm.Fields & pc.DM_PAPERSIZE), "Fields 置位 DM_PAPERSIZE")
    check(bool(dm.Fields & pc.DM_DEFAULTSOURCE), "Fields 置位 DM_DEFAULTSOURCE")

    h = pc.win32print.OpenPrinter(default_printer)
    dm2 = pc.win32print.GetPrinter(h, 2)["pDevMode"]
    pc.win32print.ClosePrinter(h)
    before = (dm2.PaperSize, dm2.Fields)
    pc._modify_devmode(dm2, paper_size=None, paper_source=None)
    check((dm2.PaperSize, dm2.Fields) == before, "未设置时不动纸型字段")

check(pc.paper_devmode_values({"paper": 8, "source": 1}) == (8, 1), "opts -> devmode 值")
check(pc.paper_devmode_values({}) == (None, None), "空 opts -> 不改")

# ---------------------------------------------------------------- 3. 纸型真的生效
section("3. 纸型改变 DC 物理尺寸")
if a4 and "Fax" not in default_printer:
    try:
        dc = pc._open_printer_dc(default_printer, paper_size=8)      # A3
        w3, h3 = pc.dc_paper_mm(dc)
        dc.DeleteDC()
        dc = pc._open_printer_dc(default_printer, paper_size=9)      # A4
        w4, h4 = pc.dc_paper_mm(dc)
        dc.DeleteDC()
        check(abs(w3 - 297) <= 3 and abs(h3 - 420) <= 3, "A3 DC 是 297×420mm", f"{w3}×{h3}")
        check(abs(w4 - 210) <= 3 and abs(h4 - 297) <= 3, "A4 DC 是 210×297mm", f"{w4}×{h4}")
    except Exception as e:
        check(False, "创建 DC 失败", repr(e))

# ---------------------------------------------------------------- 4. 逐页打印
section("4. 逐页打印 (假 DC, 不出纸)")


class FakeDC:
    """只是一个记录器: 记下 StartDoc/StartPage/EndPage/EndDoc 的调用顺序"""

    def __init__(self):
        self.calls = []

    def StartDoc(self, name):
        self.calls.append(("doc", name))

    def StartPage(self):
        self.calls.append(("page",))

    def EndPage(self):
        self.calls.append(("endpage",))

    def EndDoc(self):
        self.calls.append(("enddoc",))

    def GetDeviceCaps(self, cap):
        return {8: 6600, 10: 5100, 88: 600, 110: 6600, 112: 5100,
                114: 0, 115: 0}.get(cap, 0)

    def GetHandleOutput(self):
        return 0

    def GetTextExtent(self, _s):
        return (50, 20)

    def SelectObject(self, _o):
        return True

    def TextOut(self, *_a):
        return True

    def DeleteDC(self):
        self.calls.append(("deletedc",))


class FakeDib:
    def __init__(self, im):
        self.size = im.size

    def draw(self, _hdc, _box):
        return True


tmp = tempfile.mkdtemp(prefix="bp_paper_")
try:
    from PIL import Image
    imgs = [Image.new("RGB", (300, 400), c) for c in ("white", "gray", "black")]
    pdf3 = os.path.join(tmp, "three.pdf")
    imgs[0].save(pdf3, "PDF", save_all=True, append_images=imgs[1:])
    txt = os.path.join(tmp, "many.txt")
    with open(txt, "w", encoding="utf-8") as f:
        f.write("\n".join(f"第 {i} 行内容" for i in range(500)))

    orig_dc = pc._open_printer_dc
    orig_dib = pc.ImageWin.Dib
    fake = FakeDC()
    try:
        pc._open_printer_dc = lambda *a, **k: fake
        pc.ImageWin.Dib = FakeDib
        logs = []

        pc.print_pdf_gdi(pdf3, "假打印机", 1, None, None, None, logs.append,
                         per_page=True, page_gap=0)
        docs = [c for c in fake.calls if c[0] == "doc"]
        check(len(docs) == 3, "3 页 PDF -> 3 个打印作业", str(docs))
        check(fake.calls.count(("enddoc",)) == 3, "每页各自 EndDoc")
        check(all("第%d页" % (i + 1) in docs[i][1] for i in range(len(docs))),
              "作业名带页码", str([d[1] for d in docs]))
        check(logs and "three.pdf" in logs[-1], "逐页模式日志正常", logs[-1])

        fake.calls.clear()
        pc.print_pdf_gdi(pdf3, "假打印机", 1, None, None, None, logs.append,
                         per_page=False)
        docs = [c for c in fake.calls if c[0] == "doc"]
        check(len(docs) == 1 and fake.calls.count(("enddoc",)) == 1,
              "非逐页: 整份一个作业", str(docs))
        check(fake.calls.count(("page",)) == 3, "非逐页仍是 3 页")

        # 页码 + 逐页: 只打第 1、3 页 -> 2 个作业
        fake.calls.clear()
        pc.print_pdf_gdi(pdf3, "假打印机", 1, None, None, {1, 3}, logs.append,
                         per_page=True, page_gap=0)
        docs = [c for c in fake.calls if c[0] == "doc"]
        check(len(docs) == 2, "页码范围 + 逐页 -> 2 个作业", str(docs))

        # 文本: 500 行 -> 多页 -> 逐页拆作业
        fake.calls.clear()
        pc.print_text_gdi(txt, "假打印机", 1, None, None, None, logs.append,
                          per_page=True, page_gap=0)
        docs = [c for c in fake.calls if c[0] == "doc"]
        check(len(docs) >= 2, "多页文本按页拆作业", f"{len(docs)} 个作业")
        check(fake.calls.count(("enddoc",)) == len(docs), "文本逐页 EndDoc 次数一致")

        # 图片: 3 帧 GIF 逐页
        gif = os.path.join(tmp, "multi.gif")
        imgs[0].save(gif, save_all=True, append_images=imgs[1:], duration=100)
        fake.calls.clear()
        pc.print_image_gdi(gif, "假打印机", 1, None, None, None, logs.append,
                           per_page=True, page_gap=0)
        docs = [c for c in fake.calls if c[0] == "doc"]
        check(len(docs) == 3, "3 帧图片 -> 3 个作业", str(len(docs)))
    finally:
        pc._open_printer_dc = orig_dc
        pc.ImageWin.Dib = orig_dib

    # ------------------------------------------------------------ 5. 参数贯通
    section("5. print_file 参数贯通")
    seen = {}

    def fake_pdf_gdi(path, printer, copies, color, duplex, page_set, log, dpi=200,
                     paper_size=None, paper_source=None, per_page=False, page_gap=0.0):
        seen.update(dict(path=path, dpi=dpi, paper_size=paper_size,
                         paper_source=paper_source, per_page=per_page,
                         page_gap=page_gap))

    orig = pc.print_pdf_gdi
    try:
        pc.print_pdf_gdi = fake_pdf_gdi
        pc.print_file(pdf3, "假打印机",
                      {"paper": 8, "source": 1, "per_page": True, "page_gap": 0.5},
                      {}, "", lambda _m: None)
    finally:
        pc.print_pdf_gdi = orig
    check(seen.get("paper_size") == 8, "纸型传到 PDF 通道", str(seen))
    check(seen.get("paper_source") == 1, "纸张来源传到 PDF 通道")
    check(seen.get("per_page") is True and abs(seen.get("page_gap") - 0.5) < 1e-6,
          "逐页与页间距传到 PDF 通道")

    # Word/Excel: COM 导出 PDF 后同样走 PDF 通道 (这里只验证参数贯通, 不打印)
    docx = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_probe", "probe.docx")
    if os.path.isfile(docx):
        seen.clear()
        logs = []
        orig = pc.print_pdf_gdi
        try:
            pc.print_pdf_gdi = fake_pdf_gdi
            pc.print_file(docx, "假打印机",
                          {"paper": 8, "per_page": True, "page_gap": 0},
                          {}, "", logs.append)
        except Exception as e:
            check(False, "Office 文件走 PDF 通道", repr(e))
        finally:
            pc.print_pdf_gdi = orig
            import office_core
            office_core.close_session()
        check(seen.get("paper_size") == 8 and seen.get("per_page") is True,
              "Office 文件带纸型/逐页进 PDF 通道", str(seen.get("paper_size")))
finally:
    import shutil
    shutil.rmtree(tmp, ignore_errors=True)

# ---------------------------------------------------------------- 6. 纯逻辑
section("6. 纯逻辑")
check(pc.plan_pages(3, None) == [1, 2, 3], "plan_pages 全部")
check(pc.plan_pages(3, {1, 3}) == [1, 3], "plan_pages 指定页")
check(pc.plan_pages(0, None) == [], "plan_pages 空文档")
check(pc._range_text({1, 2, 3, 5, 8, 9}) == "1-3,5,8-9", "页码范围拼串")
check(pc._pages_desc([1, 2, 3], 3) == "", "全部页不写页码")
check(pc._pages_desc([2], 3) == " (第2页)", "部分页写页码", pc._pages_desc([2], 3))
check(pc.to_float("2.5", 0) == 2.5 and pc.to_float(None, 1) == 1, "to_float")

print()
print(f"结果: {PASS} 通过, {FAIL} 失败")
sys.exit(1 if FAIL else 0)
