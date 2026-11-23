# -*- coding: utf-8 -*-
"""
printer_core.py - 批量打印核心引擎
支持引擎:
  1. GDI 直印 (文本/图片): 完全控制份数/色彩/双面/页码范围
  2. SumatraPDF (PDF):    命令行完整控制
  3. Office COM (Word/Excel/PPT): 改版面 -> 转 PDF -> 走 PDF 通道, 全程静默
  4. Shell 打印动词 (其他未知格式): 临时切默认打印机 + printto
  5. RAW 直写 (prn/ps/pcl): 字节流直送打印机
"""

import os
import time
import glob
import shutil
import subprocess
import contextlib
import threading

import win32print
import win32api
import win32con
import win32gui
import win32ui
from PIL import Image, ImageWin

import office_core
from office_core import OfficeUnavailable

# ---------------------------------------------------------------- 常量

# 文件类别
CAT_PDF = "pdf"
CAT_IMAGE = "image"
CAT_TEXT = "text"
CAT_OFFICE = "office"
CAT_RAW = "raw"

CATEGORY_EXTS = {
    CAT_PDF: {".pdf"},
    CAT_IMAGE: {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff"},
    CAT_TEXT: {".txt", ".log", ".csv", ".ini", ".md", ".bat", ".py", ".json", ".xml", ".html", ".htm"},
    CAT_OFFICE: {".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".rtf",
                 ".odt", ".ods", ".odp", ".wps", ".et", ".dps", ".dot", ".dotx",
                 ".xlsm", ".docm", ".pptm", ".csvx"},
    CAT_RAW: {".prn", ".ps", ".pcl"},
}

ARCHIVE_EXTS = {".zip", ".7z", ".rar", ".tar", ".gz", ".tgz", ".bz2", ".xz"}

# DEVMODE 常量
DM_ORIENTATION = 0x00000001
DM_PAPERSIZE = 0x00000002
DM_COPIES = 0x00000100
DM_DEFAULTSOURCE = 0x00000200
DM_DUPLEX = 0x00001000
DM_COLOR = 0x00000800
DM_COLLATE = 0x00008000

DMCOLOR_MONOCHROME = 1
DMCOLOR_COLOR = 2
DMDUP_SIMPLEX = 1
DMDUP_VERTICAL = 2      # 长边翻页
DMDUP_HORIZONTAL = 3    # 短边翻页
DMCOLLATE_TRUE = 1

# 纸张方向: 横图放纵向纸上只能占一半, 改成横向纸就能铺满
DMORIENT_PORTRAIT = 1
DMORIENT_LANDSCAPE = 2
ORIENT_MAP = {"portrait": DMORIENT_PORTRAIT, "landscape": DMORIENT_LANDSCAPE}

# DeviceCapabilities 能力号 (注意 DC_BINS 是 6, 不是 5)
DC_PAPERS = 2
DC_PAPERSIZE = 3
DC_BINS = 6
DC_BINNAMES = 12
DC_PAPERNAMES = 16

# 设置值 -> DEVMODE
COLOR_MAP = {"mono": DMCOLOR_MONOCHROME, "color": DMCOLOR_COLOR}
DUPLEX_MAP = {"simplex": DMDUP_SIMPLEX, "duplex_long": DMDUP_VERTICAL, "duplex_short": DMDUP_HORIZONTAL}

SUMATRA_SEARCH_PATHS = [
    r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
    r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "SumatraPDF", "SumatraPDF.exe"),
]


def all_known_exts(ext_map=None):
    """受支持的扩展名 (含映射别名)"""
    exts = set()
    for s in CATEGORY_EXTS.values():
        exts |= s
    exts |= ARCHIVE_EXTS
    if ext_map:
        exts |= {("." + k.lower().lstrip(".")) for k in ext_map}
    return exts


def resolve_category(ext, ext_map):
    """根据扩展名 + 映射表解析文件类别"""
    e = ext.lower().lstrip(".")
    mapped = ext_map.get(e, e)          # 别名映射 (键/值均不含点)
    e = "." + mapped.lstrip(".")
    for cat, s in CATEGORY_EXTS.items():
        if e in s:
            return cat
    return None


# ---------------------------------------------------------------- 页码范围

def parse_page_range(spec):
    """'1-3,5,8-10' -> set{1,2,3,5,8,9,10}; None/'' -> None(全部)"""
    if not spec or not spec.strip():
        return None
    pages = set()
    for part in spec.replace("，", ",").split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, _, b = part.partition("-")
            pages.update(range(int(a), int(b) + 1))
        else:
            pages.add(int(part))
    return pages


# ---------------------------------------------------------------- 打印机

def list_printers():
    try:
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        return [p[2] for p in win32print.EnumPrinters(flags)]
    except Exception:
        return []


def get_default_printer():
    try:
        return win32print.GetDefaultPrinter()
    except Exception:
        return ""


def find_sumatra(configured_path=""):
    if configured_path and os.path.isfile(configured_path):
        return configured_path
    for p in SUMATRA_SEARCH_PATHS:
        if p and os.path.isfile(p):
            return p
    for p in glob.glob(r"C:\Program Files*\*umatra*\SumatraPDF.exe"):
        return p
    which = shutil.which("SumatraPDF.exe")
    return which or ""


class PrinterSettingsError(Exception):
    pass


# ---------------------------------------------------------------- 纸型 / 纸张来源

def printer_port(printer_name):
    """打印机端口名 (DeviceCapabilities 需要 device + port 两个参数)"""
    try:
        h = win32print.OpenPrinter(printer_name)
        try:
            return win32print.GetPrinter(h, 2).get("pPortName", "") or ""
        finally:
            win32print.ClosePrinter(h)
    except Exception:
        return ""


def _capabilities(printer_name, cap):
    """DeviceCapabilities 的安全包装: 失败一律返回空元组"""
    try:
        port = printer_port(printer_name)
        return win32print.DeviceCapabilities(printer_name, port, cap) or ()
    except Exception:
        return ()


def _point_mm(p):
    """DC_PAPERSIZE 的 POINT(0.1mm) -> (宽mm, 高mm); 兼容 dict / tuple"""
    try:
        if isinstance(p, dict):
            w, h = int(p.get("x") or 0), int(p.get("y") or 0)
        else:
            w, h = int(p[0]), int(p[1])
        return int(round(w / 10.0)), int(round(h / 10.0))
    except Exception:
        return 0, 0


def list_paper_sizes(printer_name):
    """打印机支持的纸型: [(纸型ID, 名称, 宽mm, 高mm), ...]

    名称直接取驱动自带的名字 ("A4(210x297mm)" / "信纸")，中文系统下就是用户
    在打印机驱动里看到的那套叫法。
    """
    if not printer_name:
        return []
    ids = _capabilities(printer_name, DC_PAPERS)
    names = _capabilities(printer_name, DC_PAPERNAMES)
    sizes = _capabilities(printer_name, DC_PAPERSIZE)
    out, seen = [], set()
    for i, pid in enumerate(ids):
        try:
            pid = int(pid)
        except Exception:
            continue
        if pid in seen or pid <= 0:
            continue
        seen.add(pid)
        name = str(names[i]).strip() if i < len(names) else ""
        w, h = _point_mm(sizes[i]) if i < len(sizes) else (0, 0)
        out.append((pid, name or ("纸型 %d" % pid), w, h))
    return out


def list_paper_bins(printer_name):
    """打印机纸盒 / 纸张来源: [(来源ID, 名称), ...]"""
    if not printer_name:
        return []
    ids = _capabilities(printer_name, DC_BINS)
    names = _capabilities(printer_name, DC_BINNAMES)
    out, seen = [], set()
    for i, bid in enumerate(ids):
        try:
            bid = int(bid)
        except Exception:
            continue
        if bid in seen:
            continue
        seen.add(bid)
        name = str(names[i]).strip() if i < len(names) else ""
        out.append((bid, name or ("纸盒 %d" % bid)))
    return out


def _paper_size_to_int(v):
    try:
        return int(v) if v else None
    except (TypeError, ValueError):
        return None


def _modify_devmode(dm, copies=None, color=None, duplex=None, collate=None,
                    paper_size=None, paper_source=None, orientation=None):
    """在 PyDEVMODE 上应用设置并置位对应 Fields"""
    fields = 0
    if copies is not None and copies > 0:
        dm.Copies = copies
        if collate:
            dm.Collate = DMCOLLATE_TRUE
            fields |= DM_COLLATE
        fields |= DM_COPIES
    if color is not None:
        dm.Color = color
        fields |= DM_COLOR
    if duplex is not None:
        dm.Duplex = duplex
        fields |= DM_DUPLEX
    if paper_size:
        dm.PaperSize = int(paper_size)
        fields |= DM_PAPERSIZE
    if paper_source:
        dm.DefaultSource = int(paper_source)
        fields |= DM_DEFAULTSOURCE
    if orientation:
        dm.Orientation = int(orientation)
        fields |= DM_ORIENTATION
    if fields:
        dm.Fields |= fields
    return dm


def apply_printer_defaults(printer_name, copies=None, color=None, duplex=None,
                           paper_size=None, paper_source=None):
    """临时把设置写入打印机默认 DEVMODE, 返回恢复句柄 (传给 restore_printer_defaults)"""
    h = win32print.OpenPrinter(printer_name, {"DesiredAccess": win32print.PRINTER_ALL_ACCESS})
    try:
        info = win32print.GetPrinter(h, 2)
        original = info["pDevMode"]
        dm = info["pDevMode"]
        _modify_devmode(dm, copies=copies, color=color, duplex=duplex,
                        collate=(copies or 0) > 1,
                        paper_size=paper_size, paper_source=paper_source)
        info["pDevMode"] = dm
        win32print.SetPrinter(h, 2, info, 0)
        return h, original
    except Exception as e:
        try:
            win32print.ClosePrinter(h)
        except Exception:
            pass
        raise PrinterSettingsError(str(e))


@contextlib.contextmanager
def temp_printer_defaults(printer_name, log=None, settle=0.8, **kw):
    """临时改打印机默认的纸型/纸盒, 退出时恢复 (给 Sumatra / Shell 通道用)

    这两个通道自己创建 DC 时读的是打印机默认 DEVMODE, 拿不到我们传入的
    devmode, 所以只能临时改默认值再改回来。
    """
    ctx = None
    if any(kw.get(k) for k in ("paper_size", "paper_source")):
        try:
            ctx = apply_printer_defaults(printer_name, **kw)
        except PrinterSettingsError as e:
            if log:
                log(f"提示: 无法临时应用纸型/进纸盒 ({e}), 将按打印机当前默认值打印")
    try:
        yield
    finally:
        if ctx:
            time.sleep(settle)          # 等驱动把作业读走再恢复
            restore_printer_defaults(*ctx)



def restore_printer_defaults(h, original):
    try:
        info = win32print.GetPrinter(h, 2)
        info["pDevMode"] = original
        win32print.SetPrinter(h, 2, info, 0)
    except Exception:
        pass
    finally:
        try:
            win32print.ClosePrinter(h)
        except Exception:
            pass


# ---------------------------------------------------------------- 默认打印机临时切换

def get_default_printer_name():
    """当前系统默认打印机名 (失败返回空串)"""
    try:
        return win32print.GetDefaultPrinter()
    except Exception:
        return ""


def printer_driver_port(printer_name):
    """返回 (驱动名, 端口名), 供 printto 动词拼参数"""
    try:
        h = win32print.OpenPrinter(printer_name)
        try:
            info = win32print.GetPrinter(h, 2)
            return info.get("pDriverName", "") or "", info.get("pPortName", "") or ""
        finally:
            win32print.ClosePrinter(h)
    except Exception:
        return "", ""


class default_printer_guard:
    """打印期间把系统默认打印机临时切到目标打印机, 退出时恢复.

    必要性: ShellExecute 的 'print' 动词没有指定打印机的参数, 关联程序一律打到
    系统默认打印机。不切换的话, 界面上选的打印机对 Office/未知格式完全无效。
    """

    def __init__(self, printer_name, log=None):
        self.target = (printer_name or "").strip()
        self.log = log
        self.original = ""
        self.changed = False

    def __enter__(self):
        self.original = get_default_printer_name()
        if not self.target or self.target == self.original:
            return self
        try:
            win32print.SetDefaultPrinter(self.target)
            self.changed = True
            if self.log:
                self.log(f"临时把系统默认打印机切到「{self.target}」"
                         f"(打印结束后自动恢复为「{self.original}」)")
        except Exception as e:
            if self.log:
                self.log(f"⚠ 无法临时切换默认打印机 ({e}), 将尝试用 printto 动词指定打印机")
        return self

    def __exit__(self, *_exc):
        if self.changed and self.original:
            try:
                win32print.SetDefaultPrinter(self.original)
                if self.log:
                    self.log(f"已恢复系统默认打印机为「{self.original}」")
            except Exception:
                pass
        return False


# ---------------------------------------------------------------- GDI 直印 (文本/图片)

def _open_printer_dc(printer_name, copies=None, color=None, duplex=None,
                     paper_size=None, paper_source=None, orientation=None):
    hprinter = win32print.OpenPrinter(printer_name)
    dm = win32print.GetPrinter(hprinter, 2)["pDevMode"]
    win32print.ClosePrinter(hprinter)
    _modify_devmode(dm, copies=copies, color=color, duplex=duplex,
                    collate=(copies or 0) > 1,
                    paper_size=paper_size, paper_source=paper_source,
                    orientation=orientation)
    hdc = win32gui.CreateDC("WINSPOOL", printer_name, dm)
    return win32ui.CreateDCFromHandle(hdc)


def image_orientation(path):
    """图片本身是横的还是竖的: 'landscape' / 'portrait' (取不到就当竖向)"""
    try:
        with Image.open(path) as im:
            w, h = im.size
            try:
                from PIL import ImageOps
                exif = im.getexif()
                if exif and int(exif.get(0x0112, 1) or 1) in (5, 6, 7, 8):
                    w, h = h, w         # 手机竖拍(EXIF 旋转)要先摆正再判断
            except Exception:
                pass
        return "landscape" if w > h else "portrait"
    except Exception:
        return "portrait"


def resolve_orientation(setting, paths):
    """'auto' / 'portrait' / 'landscape' -> DEVMODE 方向值 (None = 跟随驱动)

    auto: 看这批图片里第一张的实际朝向 —— 横图就用横向纸, 不然横图挤在竖向 A4 上
    只能占中间一条, 白白浪费一半纸面。
    """
    setting = (setting or "auto").lower()
    if setting in ORIENT_MAP:
        return ORIENT_MAP[setting]
    for p in paths:
        try:
            return ORIENT_MAP[image_orientation(p)]
        except Exception:
            continue
    return None


def dc_paper_mm(dc):
    """DC 上驱动实际回报的纸张尺寸 (毫米), 用来确认纸型真的生效了"""
    try:
        lpx = dc.GetDeviceCaps(win32con.LOGPIXELSX) or 600
        w = dc.GetDeviceCaps(win32con.PHYSICALWIDTH)
        h = dc.GetDeviceCaps(win32con.PHYSICALHEIGHT)
        return int(round(w * 25.4 / lpx)), int(round(h * 25.4 / lpx))
    except Exception:
        return 0, 0


def plan_pages(count, page_set):
    """要打印的页码列表 (1..count 有序, 按页码范围过滤; 空列表表示都不用打)"""
    if count <= 0:
        return []
    if page_set is None:
        return list(range(1, count + 1))
    return [p for p in range(1, count + 1) if p in page_set]


def _pages_desc(pages, total):
    """日志用: 只打了部分页面时列出页码"""
    if not pages or len(pages) == total:
        return ""
    txt = ",".join(str(p) for p in pages[:20])
    return " (第%s页%s)" % (txt, "…" if len(pages) > 20 else "")


class _DocWriter:
    """管理 StartDoc/EndDoc 的生命周期

    普通模式: 整个文件一个打印作业; 逐页模式: 每页一个独立作业 (作业名带页码),
    页与页之间可停一会儿——卡纸/漏页时容易从任意一页补打, 手动放纸也能一页一页来。
    """

    def __init__(self, dc, base, per_page=False, gap=0.0):
        self.dc = dc
        self.base = base
        self.per_page = bool(per_page)
        try:
            self.gap = max(0.0, float(gap or 0))
        except (TypeError, ValueError):
            self.gap = 0.0
        self.open = False

    def begin(self, page_no):
        if self.per_page:
            self.dc.StartDoc("%s - 第%d页" % (self.base, page_no))
            self.open = True
        elif not self.open:
            self.dc.StartDoc(self.base)
            self.open = True

    def end(self, last=False):
        self.dc.EndPage()
        if self.per_page:
            self.dc.EndDoc()
            self.open = False
            if not last and self.gap:
                time.sleep(self.gap)

    def close(self):
        if self.open:
            try:
                self.dc.EndDoc()
            except Exception:
                pass
            self.open = False


def print_text_gdi(path, printer_name, copies, color, duplex, page_set, log,
                   paper_size=None, paper_source=None, per_page=False, page_gap=0.0,
                   orientation=None):
    """文本文件 GDI 直印, 支持页码范围 / 逐页打印"""
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        lines = f.read().replace("\r\n", "\n").replace("\r", "\n").split("\n")

    base = os.path.basename(path)
    dc = _open_printer_dc(printer_name, copies=copies, color=color, duplex=duplex,
                          paper_size=paper_size, paper_source=paper_source,
                          orientation=orientation)
    try:
        font = win32ui.CreateFont({"name": "Consolas", "height": -110, "weight": 400})
        dc.SelectObject(font)
        lh = dc.GetTextExtent("Ag")[1] + 6
        margin_x = 200
        try:
            lpx = dc.GetDeviceCaps(win32con.LOGPIXELSX)
            margin_x = int(lpx * 0.35)
        except Exception:
            pass

        page_h = dc.GetDeviceCaps(win32con.VERTRES)
        per = max(1, (page_h - margin_x * 2) // lh)     # 每页能排多少行
        total = len(lines)
        if total <= 0:
            log(f"跳过 {base}: 文件内容为空")
            return
        page_count = (total + per - 1) // per
        todo = plan_pages(page_count, page_set)
        if not todo:
            log(f"跳过 {base}: 页码范围不包含任何页")
            return

        writer = _DocWriter(dc, base, per_page, page_gap)
        for k, pno in enumerate(todo):
            chunk = lines[(pno - 1) * per: pno * per]
            writer.begin(pno)
            dc.StartPage()
            y = margin_x
            for line in chunk:
                dc.TextOut(margin_x, y, line[:180])
                y += lh
            writer.end(last=(k == len(todo) - 1))
        writer.close()
    finally:
        dc.DeleteDC()
    log(f"GDI 文本打印完成: {base}" + _pages_desc(todo, page_count))


def print_image_gdi(path, printer_name, copies, color, duplex, page_set, log,
                    paper_size=None, paper_source=None, per_page=False, page_gap=0.0,
                    orientation=None):
    """图片 GDI 直印 (等比缩放居中), 多帧 tif/tiff 按页处理"""
    img = Image.open(path)
    frames = []
    need = max(page_set) if page_set else 0      # 只读需要的那几帧, 避免大 TIFF 撑爆内存
    try:
        n = getattr(img, "n_frames", 1)
        for i in range(n):
            img.seek(i)
            frames.append(img.convert("RGB").copy())
            if need and len(frames) >= need:
                break
    except Exception:
        frames = [img.convert("RGB")]

    base = os.path.basename(path)
    todo = plan_pages(len(frames), page_set)
    if not todo:
        log(f"跳过 {base}: 页码范围不包含任何页")
        return

    dc = _open_printer_dc(printer_name, copies=copies, color=color, duplex=duplex,
                          paper_size=paper_size, paper_source=paper_source,
                          orientation=orientation)
    try:
        pw, ph = dc.GetDeviceCaps(win32con.HORZRES), dc.GetDeviceCaps(win32con.VERTRES)
        try:
            lpx = dc.GetDeviceCaps(win32con.LOGPIXELSX)
            m = int(lpx * 0.2)
        except Exception:
            m = 30
        pw, ph = pw - m * 2, ph - m * 2

        writer = _DocWriter(dc, base, per_page, page_gap)
        for k, pno in enumerate(todo):
            writer.begin(pno)
            dc.StartPage()
            dib = ImageWin.Dib(frames[pno - 1])
            iw, ih = dib.size
            scale = min(pw / iw, ph / ih)
            w, h = int(iw * scale), int(ih * scale)
            x0 = m + (pw - w) // 2
            y0 = m + (ph - h) // 2
            dib.draw(dc.GetHandleOutput(), (x0, y0, x0 + w, y0 + h))
            writer.end(last=(k == len(todo) - 1))
        writer.close()
    finally:
        dc.DeleteDC()
    log(f"GDI 图片打印完成: {base}" + _pages_desc(todo, len(frames)))


# ---------------------------------------------------------------- PDF GDI 直印 (pypdfium2 渲染)

def print_pdf_gdi(path, printer_name, copies, color, duplex, page_set, log, dpi=200,
                  paper_size=None, paper_source=None, per_page=False, page_gap=0.0,
                  orientation=None):
    """用 pypdfium2 渲染 PDF 后经 GDI 直印, 完整支持份数/色彩/双面/页码/逐页"""
    try:
        import pypdfium2 as pdfium
    except ImportError:
        raise RuntimeError("PDF 引擎不可用: 请安装 pypdfium2 (pip install pypdfium2) 或 SumatraPDF")

    base = os.path.basename(path)
    try:
        pdf = pdfium.PdfDocument(path)
    except Exception as e:
        raise RuntimeError(f"PDF 打开失败: {e}")
    try:
        n = len(pdf)
        todo = plan_pages(n, page_set)
        if not todo:
            log(f"跳过 {base}: 页码范围不包含任何页")
            return

        dc = _open_printer_dc(printer_name, copies=copies, color=color, duplex=duplex,
                              paper_size=paper_size, paper_source=paper_source,
                              orientation=orientation)
        try:
            pw, ph = dc.GetDeviceCaps(win32con.HORZRES), dc.GetDeviceCaps(win32con.VERTRES)
            try:
                lpx = dc.GetDeviceCaps(win32con.LOGPIXELSX)
                m = int(lpx * 0.15)
            except Exception:
                m = 25
            pw, ph = pw - m * 2, ph - m * 2

            writer = _DocWriter(dc, base, per_page, page_gap)
            for k, pno in enumerate(todo):
                page = pdf[pno - 1]
                scale = dpi / 72.0
                pil = page.render(scale=scale).to_pil().convert("RGB")
                page.close()
                writer.begin(pno)
                dc.StartPage()
                dib = ImageWin.Dib(pil)
                iw, ih = dib.size
                s = min(pw / iw, ph / ih)
                w, h = int(iw * s), int(ih * s)
                x0 = m + (pw - w) // 2
                y0 = m + (ph - h) // 2
                dib.draw(dc.GetHandleOutput(), (x0, y0, x0 + w, y0 + h))
                writer.end(last=(k == len(todo) - 1))
            writer.close()
        finally:
            dc.DeleteDC()
    finally:
        try:
            pdf.close()
        except Exception:
            pass
    log(f"PDF 直印完成 (内置引擎): {base}" + _pages_desc(todo, n))



# ---------------------------------------------------------------- 多图拼版 (双面)

IMAGE_PER_SIDE_CHOICES = (1, 2, 4)      # 每面可放 1/2/4 张图

DC_DUPLEX = 7                           # DeviceCapabilities: 是否支持双面


def printer_supports_duplex(printer_name):
    """打印机是否支持自动双面: True / False / None(驱动没报, 未知)"""
    if not printer_name:
        return None
    try:
        v = win32print.DeviceCapabilities(printer_name, printer_port(printer_name), DC_DUPLEX)
        return int(v or 0) > 0
    except Exception:
        return None


def image_page_count(path):
    """一张图片的页数: 多帧 TIFF 的每一帧算一页"""
    try:
        with Image.open(path) as im:
            return max(1, int(getattr(im, "n_frames", 1) or 1))
    except Exception:
        return 1


def _grid_shape(per_side):
    """每面张数 -> (列, 行)"""
    per_side = max(1, int(per_side or 1))
    if per_side >= 4:
        return 2, 2
    if per_side == 3:
        return 1, 3
    return 1, per_side


def _open_frame(path, index=0):
    """打开图片第 index 帧, 按 EXIF 方向摆正后转 RGB (手机竖拍的图不会躺倒)"""
    im = Image.open(path)
    try:
        if index:
            im.seek(index)
        try:
            from PIL import ImageOps
            im = ImageOps.exif_transpose(im)
        except Exception:
            pass
        return im.convert("RGB")
    finally:
        try:
            im.close()
        except Exception:
            pass


def print_images_merged(paths, printer_name, copies, color, duplex, log,
                        per_side=1, paper_size=None, paper_source=None,
                        margin_mm=4.0, gap_mm=3.0, orientation=None):
    """把多张图片按顺序拼成**一份多页文档**一次发出, 双面打印时正反自动配对.

    为什么需要它: 每个图片文件单独发一个作业时, 每个作业只有一页, 双面设置等于没用
    ——打印机只会往每张纸的正面打一张图, 背面全空。合并成一个作业后, 页序是
    图1(正面)、图2(背面)、图3(正面)…, 纸耗直接减半。

    per_side=1: 每页一张图 → 一张纸正反各一张 (最常用)
    per_side=2/4: 每面再拼 2/4 张 → 一张纸 4/8 张图 (更省纸, 图会缩小)

    返回 (页数, 图片帧数)
    """
    frames = []
    for p in paths:
        n = image_page_count(p)
        frames.extend((p, i) for i in range(n))
    if not frames:
        return 0, 0

    per_side = max(1, int(per_side or 1))
    cols, rows = _grid_shape(per_side)
    pages = [frames[i:i + per_side] for i in range(0, len(frames), per_side)]
    label = os.path.basename(paths[0]) if len(paths) == 1 else f"{len(paths)} 张图片"

    dc = _open_printer_dc(printer_name, copies=copies, color=color, duplex=duplex,
                          paper_size=paper_size, paper_source=paper_source,
                          orientation=orientation)
    try:
        dpi = dc.GetDeviceCaps(win32con.LOGPIXELSX) or 600
        pw, ph = dc.GetDeviceCaps(win32con.HORZRES), dc.GetDeviceCaps(win32con.VERTRES)
        margin = int(dpi * max(0.0, to_float(margin_mm, 4.0)) / 25.4)
        gap = int(dpi * max(0.0, to_float(gap_mm, 3.0)) / 25.4)
        cell_w = max(1, (pw - margin * 2 - gap * (cols - 1)) // cols)
        cell_h = max(1, (ph - margin * 2 - gap * (rows - 1)) // rows)

        writer = _DocWriter(dc, f"{label} (拼版{per_side}张/面)")
        try:
            for k, group in enumerate(pages):
                writer.begin(k + 1)
                dc.StartPage()
                for j, (path, idx) in enumerate(group):
                    r, c = divmod(j, cols)
                    x0 = margin + c * (cell_w + gap)
                    y0 = margin + r * (cell_h + gap)
                    try:
                        im = _open_frame(path, idx)
                    except Exception as e:
                        log(f"⚠ 跳过 {os.path.basename(path)}: {e}")
                        continue
                    try:
                        dib = ImageWin.Dib(im)
                        iw, ih = dib.size
                        if iw <= 0 or ih <= 0:
                            continue
                        s = min(cell_w / float(iw), cell_h / float(ih))
                        w, h = max(1, int(iw * s)), max(1, int(ih * s))
                        dx = x0 + (cell_w - w) // 2
                        dy = y0 + (cell_h - h) // 2
                        dib.draw(dc.GetHandleOutput(), (dx, dy, dx + w, dy + h))
                    finally:
                        im.close()
                writer.end(last=(k == len(pages) - 1))
        finally:
            writer.close()
    finally:
        dc.DeleteDC()

    duplex_txt = ""
    if duplex == DMDUP_VERTICAL:
        duplex_txt = " · 双面(长边)"
    elif duplex == DMDUP_HORIZONTAL:
        duplex_txt = " · 双面(短边)"
    sheets = (len(pages) + 1) // 2 if duplex_txt else len(pages)
    log(f"多图拼版打印完成: {len(frames)} 张图 → {len(pages)} 页"
        f"(每面 {per_side} 张){duplex_txt} ≈ {sheets} 张纸")
    return len(pages), len(frames)


# ---------------------------------------------------------------- 连续打印 (跨文件接续)

CONT_CATS = (CAT_IMAGE, CAT_PDF, CAT_OFFICE, CAT_TEXT)   # 能接续成一个作业的格式


def _text_font(size):
    """等宽字体 (与 GDI 文本打印观感一致), 都找不到就用默认位图字体"""
    from PIL import ImageFont
    for name in ("consola.ttf", "cour.ttf", "msyh.ttc", "simsun.ttc"):
        try:
            return ImageFont.truetype(name, size)
        except Exception:
            continue
    try:
        return ImageFont.load_default()
    except Exception:
        return None


def _render_text_pages(path, pw, ph, page_set, dpi, log):
    """把文本文件排成整页图像, 逐页 yield"""
    from PIL import ImageDraw
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            text = f.read()
    except Exception as e:
        log(f"⚠ 跳过 {os.path.basename(path)}: {e}")
        return
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    size = max(10, int(dpi * 10 / 72.0))          # 10pt
    font = _text_font(size)
    lh = max(1, int(size * 1.35))
    margin = max(10, int(dpi * 3.5 / 25.4))       # 3.5mm
    per = max(1, (ph - margin * 2) // lh)
    total = max(1, (len(lines) + per - 1) // per)
    for pno in plan_pages(total, page_set):
        img = Image.new("RGB", (pw, ph), (255, 255, 255))
        draw = ImageDraw.Draw(img)
        y = margin
        for line in lines[(pno - 1) * per: pno * per]:
            if font is not None:
                draw.text((margin, y), line[:220], fill=(0, 0, 0), font=font)
            y += lh
        yield img


def _iter_job_pages(job, pw, ph, dpi, log):
    """一个文件 -> 逐页的 PIL 图像 (失败只跳过这一个文件, 不影响整批)"""
    cat = job.get("cat")
    path = job.get("path")
    page_set = job.get("pages")
    name = os.path.basename(path)

    if cat == CAT_IMAGE:
        n = image_page_count(path)
        for i in plan_pages(n, page_set):
            try:
                yield _open_frame(path, i - 1)
            except Exception as e:
                log(f"⚠ 跳过 {name} 第{i}页: {e}")
        return

    if cat == CAT_TEXT:
        for img in _render_text_pages(path, pw, ph, page_set, dpi, log):
            yield img
        return

    # PDF / Office: 拿一个 PDF 再逐页渲染
    work = None
    try:
        if cat == CAT_OFFICE:
            pdf_path, work = office_core.export_pdf(
                path, job.get("layout"), log, kind=office_core.ext_kind(os.path.splitext(path)[1]))
        else:
            pdf_path = path
        try:
            import pypdfium2 as pdfium
        except ImportError:
            log(f"⚠ 跳过 {name}: 没有 PDF 渲染引擎 (pypdfium2)")
            return
        try:
            pdf = pdfium.PdfDocument(pdf_path)
        except Exception as e:
            log(f"⚠ 跳过 {name}: PDF 打开失败 ({e})")
            return
        try:
            n = len(pdf)
            for pno in plan_pages(n, page_set):
                try:
                    page = pdf[pno - 1]
                    yield page.render(scale=dpi / 72.0).to_pil().convert("RGB")
                    page.close()
                except Exception as e:
                    log(f"⚠ {name} 第{pno}页渲染失败: {e}")
        finally:
            try:
                pdf.close()
            except Exception:
                pass
    finally:
        if work:
            shutil.rmtree(work, ignore_errors=True)


def print_files_continuous(jobs, printer_name, copies, color, duplex, log,
                           per_side=1, paper_size=None, paper_source=None,
                           dpi=200, margin_mm=4.0, gap_mm=3.0, orientation=None):
    """把多个文件的所有页**接续成同一个打印作业**发出.

    为什么: 每个文件单独成作业时, 打印机总是从新纸的正面开始 —— 前一个文件页数为奇数
    时, 最后那张纸的背面就空着, 下一个文件又从新纸开始, 中间白白浪费一个面。
    接续成一个作业后, 页码连续排下去, 双面自动正反配对, 文件之间不留空面。
    (注意: 页数总量为奇数时, 最后一张纸的背面仍然会是空的, 这是物理必然)

    jobs: [{"path", "cat", "pages"(set|None), "layout"(dict|None)}, ...]
    返回 (总页数, 成功输出的文件数)
    """
    if not jobs:
        return 0, 0

    per_side = max(1, int(per_side or 1))
    cols, rows = _grid_shape(per_side)
    dc = _open_printer_dc(printer_name, copies=copies, color=color, duplex=duplex,
                          paper_size=paper_size, paper_source=paper_source,
                          orientation=orientation)
    pages_total = 0
    files_done = 0
    try:
        dpi_dc = dc.GetDeviceCaps(win32con.LOGPIXELSX) or dpi
        pw, ph = dc.GetDeviceCaps(win32con.HORZRES), dc.GetDeviceCaps(win32con.VERTRES)
        margin = int(dpi_dc * max(0.0, to_float(margin_mm, 4.0)) / 25.4)
        gap = int(dpi_dc * max(0.0, to_float(gap_mm, 3.0)) / 25.4)
        cell_w = max(1, (pw - margin * 2 - gap * (cols - 1)) // cols)
        cell_h = max(1, (ph - margin * 2 - gap * (rows - 1)) // rows)

        # 文本页按"渲染 dpi"生成 (直接用打印机 600dpi 像素开会是近百 MB 一张图)
        pw_mm = pw * 25.4 / max(1, dpi_dc)
        ph_mm = ph * 25.4 / max(1, dpi_dc)
        render_w = max(1, int(pw_mm * dpi / 25.4))
        render_h = max(1, int(ph_mm * dpi / 25.4))

        writer = _DocWriter(dc, "连续打印")
        bucket = []          # 当前这一面待放的图

        def flush_side():
            nonlocal pages_total
            if not bucket:
                return
            writer.begin(pages_total + 1)
            dc.StartPage()
            for j, im in enumerate(bucket):
                try:
                    r, c = divmod(j, cols)
                    x0 = margin + c * (cell_w + gap)
                    y0 = margin + r * (cell_h + gap)
                    dib = ImageWin.Dib(im)
                    iw, ih = dib.size
                    if iw <= 0 or ih <= 0:
                        continue
                    s = min(cell_w / float(iw), cell_h / float(ih))
                    w, h = max(1, int(iw * s)), max(1, int(ih * s))
                    dx = x0 + (cell_w - w) // 2
                    dy = y0 + (cell_h - h) // 2
                    dib.draw(dc.GetHandleOutput(), (dx, dy, dx + w, dy + h))
                finally:
                    try:
                        im.close()
                    except Exception:
                        pass
            writer.end()
            pages_total += 1
            bucket.clear()

        for job in jobs:
            got = 0
            try:
                for im in _iter_job_pages(job, render_w, render_h, dpi, log):
                    bucket.append(im)
                    got += 1
                    if len(bucket) >= per_side:
                        flush_side()
            except Exception as e:
                log(f"⚠ {os.path.basename(job.get('path', ''))}: {e}")
            if got:
                files_done += 1
        flush_side()
        writer.close()
    finally:
        dc.DeleteDC()

    duplex_txt = {DMDUP_VERTICAL: "双面(长边)", DMDUP_HORIZONTAL: "双面(短边)"}.get(duplex, "")
    sheets = (pages_total + 1) // 2 if duplex_txt else pages_total
    log(f"连续打印: {files_done} 个文件接续成 {pages_total} 页"
        + (f" · {duplex_txt} ≈ {sheets} 张纸" if duplex_txt else f" ≈ {sheets} 张纸"))
    return pages_total, files_done


# ---------------------------------------------------------------- RAW 直写

def print_raw(path, printer_name, log):
    with open(path, "rb") as f:
        data = f.read()
    h = win32print.OpenPrinter(printer_name)
    try:
        win32print.StartDocPrinter(h, 1, (os.path.basename(path), None, "RAW"))
        win32print.StartPagePrinter(h)
        win32print.WritePrinter(h, data)
        win32print.EndPagePrinter(h)
        win32print.EndDocPrinter(h)
    finally:
        win32print.ClosePrinter(h)
    log(f"RAW 直写完成: {os.path.basename(path)}")


# ---------------------------------------------------------------- Shell 动词

_SHELL_ERRORS = {
    0: "内存不足", 2: "找不到文件", 3: "找不到路径", 5: "拒绝访问",
    8: "内存不足", 26: "共享冲突", 27: "文件名无效", 28: "命令行过长",
    29: "磁盘空间不足", 30: "内存不足", 31: "关联程序无打印动词或发生错误",
}


def _shell_printto(path, printer_name, dirname, log):
    """用 printto 动词显式指定打印机 (参数: 打印机名 驱动名 端口名)"""
    driver, port = printer_driver_port(printer_name)
    params = f'"{printer_name}"'
    if driver:
        params += f' "{driver}"'
        if port:
            params += f' "{port}"'
    ret = win32api.ShellExecute(0, "printto", path, params, dirname, win32con.SW_HIDE)
    if ret > 32:
        if log:
            log(f"Shell 打印已发送 (printto → {printer_name}): {os.path.basename(path)}")
        return True
    return False


def shell_print(path, printer_name="", log=None):
    """调用 Windows Shell 打印动词 (由文件关联程序处理)

    注意: 'print' 动词不能指定打印机, 只会打到系统默认打印机——所以调用方应先用
    default_printer_guard 把默认打印机切到目标。若默认打印机仍不是目标打印机
    (切换失败), 优先用 'printto' 动词显式指定, 避免打到错误的打印机。
    """
    dirname = os.path.dirname(path)
    base = os.path.basename(path)

    default_is_target = (not printer_name) or (get_default_printer_name() == printer_name)

    # 默认打印机不是目标时: 先用 printto (能显式指定打印机)
    if not default_is_target and _shell_printto(path, printer_name, dirname, log):
        return

    ret = win32api.ShellExecute(0, "print", path, None, dirname, win32con.SW_HIDE)
    if ret > 32:
        if log:
            log(f"Shell 打印已发送: {base}"
                + ("" if default_is_target else f" (目标打印机 {printer_name})"))
        return

    # 'print' 不可用 (常见于默认打开程序是 UWP 应用) -> 再用 printto 试一次
    if printer_name and default_is_target and _shell_printto(path, printer_name, dirname, log):
        return

    raise RuntimeError(
        f"Shell 打印失败: {_SHELL_ERRORS.get(ret, f'错误码 {ret}')} ({os.path.splitext(path)[1]})。"
        f"常见原因: 该格式的默认打开程序是 UWP 应用 (如 Edge/照片) 或未注册经典打印动词。"
        f"建议: 在系统设置中把该格式默认打开方式改为传统程序 (如 Adobe/WPS/Office), "
        f"或在扩展名映射中将其映射到其他格式")


# ---------------------------------------------------------------- SumatraPDF

def print_pdf_sumatra(sumatra_exe, path, printer_name, copies, color, duplex, page_set, log,
                      paper_size=None, paper_source=None):
    settings = []
    if copies and copies > 1:
        settings.append(f"{copies}x")
    if duplex == DMDUP_SIMPLEX:
        settings.append("simplex")
    elif duplex == DMDUP_VERTICAL:
        settings.append("duplex")
    elif duplex == DMDUP_HORIZONTAL:
        settings.append("duplexshort")
    if color == DMCOLOR_MONOCHROME:
        settings.append("monochrome")
    elif color == DMCOLOR_COLOR:
        settings.append("color")
    if page_set:
        settings.append(_range_text(page_set))

    cmd = [sumatra_exe, "-print-to", printer_name, "-silent"]
    if settings:
        cmd += ["-print-settings", ",".join(settings)]
    cmd.append(path)
    # Sumatra 自己创建 DC, 拿不到我们传的 devmode; 纸型/进纸盒只能临时写进打印机默认值
    with temp_printer_defaults(printer_name, log,
                               paper_size=paper_size, paper_source=paper_source):
        proc = subprocess.run(cmd, capture_output=True, timeout=120)
    if proc.returncode != 0:
        err = (proc.stderr or b"").decode("utf-8", "replace").strip()
        raise RuntimeError(f"SumatraPDF 打印失败: {err or proc.returncode}")
    log(f"SumatraPDF 打印已发送: {os.path.basename(path)}" + (f" (设置: {','.join(settings)})" if settings else ""))


def _range_text(page_set):
    """{1,2,3,5} -> '1-3,5'"""
    rng = []
    pages = sorted(page_set)
    start = prev = pages[0]
    for p in pages[1:]:
        if p == prev + 1:
            prev = p
            continue
        rng.append(f"{start}-{prev}" if prev > start else f"{start}")
        start = prev = p
    rng.append(f"{start}-{prev}" if prev > start else f"{start}")
    return ",".join(rng)


# ---------------------------------------------------------------- 统一打印入口

def settings_to_devmode_values(copies=None, color=None, duplex=None):
    """把 UI 层设置转换为 DEVMODE 值; None 表示不修改(跟随打印机)"""
    c = int(copies) if copies else None
    col = COLOR_MAP.get(color) if color else None
    dup = DUPLEX_MAP.get(duplex) if duplex else None
    return c, col, dup


def to_float(v, default=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def paper_devmode_values(opts):
    """opts 里的纸型/进纸盒 -> (dmPaperSize, dmDefaultSource); None = 跟随打印机"""
    return _paper_size_to_int(opts.get("paper")), _paper_size_to_int(opts.get("source"))


def _print_pdf_path(pdf, printer_name, copies, color, duplex, page_set, log, opts,
                    sumatra_exe, dpi=200):
    """PDF 统一通道: 有 Sumatra 且不逐页 -> Sumatra; 否则走内置渲染引擎"""
    paper_size, paper_source = paper_devmode_values(opts)
    per_page = bool(opts.get("per_page"))
    gap = to_float(opts.get("page_gap"), 0.0)
    if sumatra_exe and not per_page:
        print_pdf_sumatra(sumatra_exe, pdf, printer_name, copies, color, duplex,
                          page_set, log, paper_size=paper_size,
                          paper_source=paper_source)
        return
    try:
        import pypdfium2  # noqa: F401
    except ImportError:
        if sumatra_exe:
            # 逐页打印只能靠内置引擎拆页; 没有它就按整份发一份, 不要因此打不出来
            log("提示: 逐页打印需要内置 PDF 引擎, 本次按整份发送")
            print_pdf_sumatra(sumatra_exe, pdf, printer_name, copies, color, duplex,
                              page_set, log, paper_size=paper_size,
                              paper_source=paper_source)
            return
        raise ImportError("PDF 引擎不可用: 请安装 pypdfium2 或 SumatraPDF")
    print_pdf_gdi(pdf, printer_name, copies, color, duplex, page_set, log, dpi=dpi,
                  paper_size=paper_size, paper_source=paper_source,
                  per_page=per_page, page_gap=gap)


def print_office_com(path, printer_name, copies, color, duplex, page_set,
                     opts, sumatra_exe, log, kind=None):
    """Office 文档: COM 改版面 -> 导出 PDF -> 走 PDF 通道静默打印.

    这一步彻底摆脱了 Shell 的 print 动词: 不会弹程序窗口、不用切系统默认打印机,
    并且能真正控制页边距/缩放 (Word 页边距、Excel 所有列缩到一页宽)。
    """
    layout = opts.get("layout")
    pdf, work = office_core.export_pdf(path, layout, log, kind=kind)
    try:
        # Office 文档文字多, 用 300dpi 渲染, 比默认 200dpi 更清晰
        _print_pdf_path(pdf, printer_name, copies, color, duplex, page_set, log,
                        opts, sumatra_exe, dpi=300)
    finally:
        shutil.rmtree(work, ignore_errors=True)


def print_file(path, printer_name, opts, ext_map, sumatra_exe, log):
    """
    打印单个文件.
    opts: dict(copies=int|None, color='color'|'mono'|None,
               duplex='simplex'|'duplex_long'|'duplex_short'|None,
               pages=str|None, layout=dict|None, office_com=bool,
               paper=int|None 纸型ID(按纸型缩放), source=int|None 进纸盒ID,
               per_page=bool 逐页打印, page_gap=float 页间间隔秒)
    """
    ext = os.path.splitext(path)[1].lower().lstrip(".")
    ext = ext_map.get(ext, ext)
    cat = resolve_category(ext, ext_map)
    if cat is None:
        raise RuntimeError(f"不支持的格式: .{ext} (可在扩展名映射中添加别名)")
    page_set = parse_page_range(opts.get("pages"))
    copies, color, duplex = settings_to_devmode_values(opts.get("copies"), opts.get("color"), opts.get("duplex"))
    paper_size, paper_source = paper_devmode_values(opts)
    per_page = bool(opts.get("per_page"))
    page_gap = to_float(opts.get("page_gap"), 0.0)

    if cat == CAT_RAW:
        if paper_size or paper_source or per_page:
            log("提示: RAW (prn/ps/pcl) 由文件内容自带版面, 纸型/逐页设置对它无效")
        print_raw(path, printer_name, log)
    elif cat == CAT_PDF:
        # 优先内置 PDF 渲染引擎 / Sumatra, 都不行再退回 Shell 动词
        try:
            _print_pdf_path(path, printer_name, copies, color, duplex, page_set, log,
                            opts, sumatra_exe)
        except ImportError:
            _shell_print_with_defaults(path, printer_name, copies, color, duplex, log,
                                       paper_size=paper_size, paper_source=paper_source)
    elif cat in (CAT_TEXT, CAT_IMAGE):
        fn = print_text_gdi if cat == CAT_TEXT else print_image_gdi
        fn(path, printer_name, copies, color, duplex, page_set, log,
           paper_size=paper_size, paper_source=paper_source,
           per_page=per_page, page_gap=page_gap,
           orientation=(opts.get("orientation") if cat == CAT_IMAGE else None))
    else:
        done = False
        if opts.get("office_com", True):
            try:
                print_office_com(path, printer_name, copies, color, duplex,
                                 page_set, opts, sumatra_exe, log,
                                 kind=office_core.ext_kind("." + ext))
                done = True
            except OfficeUnavailable as e:
                log(f"提示: {e} — 改用系统关联程序打印")
            except Exception as e:
                log(f"⚠ Office 静默打印失败: {e} — 退回系统关联程序打印")
        if not done:
            _shell_print_with_defaults(path, printer_name, copies, color, duplex, log,
                                       paper_size=paper_size, paper_source=paper_source,
                                       per_page=per_page)


def _shell_print_with_defaults(path, printer_name, copies, color, duplex, log,
                               paper_size=None, paper_source=None, per_page=False):
    ctx = None
    note = []
    try:
        ctx = apply_printer_defaults(printer_name, copies=copies, color=color,
                                     duplex=duplex, paper_size=paper_size,
                                     paper_source=paper_source)
        note.append("已临时应用份数/色彩/双面/纸型到打印机默认值")
    except PrinterSettingsError as e:
        log(f"提示: 无法修改打印机默认设置 ({e}), 将按打印机当前默认值打印")
    if per_page:
        log("提示: 该格式由系统关联程序打印, 无法逐页拆分成多个作业")
    try:
        # 单文件单独调用时也要保证打到所选打印机 (批量打印时外层已切换, 这里是无操作)
        with default_printer_guard(printer_name, log) as guard:
            shell_print(path, printer_name, log)
            if guard.changed:
                # 给关联程序读取"当前默认打印机"的时间, 再恢复
                time.sleep(1.5)
        if note:
            log("  - " + note[0])
    finally:
        if ctx:
            # 等待关联程序把作业读入后台再恢复默认值
            time.sleep(2.5)
            restore_printer_defaults(*ctx)

