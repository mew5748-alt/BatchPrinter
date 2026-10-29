# -*- coding: utf-8 -*-
"""
office_core.py - Word / Excel / PowerPoint 静默打印与版面控制 (COM 自动化)

为什么不用 Shell 的 print 动词:
    print 动词没有"打印机"和"版面"参数, 行为完全由关联程序决定 (只能打默认打印机,
    也无法改页边距/缩放), 还会弹程序窗口。

这里的做法:
    1. 用 COM 只读打开文档 (原文件不会被修改)
    2. 按需设置版面: 页边距 / 纸张方向 / Excel "所有列缩到一页宽"
    3. 导出为临时 PDF
    4. 交给 printer_core 的 PDF 通道静默打印 (SumatraPDF 或内置 pypdfium2)
    好处: 全程无对话框, 版面精确可控, 打印预览与真正打印用的是同一份 PDF。

只读失败时自动用临时副本重试 (某些文档只读时不允许改版面)。
"""

import os
import shutil
import tempfile
import threading

MM2PT = 72.0 / 25.4          # 毫米 -> 磅 (Word/Excel 的页边距单位是磅)

# ---------------------------------------------------------------- 格式分类

WORD_EXTS = {".doc", ".docx", ".docm", ".dot", ".dotx", ".rtf", ".odt", ".wps", ".wpt"}
EXCEL_EXTS = {".xls", ".xlsx", ".xlsm", ".xlsb", ".et", ".ett", ".ods"}
PPT_EXTS = {".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".dps", ".odp"}

PROGIDS = {"word": "Word.Application",
           "excel": "Excel.Application",
           "ppt": "PowerPoint.Application"}
APP_NAMES = {"word": "Word", "excel": "Excel", "ppt": "PowerPoint"}

# 导出格式常量
WD_FORMAT_PDF = 17
XL_TYPE_PDF = 0
PP_FIXED_FORMAT_PDF = 2
PP_SAVE_AS_PDF = 32

# 方向常量 (两个组件取值不同, 别搞混)
WD_ORIENT_PORTRAIT, WD_ORIENT_LANDSCAPE = 0, 1
XL_ORIENT_PORTRAIT, XL_ORIENT_LANDSCAPE = 1, 2


class OfficeError(Exception):
    """Office 处理失败 (文档本身的问题: 加密/损坏/受保护等)"""


class OfficeUnavailable(OfficeError):
    """没有可用的 Office/WPS, 应回退到系统关联程序打印"""


def ext_kind(path):
    """按扩展名判断由哪个组件处理: word/excel/ppt/None"""
    e = os.path.splitext(str(path))[1].lower()
    if e in WORD_EXTS:
        return "word"
    if e in EXCEL_EXTS:
        return "excel"
    if e in PPT_EXTS:
        return "ppt"
    return None


def is_office_ext(ext):
    """ext 可带点也可不带"""
    e = str(ext).lower()
    if not e.startswith("."):
        e = "." + e
    return e in (WORD_EXTS | EXCEL_EXTS | PPT_EXTS)


# ---------------------------------------------------------------- 版面参数

DEFAULT_LAYOUT = {
    "margins_mm": None,      # (上, 下, 左, 右) 毫米; None = 保持原文档
    "orientation": "",       # "" 原样 / portrait 纵向 / landscape 横向
    "excel_fit": "",         # "" 不缩放 / width 一页宽 / page 一页(宽+高) / percent
    "excel_scale": 0,        # excel_fit=percent 时的百分比 (10-400)
    "center": None,          # Excel 水平居中; None = 不改
    "gridlines": None,       # Excel 打印网格线; None = 不改
}


def normalize_layout(raw):
    """把 UI/配置里的 dict 规范化为内部结构 (容错)"""
    lay = dict(DEFAULT_LAYOUT)
    if not isinstance(raw, dict):
        return lay

    m = raw.get("margins_mm")
    if m:
        try:
            vals = [float(x) for x in (m if isinstance(m, (list, tuple)) else [m] * 4)]
            if len(vals) == 4 and all(0 <= v <= 100 for v in vals):
                lay["margins_mm"] = tuple(round(v, 2) for v in vals)
        except (TypeError, ValueError):
            pass

    o = str(raw.get("orientation") or "").lower()
    if o in ("portrait", "vertical", "纵向"):
        lay["orientation"] = "portrait"
    elif o in ("landscape", "horizontal", "横向"):
        lay["orientation"] = "landscape"

    fit = str(raw.get("excel_fit") or "").lower()
    if fit in ("width", "page", "percent"):
        lay["excel_fit"] = fit
    if lay["excel_fit"] == "percent":
        try:
            s = int(float(raw.get("excel_scale") or 0))
            if 10 <= s <= 400:
                lay["excel_scale"] = s
            else:
                lay["excel_fit"] = ""      # 百分比无效 -> 不缩放
        except (TypeError, ValueError):
            lay["excel_fit"] = ""

    for key in ("center", "gridlines"):
        if isinstance(raw.get(key), bool):
            lay[key] = raw[key]
    return lay


def layout_summary(lay):
    """给日志/界面用的一句话摘要; 无任何设置时返回空串"""
    if not isinstance(lay, dict):
        return ""
    lay = normalize_layout(lay)
    parts = []
    if lay["margins_mm"]:
        t, b, l, r = lay["margins_mm"]
        parts.append(f"页边距 上{t:g}/下{b:g}/左{l:g}/右{r:g}mm")
    if lay["orientation"] == "portrait":
        parts.append("纵向")
    elif lay["orientation"] == "landscape":
        parts.append("横向")
    if lay["excel_fit"] == "width":
        parts.append("列缩到一页宽")
    elif lay["excel_fit"] == "page":
        parts.append("整表缩到一页")
    elif lay["excel_fit"] == "percent":
        parts.append(f"缩放{lay['excel_scale']}%")
    if lay["center"]:
        parts.append("水平居中")
    if lay["gridlines"]:
        parts.append("打印网格线")
    return ", ".join(parts)


# ---------------------------------------------------------------- 组件会话

_tls = threading.local()


class OfficeSession:
    """一个线程内复用同一套 Word/Excel/PPT 进程, 批量打印时省掉反复启动.

    COM 对象有线程亲和性, 所以会话按线程缓存 (_tls), 打印线程与主线程互不干扰。
    """

    def __init__(self, log=None):
        self.log = log or (lambda m: None)
        self._apps = {}

    # -------------------------------------------------- 启动/关闭

    def app(self, kind):
        if kind in self._apps:
            return self._apps[kind]
        app = self._launch(kind)
        self._apps[kind] = app
        return app

    def _launch(self, kind):
        try:
            import pythoncom                      # noqa: F401
            import win32com.client as win32
        except Exception as e:                    # pywin32 缺失
            raise OfficeUnavailable(f"缺少 pywin32 (win32com): {e}")
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass

        prog = PROGIDS[kind]
        try:
            app = win32.DispatchEx(prog)
        except Exception as e:
            raise OfficeUnavailable(
                f"未检测到可用的 {APP_NAMES[kind]} (或 COM 被禁用): {e}")

        # 静默化: 不显示窗口、不弹任何确认框
        try:
            app.DisplayAlerts = 0 if kind == "word" else False
        except Exception:
            pass
        if kind == "word":
            try:
                app.Visible = False
            except Exception:
                pass
            try:
                app.Options.ConfirmConversions = False
            except Exception:
                pass
            try:
                app.Options.UpdateLinksAtOpen = False
            except Exception:
                pass
        elif kind == "excel":
            try:
                app.Visible = False
            except Exception:
                pass
            try:
                app.AutomationSecurity = 3        # 禁用宏
            except Exception:
                pass
            try:
                app.AskToUpdateLinks = False
            except Exception:
                pass
            try:
                app.EnableEvents = False
            except Exception:
                pass
        else:                                     # PowerPoint: 必须可见
            try:
                app.Visible = True
            except Exception:
                pass
            try:
                app.WindowState = 2               # ppWindowMinimized
            except Exception:
                pass
        return app

    def invalidate(self, kind):
        """组件异常/崩溃后丢弃句柄, 下次自动重启"""
        app = self._apps.pop(kind, None)
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass

    def close(self):
        for kind in list(self._apps):
            self.invalidate(kind)


def get_session(log=None):
    """取当前线程的会话 (没有则创建)"""
    sess = getattr(_tls, "session", None)
    if sess is None:
        sess = OfficeSession(log)
        _tls.session = sess
    elif log is not None:
        sess.log = log
    return sess


def close_session():
    """结束当前线程的 Office 进程 (批量打印结束时调用)"""
    sess = getattr(_tls, "session", None)
    if sess is not None:
        sess.close()
        _tls.session = None


# ---------------------------------------------------------------- 版面应用

def _apply_word_layout(doc, lay):
    notes = []
    ps = doc.PageSetup
    if lay["margins_mm"]:
        t, b, l, r = [v * MM2PT for v in lay["margins_mm"]]
        ps.TopMargin, ps.BottomMargin = t, b
        ps.LeftMargin, ps.RightMargin = l, r
        m = lay["margins_mm"]
        notes.append(f"页边距 上{m[0]:g}/下{m[1]:g}/左{m[2]:g}/右{m[3]:g}mm")
    if lay["orientation"] == "portrait":
        ps.Orientation = WD_ORIENT_PORTRAIT
        notes.append("纵向")
    elif lay["orientation"] == "landscape":
        ps.Orientation = WD_ORIENT_LANDSCAPE
        notes.append("横向")
    return notes


def _apply_excel_layout(wb, lay, log):
    notes = []
    applied = set()
    for ws in wb.Worksheets:
        try:
            ps = ws.PageSetup
        except Exception:
            continue                          # 图表工作表等无 PageSetup
        if lay["margins_mm"]:
            t, b, l, r = [v * MM2PT for v in lay["margins_mm"]]
            ps.TopMargin, ps.BottomMargin = t, b
            ps.LeftMargin, ps.RightMargin = l, r
            # 页眉/页脚位置不能超出页边距, 否则 Excel 会报错
            try:
                if float(ps.HeaderMargin) > t:
                    ps.HeaderMargin = t
                if float(ps.FooterMargin) > b:
                    ps.FooterMargin = b
            except Exception:
                pass
            applied.add("margins")
        if lay["orientation"] == "portrait":
            ps.Orientation = XL_ORIENT_PORTRAIT
            applied.add("portrait")
        elif lay["orientation"] == "landscape":
            ps.Orientation = XL_ORIENT_LANDSCAPE
            applied.add("landscape")
        if lay["excel_fit"] == "width":
            ps.Zoom = False                   # 必须先关掉百分比缩放
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = False         # 高度不限
            applied.add("width")
        elif lay["excel_fit"] == "page":
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1
            applied.add("page")
        elif lay["excel_fit"] == "percent" and lay["excel_scale"]:
            ps.Zoom = int(lay["excel_scale"])
            ps.FitToPagesWide = 1             # 复位"调整为"设置
            ps.FitToPagesTall = 1
            applied.add("percent")
        if lay["center"] is not None:
            ps.CenterHorizontally = bool(lay["center"])
            if lay["center"]:
                applied.add("center")
        if lay["gridlines"] is not None:
            ps.PrintGridlines = bool(lay["gridlines"])
            if lay["gridlines"]:
                applied.add("gridlines")

    if "margins" in applied:
        m = lay["margins_mm"]
        notes.append(f"页边距 上{m[0]:g}/下{m[1]:g}/左{m[2]:g}/右{m[3]:g}mm")
    if "portrait" in applied:
        notes.append("纵向")
    if "landscape" in applied:
        notes.append("横向")
    if "width" in applied:
        notes.append("所有列缩到一页宽")
    if "page" in applied:
        notes.append("整表缩到一页")
    if "percent" in applied:
        notes.append(f"缩放{lay['excel_scale']}%")
    if "center" in applied:
        notes.append("水平居中")
    if "gridlines" in applied:
        notes.append("打印网格线")
    return notes


# ---------------------------------------------------------------- 导出 PDF

def _open_word(app, path, readonly):
    return app.Documents.Open(FileName=os.path.abspath(path),
                              ConfirmConversions=False, ReadOnly=readonly,
                              AddToRecentFiles=False, Revert=False)


def _open_excel(app, path, readonly):
    return app.Workbooks.Open(Filename=os.path.abspath(path), UpdateLinks=0,
                              ReadOnly=readonly, IgnoreReadOnlyRecommended=True,
                              AddToMru=False)


def _open_ppt(app, path, readonly):
    return app.Presentations.Open(FileName=os.path.abspath(path),
                                  ReadOnly=readonly, Untitled=False, WithWindow=False)


def _dispatch(kind, app, path, pdf, lay, log):
    """按类型打开 -> 设版面 -> 导出 PDF"""
    if kind == "word":
        doc = _open_word(app, path, True)
        try:
            notes = _apply_word_layout(doc, lay)
            if notes:
                log("  " + "; ".join(notes))
            try:
                doc.ExportAsFixedFormat(pdf, WD_FORMAT_PDF)
            except Exception:
                doc.SaveAs2(pdf, FileFormat=WD_FORMAT_PDF)
        finally:
            try:
                doc.Close(SaveChanges=0)
            except Exception:
                pass
    elif kind == "excel":
        wb = _open_excel(app, path, True)
        try:
            notes = _apply_excel_layout(wb, lay, log)
            if notes:
                log("  " + "; ".join(notes))
            wb.ExportAsFixedFormat(XL_TYPE_PDF, pdf)
        finally:
            try:
                wb.Close(False)
            except Exception:
                pass
    else:
        pres = _open_ppt(app, path, True)
        try:
            try:
                pres.ExportAsFixedFormat(pdf, PP_FIXED_FORMAT_PDF)
            except Exception:
                pres.SaveAs(pdf, PP_SAVE_AS_PDF)
        finally:
            try:
                pres.Close()
            except Exception:
                pass


def _friendly_error(e):
    msg = str(e)
    low = msg.lower()
    if "password" in low or "密码" in msg:
        return "文件已加密/需要密码, 无法自动打印"
    if "0x800a1522" in low or "命令失败" in msg:
        return "文档受保护或已损坏, 无法处理"
    return msg


def export_pdf(path, layout=None, log=None, session=None, out_dir=None, kind=None):
    """把 Office 文档导出为 PDF.

    kind: 强制指定 word/excel/ppt (供扩展名映射后的文件使用); 默认按扩展名判断.
    返回 (pdf_path, work_dir); work_dir 是本次创建/复用的临时目录,
    调用方用完 PDF 后应 shutil.rmtree(work_dir), 传入 out_dir 时不会自动建目录。
    """
    log = log or (lambda m: None)
    kind = kind or ext_kind(path)
    if kind is None:
        raise OfficeError(f"不是 Office 文档: {os.path.basename(path)}")

    lay = normalize_layout(layout)
    sess = session or get_session(log)
    if session is not None:
        session.log = log
    own_dir = out_dir is None
    out_dir = out_dir or tempfile.mkdtemp(prefix="bp_office_")
    pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0] + ".pdf")
    if os.path.exists(pdf):
        pdf = os.path.join(out_dir, "office_%d.pdf" % (abs(hash(path)) % 100000))

    try:
        try:
            _dispatch(kind, sess.app(kind), path, pdf, lay, log)
        except OfficeUnavailable:
            raise
        except Exception as e:
            # 只读受限制/组件卡住: 换临时副本重试一次
            log(f"  ⚠ 只读处理失败 ({_friendly_error(e)}), 改用临时副本重试…")
            sess.invalidate(kind)
            tmpd = tempfile.mkdtemp(prefix="bp_office_copy_")
            try:
                copy = os.path.join(tmpd, os.path.basename(path))
                shutil.copy2(path, copy)
                _dispatch(kind, sess.app(kind), copy, pdf, lay, log)
            finally:
                shutil.rmtree(tmpd, ignore_errors=True)

        if not os.path.isfile(pdf) or os.path.getsize(pdf) == 0:
            raise OfficeError("导出 PDF 失败 (没有生成文件)")
    except OfficeError:
        if own_dir:
            shutil.rmtree(out_dir, ignore_errors=True)
        raise
    except Exception as e:
        if own_dir:
            shutil.rmtree(out_dir, ignore_errors=True)
        raise OfficeError(_friendly_error(e))
    return pdf, out_dir


# ---------------------------------------------------------------- 引擎探测 (给界面显示)

def detect_engine():
    """粗查本机 Word/Excel 由谁提供: 返回 (标签, 说明)"""
    try:
        import winreg
    except Exception:
        return "", ""
    found = {}
    for exe, key in (("WINWORD.EXE", "word"), ("EXCEL.EXE", "excel")):
        p = ""
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
                try:
                    with winreg.OpenKey(
                            root, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\%s" % exe,
                            0, winreg.KEY_READ | view) as k:
                        p = winreg.QueryValueEx(k, "")[0]
                        break
                except Exception:
                    continue
            if p:
                break
        if p:
            found[key] = p
    if not found:
        return "", ""
    low = " ".join(found.values()).lower()
    if "wps" in low or "kingsoft" in low:
        return "WPS Office", "由 WPS 提供 COM 能力"
    return "Microsoft Office", "由 Office 提供 COM 能力"
