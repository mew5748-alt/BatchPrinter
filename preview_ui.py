# -*- coding: utf-8 -*-
"""
preview_ui.py - 打印预览面板

设计要点:
  1. 内嵌在主窗口里 (不弹窗): 占一列, 拖动分隔条调宽窄, 可整块收起
  2. 单击列表文件立即刷新预览: 转换在后台线程做, 界面不会卡
  3. 预览内容与真正打印用的是同一份 PDF (Office 先按当前版面转 PDF)
  4. 转好的 PDF 按 (文件, 修改时间, 版面) 缓存在临时目录, 来回点选秒开
"""

import os
import json
import queue
import shutil
import hashlib
import tempfile
import threading

from PySide6.QtCore import Qt, QObject, Signal, QTimer
from PySide6.QtGui import QPixmap, QImage
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QCheckBox,
    QScrollArea,
)

import office_core
from printer_core import resolve_category

PREVIEW_CACHE_DIR = os.path.join(tempfile.gettempdir(), "bp_preview_cache")


def _cache_path(path, layout, kind):
    """缓存文件名: 源文件路径 + 修改时间 + 版面参数 (任一变化都会重新生成)"""
    try:
        mt = "%.3f" % os.path.getmtime(path)
    except OSError:
        mt = "0"
    raw = "|".join([os.path.abspath(path).lower(), mt,
                    json.dumps(office_core.normalize_layout(layout), sort_keys=True),
                    str(kind)])
    return os.path.join(PREVIEW_CACHE_DIR,
                        hashlib.md5(raw.encode("utf-8")).hexdigest()[:20] + ".pdf")


def prune_cache(keep=120):
    """清理预览缓存: 只保留最近 keep 个 PDF, 避免临时目录无限增长"""
    try:
        if not os.path.isdir(PREVIEW_CACHE_DIR):
            return 0
        files = []
        for name in os.listdir(PREVIEW_CACHE_DIR):
            p = os.path.join(PREVIEW_CACHE_DIR, name)
            try:
                files.append((os.path.getmtime(p), p))
            except OSError:
                continue
        files.sort(reverse=True)
        removed = 0
        for _mt, p in files[keep:]:
            try:
                os.remove(p)
                removed += 1
            except OSError:
                pass
        return removed
    except Exception:
        return 0


class PreviewSignals(QObject):
    ready = Signal(int, str, str, str)      # seq, kind(pdf|image), 路径, 说明
    failed = Signal(int, str)               # seq, 错误信息


class PreviewLoader:
    """后台生成预览数据 (Office 转 PDF / 读取图片), 单线程串行处理"""

    def __init__(self, signals):
        self.sig = signals
        self._q = queue.Queue()
        self._th = None
        self._lock = threading.Lock()

    # -------------------------------------------------- 生命周期
    def submit(self, seq, path, opts, ext_map):
        self._ensure()
        self._q.put((seq, path, opts or {}, dict(ext_map or {})))

    def stop(self):
        self._q.put(None)
        th = self._th
        if th and th.is_alive():
            th.join(timeout=5.0)
        if not th or not th.is_alive():      # 卡在 COM 里没退出的线程不要再覆盖句柄
            self._th = None

    def _ensure(self):
        with self._lock:
            if self._th and self._th.is_alive():
                return
            self._th = threading.Thread(target=self._loop, name="preview-loader", daemon=True)
            self._th.start()

    # -------------------------------------------------- 线程主体
    def _loop(self):
        while True:
            job = self._q.get()
            if job is None:
                break
            seq, path, opts, ext_map = job
            try:
                kind, payload, note = self._prepare(path, opts, ext_map)
                self.sig.ready.emit(seq, kind, payload, note)
            except Exception as e:                     # 单个文件失败不影响后续
                self.sig.failed.emit(seq, str(e))
        try:
            office_core.close_session()                # 本线程复用的 Word/Excel 关掉
        except Exception:
            pass

    @staticmethod
    def _prepare(path, opts, ext_map):
        ext = os.path.splitext(path)[1].lower().lstrip(".")
        ext = ext_map.get(ext, ext)
        cat = resolve_category(ext, ext_map)
        if cat == "office":
            kind = office_core.ext_kind("." + ext)
            pdf = PreviewLoader.pdf_cached(path, opts.get("layout"), kind)
            return "pdf", pdf, "Office 转 PDF (与真正打印同一份)"
        if cat == "pdf":
            return "pdf", path, "PDF 原文件"
        if cat == "image":
            return "image", path, "图片"
        raise RuntimeError("该格式暂不支持预览 (文本/RAW 由打印机驱动排版, 预览会失真)")

    @staticmethod
    def pdf_cached(path, layout, kind):
        """Office 文档 -> 缓存的 PDF; 已转换过直接复用"""
        os.makedirs(PREVIEW_CACHE_DIR, exist_ok=True)
        out = _cache_path(path, layout, kind)
        if os.path.isfile(out) and os.path.getsize(out) > 0:
            return out
        pdf, work = office_core.export_pdf(path, layout, None, kind=kind)
        try:
            tmp = out + ".part"
            shutil.copyfile(pdf, tmp)
            os.replace(tmp, out)                       # 原子替换, 避免半截文件
        finally:
            shutil.rmtree(work, ignore_errors=True)
        return out


class PreviewPanel(QWidget):
    """打印预览面板 (嵌在主窗口里, 单击列表文件即刷新)

    与独立窗口的区别: 不抢焦点、不用管理位置; 主界面里固定占一列,
    拖动分隔条就能调宽窄, 点工具栏按钮可整块收起。
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("previewPanel")
        self.setMinimumWidth(210)

        self._sig = PreviewSignals()
        self._sig.ready.connect(self._on_ready)
        self._sig.failed.connect(self._on_failed)
        self._loader = PreviewLoader(self._sig)

        self._seq = 0
        self._key = None
        self._pdf = None
        self._img = None
        self._pages = 0
        self._page = 0
        self._zoom = 1.0
        self._note = ""
        self._path = ""
        self._opts = {}
        self._ext_map = {}
        self._printer = ""
        self._dirty = False
        self.print_handler = None            # callable(path, opts) -> bool
        self._build()

        # 拖动分隔条时不要每一帧都重渲染, 停 140ms 再按新宽度出图
        self._resize_timer = QTimer(self)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.setInterval(140)
        self._resize_timer.timeout.connect(self._rerender)

    # ------------------------------------------------------------ 界面
    def _build(self):
        v = QVBoxLayout(self)
        v.setContentsMargins(6, 2, 2, 2)
        v.setSpacing(4)

        self.title_lbl = QLabel("打印预览")
        self.title_lbl.setStyleSheet("font-weight:bold;font-size:12px;color:#333;")
        self.title_lbl.setTextInteractionFlags(Qt.NoTextInteraction)
        v.addWidget(self.title_lbl)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(False)
        self.scroll.setAlignment(Qt.AlignCenter)
        self.scroll.setStyleSheet(
            "QScrollArea{border:1px solid #c8c8c8;background:#e6e6e6;}"
            "QScrollArea > QWidget > QWidget{background:#e6e6e6;}")
        self.canvas = QLabel("单击列表中的文件\n即可在这里预览")
        self.canvas.setAlignment(Qt.AlignCenter)
        self.canvas.setWordWrap(True)
        self.canvas.setStyleSheet("background:#e6e6e6;color:#888;padding:8px;"
                                  "font-size:11px;")
        self.canvas.setMinimumSize(150, 120)
        self.scroll.setWidget(self.canvas)
        v.addWidget(self.scroll, 1)

        nav = QHBoxLayout()
        nav.setSpacing(2)
        self.prev_btn = QPushButton("◀")
        self.prev_btn.setFixedSize(24, 22)
        self.prev_btn.setToolTip("上一页 (或按 ↑↓ 换文件)")
        self.next_btn = QPushButton("▶")
        self.next_btn.setFixedSize(24, 22)
        self.next_btn.setToolTip("下一页")
        self.page_lbl = QLabel("—")
        self.page_lbl.setAlignment(Qt.AlignCenter)
        self.page_lbl.setMinimumWidth(40)
        self.page_lbl.setStyleSheet("font-size:11px;")
        self.prev_btn.clicked.connect(lambda: self._go(-1))
        self.next_btn.clicked.connect(lambda: self._go(1))
        nav.addWidget(self.prev_btn)
        nav.addWidget(self.page_lbl)
        nav.addWidget(self.next_btn)
        nav.addStretch(1)

        fit = QPushButton("⤢")
        fit.setFixedSize(24, 22)
        fit.setToolTip("适合宽度")
        z_out = QPushButton("－")
        z_out.setFixedSize(22, 22)
        z_out.setToolTip("缩小")
        z_in = QPushButton("＋")
        z_in.setFixedSize(22, 22)
        z_in.setToolTip("放大")
        fit.clicked.connect(self._fit)
        z_out.clicked.connect(lambda: self._zoom_by(-0.2))
        z_in.clicked.connect(lambda: self._zoom_by(0.2))
        nav.addWidget(fit)
        nav.addWidget(z_out)
        nav.addWidget(z_in)
        v.addLayout(nav)

        self.info = QLabel("")
        self.info.setStyleSheet("color:#777;font-size:11px;")
        self.info.setMinimumWidth(60)
        self._info_full = ""
        v.addWidget(self.info)

        self.print_btn = QPushButton("🖨 打印此文件")
        self.print_btn.setFixedHeight(26)
        self.print_btn.setToolTip("只打印当前预览的这一个文件")
        self.print_btn.setEnabled(False)
        self.print_btn.clicked.connect(self._print)
        v.addWidget(self.print_btn)
        self._update_nav()

    # ------------------------------------------------------------ 对外接口
    def set_file(self, path, opts, ext_map, printer):
        """显示某个文件的预览 (同一文件 + 同一版面时直接复用, 不重复转换)"""
        self._path = path
        self._opts = dict(opts or {})
        self._ext_map = dict(ext_map or {})
        self._printer = printer or ""
        name = os.path.basename(path)
        self.title_lbl.setText("预览: " + self._elide(name, 28))
        self.title_lbl.setToolTip(path)
        self.print_btn.setEnabled(True)

        key = "%s|%s" % (os.path.abspath(path).lower(), self._layout_key())
        if key == self._key and self._pages:
            self._update_info()
            return
        self._key = key
        self._seq += 1
        self._close_doc()
        self._pages, self._page = 0, 0
        self.canvas.setPixmap(QPixmap())
        self.canvas.setText("正在生成预览…")
        self.canvas.resize(150, 110)
        self._info_full = ""
        self.info.setText("正在转换…")
        self._update_nav()
        self._loader.submit(self._seq, path, self._opts, self._ext_map)

    def refresh(self):
        """版面等设置变化后重新生成 (缓存键带版面, 会自动重转)"""
        if not self._path:
            return
        if not self.isVisible():            # 收起状态先记下来, 展开时再转
            self._dirty = True
            return
        self._dirty = False
        self._key = None
        self.set_file(self._path, self._opts, self._ext_map, self._printer)

    def clear(self):
        self._seq += 1
        self._key = None
        self._path = ""
        self._dirty = False
        self._close_doc()
        self._pages, self._page = 0, 0
        self.title_lbl.setText("打印预览")
        self.title_lbl.setToolTip("")
        self.canvas.setPixmap(QPixmap())
        self.canvas.setText("单击列表中的文件\n即可在这里预览")
        self.canvas.resize(150, 110)
        self._info_full = ""
        self.info.setText("")
        self.info.setToolTip("")
        self.print_btn.setEnabled(False)
        self._update_nav()

    def shutdown(self):
        self._seq += 1
        self._close_doc()
        try:
            self._loader.stop()
        except Exception:
            pass

    # ------------------------------------------------------------ 加载结果
    def _on_ready(self, seq, kind, payload, note):
        if seq != self._seq:                # 已切到别的文件, 丢弃过期结果
            return
        try:
            self._close_doc()
            if kind == "pdf":
                import pypdfium2 as pdfium
                self._pdf = pdfium.PdfDocument(payload)
                self._pages = len(self._pdf)
            else:
                self._img = self._load_image(payload)
                self._pages = 1
            self._note = note
            self._page = 0
            self._zoom = 1.0
            self._update_nav()
            self._show_page(0)
            self._update_info()
        except Exception as e:
            self._on_failed(seq, "无法打开预览: %s" % e)

    def _on_failed(self, seq, msg):
        if seq != self._seq:
            return
        self._close_doc()
        self._pages, self._page = 0, 0
        self.canvas.setPixmap(QPixmap())
        self.canvas.setText("无法预览\n" + self._elide(os.path.basename(self._path), 20))
        self.canvas.resize(150, 110)
        self._info_full = ""
        self.info.setText("不支持预览")
        self.info.setToolTip(msg)
        self._update_nav()

    @staticmethod
    def _load_image(path):
        img = QImage(path)
        if not img.isNull():
            return img
        from PIL import Image as PILImage       # 少见格式 (tiff 多页) 用 Pillow 兜底
        im = PILImage.open(path).convert("RGB")
        return QImage(im.tobytes("raw", "RGB"), im.width, im.height,
                      3 * im.width, QImage.Format_RGB888).copy()

    def _close_doc(self):
        if self._pdf is not None:
            try:
                self._pdf.close()
            except Exception:
                pass
            self._pdf = None
        self._img = None

    # ------------------------------------------------------------ 显示
    def _show_page(self, idx):
        if not (0 <= idx < self._pages):
            return
        self._page = idx
        self._update_nav()
        avail = max(120, self.scroll.viewport().width() - 16)
        try:
            if self._pdf is not None:
                page = self._pdf[idx]
                w, h = page.get_size()
                scale = (avail / float(w)) * self._zoom
                pil = page.render(scale=max(0.05, scale)).to_pil()
                img = QImage(pil.tobytes("raw", "RGB"), pil.width, pil.height,
                             3 * pil.width, QImage.Format_RGB888).copy()
                page.close()
            else:
                img = self._img
                target = max(40, int(avail * self._zoom))
                if img.width() != target:
                    img = img.scaledToWidth(target, Qt.SmoothTransformation)
            self.canvas.setPixmap(QPixmap.fromImage(img))
            self.canvas.resize(img.width() + 8, img.height() + 8)
        except Exception as e:
            self.canvas.setText("渲染第 %d 页失败: %s" % (idx + 1, e))

    def _go(self, delta):
        self._show_page(self._page + delta)

    def _zoom_by(self, delta):
        self._zoom = min(3.0, max(0.3, self._zoom + delta))
        self._show_page(self._page)

    def _fit(self):
        self._zoom = 1.0
        self._show_page(self._page)

    def _update_nav(self):
        n = self._pages
        self.page_lbl.setText("%d/%d" % (self._page + 1, n) if n else "—")
        self.prev_btn.setEnabled(n > 0 and self._page > 0)
        self.next_btn.setEnabled(n > 0 and self._page < n - 1)

    def _update_info(self):
        copies = self._opts.get("copies")
        lay = office_core.layout_summary(self._opts.get("layout")) or "原版面"
        self._info_full = "%d 页 · %s 份 · %s" % (self._pages, copies or "默认", lay)
        self.info.setToolTip("打印机: %s\n%s\n版面: %s"
                             % (self._printer or "—", self._note or "", lay))
        self._elide_info()

    def _elide_info(self):
        if not self._info_full:
            self.info.setText("")
            return
        w = max(40, self.info.width())
        self.info.setText(self.info.fontMetrics().elidedText(
            self._info_full, Qt.ElideRight, w))

    def _layout_key(self):
        return json.dumps(office_core.normalize_layout(self._opts.get("layout")),
                          sort_keys=True)

    @staticmethod
    def _elide(name, limit=28):
        return name if len(name) <= limit else name[:limit - 1] + "…"

    # ------------------------------------------------------------ 交互
    def _print(self):
        if not self._path or not self.print_handler:
            return
        self.print_handler(self._path, dict(self._opts))

    def _rerender(self):
        if self._pages:
            self._show_page(self._page)

    def showEvent(self, e):
        super().showEvent(e)
        if self._dirty:                     # 收起期间改过版面, 展开时补上
            self.refresh()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._elide_info()
        if self._pages and getattr(self, "_resize_timer", None) is not None:
            self._resize_timer.start()


# 兼容旧名字 (v1.2.0 的独立窗口类名)
PreviewWindow = PreviewPanel
