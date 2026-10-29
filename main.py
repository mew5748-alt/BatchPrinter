# -*- coding: utf-8 -*-
"""
批量打印助手 (Batch Printer) - Windows 批量打印软件
- 添加文件 / 文件夹(含子文件夹) / 压缩包(zip/7z/rar/tar, 自动解压, 无需事先解压)
- 全局设置 + 每文件单独设置: 份数 / 色彩模式 / 双面 / 页码范围
- 文件扩展名映射 (非标准扩展名 -> 受支持格式)
- 支持拖拽文件/文件夹到窗口
运行: python main.py   (依赖: pywin32, pypdfium2, py7zr, rarfile, Pillow, PySide6)
"""

import os
import re
import sys
import json
import ctypes
import queue
import threading
import tempfile
import shutil
from functools import cmp_to_key

from PySide6.QtCore import Qt, QObject, Signal, QSize, QTimer
from PySide6.QtGui import QColor, QAction, QPixmap, QImage, QPainter, QFont
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QComboBox, QLabel, QTableWidget, QTableWidgetItem, QSpinBox,
    QCheckBox, QLineEdit, QTextEdit, QFileDialog, QMessageBox, QDialog,
    QHeaderView, QAbstractItemView, QGroupBox, QSplitter, QToolBar, QStyle,
    QTableWidgetSelectionRange, QMenu, QDoubleSpinBox, QScrollArea, QListWidget,
    QSizePolicy, QFrame,
)

import office_core
from preview_ui import PreviewPanel, prune_cache
from printer_core import (
    list_printers, get_default_printer, find_sumatra, print_file,
    resolve_category, all_known_exts, parse_page_range,
    default_printer_guard, get_default_printer_name,
    list_paper_sizes, list_paper_bins,
)
from archive_util import extract_archive, cleanup_dir, is_archive

APP_NAME = "BatchPrinter"
PREVIEW_MIN_W = 240          # 预览列最窄
PREVIEW_DEFAULT_W = 330      # 默认宽度
PREVIEW_MAX_W = 620          # 最宽
APP_VERSION = "1.4.0"

APP_DIR = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
    else os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "config.json")


# ---------------------------------------------------------------- 文件排序
# 打印顺序 = 列表顺序。文件夹/压缩包内文件按 Windows 资源管理器顺序排列:
# 数字段按数值比较 (1, 2, 10 而非 1, 10, 2), 不区分大小写。
try:
    _StrCmpLogicalW = ctypes.windll.shlwapi.StrCmpLogicalW
    _StrCmpLogicalW.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p]
    _StrCmpLogicalW.restype = ctypes.c_int
except Exception:                                     # 非 Windows 环境兜底
    _StrCmpLogicalW = None


def _natural_key(name):
    """纯 Python 自然排序键 (StrCmpLogicalW 不可用时使用)"""
    parts = re.split(r"(\d+)", name.lower())
    return [(0, int(p), "") if p.isdigit() else (1, 0, p) for p in parts if p != ""]


def natural_cmp(a, b):
    """比较两个名称, 返回 -1/0/1"""
    if _StrCmpLogicalW is not None:
        try:
            return _StrCmpLogicalW(a, b)
        except Exception:
            pass
    ka, kb = _natural_key(a), _natural_key(b)
    return (ka > kb) - (ka < kb)


def sort_names(names):
    """按资源管理器顺序排序文件名列表"""
    return sorted(names, key=cmp_to_key(natural_cmp))

DEFAULT_LAYOUT_CFG = {
    "margins_on": False, "top": 20.0, "bottom": 20.0, "left": 20.0, "right": 20.0,
    "orientation": "", "excel_fit": "", "excel_scale": 100,
    "center": False, "gridlines": False,
}

DEFAULT_CONFIG = {
    "printer": "",
    "global": {"copies": 1, "color": "", "duplex": "", "pages": "",
               "paper": 0, "source": 0, "per_page": False, "page_gap": 0.5},
    "layout": dict(DEFAULT_LAYOUT_CFG),
    "office_com": True,
    "subfolders": True,
    "interval": 2.0,
    "sumatra_path": "",
    "ext_map": {},
    "preview_show": True,       # 主界面里是否显示预览面板
    "preview_width": 330,       # 预览面板宽度 (像素)
}

COLOR_LABELS = {"": "跟随打印机", "color": "彩色", "mono": "黑白"}
DUPLEX_LABELS = {"": "跟随打印机", "simplex": "单面",
                 "duplex_long": "双面(长边)", "duplex_short": "双面(短边)"}
COLOR_LABEL2VAL = {v: k for k, v in COLOR_LABELS.items()}
DUPLEX_LABEL2VAL = {v: k for k, v in DUPLEX_LABELS.items()}

STATUS_QUEUED = "等待中"
C_SEQ, C_NAME, C_TYPE, C_COPIES, C_COLOR, C_DUPLEX, C_PAGES, C_STATUS = range(8)
COLS = ["#", "文件名", "类型", "份数", "色彩", "双面", "页码范围", "状态"]

# 排序方式: 显示名 -> 内部键
SORT_MODES = {
    "添加顺序": "added",
    "文件名 (A→Z)": "name",
    "修改时间 (早→晚)": "mtime_asc",
    "修改时间 (晚→早)": "mtime_desc",
}


def file_mtime(path):
    """文件修改时间 (取不到时返回 0, 保证可排序)"""
    try:
        return os.path.getmtime(path)
    except Exception:
        return 0
STATUS_COLORS = {"已发送": "#1c7c2e", "失败": "#c0392b", "打印中": "#1a5fb4", "已取消": "#888"}


def load_config():
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    if os.path.isfile(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                saved = json.load(f)
            for k, v in saved.items():
                if k == "global" and isinstance(v, dict):
                    cfg["global"].update(v)
                elif k == "layout" and isinstance(v, dict):
                    cfg["layout"].update(v)
                else:
                    cfg[k] = v
        except Exception:
            pass
    return cfg


class WorkerSignals(QObject):
    log = Signal(str)
    status = Signal(int, str)
    done = Signal()


class ReorderTable(QTableWidget):
    """支持拖动行调整顺序的表格 (拖动即改变打印顺序)"""

    _INTERNAL = "application/x-qabstractitemmodeldatalist"

    def __init__(self, rows, cols, on_reorder, parent=None):
        super().__init__(rows, cols, parent)
        self._on_reorder = on_reorder

    def _is_file_drop(self, ev):
        """外部文件拖入 -> 交回主窗口处理, 不走行排序"""
        md = ev.mimeData()
        return md.hasUrls() and not md.hasFormat(self._INTERNAL)

    def dragEnterEvent(self, ev):
        if self._is_file_drop(ev):
            ev.ignore()
            return
        super().dragEnterEvent(ev)

    def dragMoveEvent(self, ev):
        if self._is_file_drop(ev):
            ev.ignore()
            return
        super().dragMoveEvent(ev)

    def dropEvent(self, ev):
        if self._is_file_drop(ev):
            ev.ignore()
            return
        pos = ev.position().toPoint() if hasattr(ev, "position") else ev.pos()
        drop_row = self.rowAt(pos.y())
        ind = self.dropIndicatorPosition()
        if drop_row < 0:
            drop_row = self.rowCount()
        elif ind == QAbstractItemView.BelowItem:
            drop_row += 1
        src = sorted({i.row() for i in self.selectedIndexes()})
        if not src:
            ev.ignore()
            return
        self._on_reorder(src, drop_row)
        ev.setDropAction(Qt.MoveAction)
        ev.accept()


FIT_LABELS = {"": "不缩放 (按原表)", "width": "所有列缩到一页宽",
              "page": "整表缩到一页 (宽+高)", "percent": "按百分比缩放"}
FIT_LABEL2KEY = {v: k for k, v in FIT_LABELS.items()}
ORIENT_LABELS = {"": "保持原样", "portrait": "纵向", "landscape": "横向"}
ORIENT_LABEL2KEY = {v: k for k, v in ORIENT_LABELS.items()}


class LayoutDialog(QDialog):
    """Word/Excel 版面设置 (只作用于打印, 不修改原文件)"""

    def __init__(self, parent, layout, office_com=True):
        super().__init__(parent)
        self.setWindowTitle("版面设置 (只影响打印, 不改原文件)")
        self.setModal(True)
        self.setMinimumWidth(450)
        lay = office_core.normalize_layout(layout)
        v = QVBoxLayout(self)
        v.setSpacing(8)

        # 页边距
        m = QGroupBox(" 页边距 (Word 与 Excel 通用) ")
        mg = QGridLayout(m)
        self.margins_chk = QCheckBox("自定义页边距")
        mm = lay["margins_mm"] or (20.0, 20.0, 20.0, 20.0)
        self.spins = {}
        for i, (key, label) in enumerate((("top", "上"), ("bottom", "下"),
                                          ("left", "左"), ("right", "右"))):
            sp = QDoubleSpinBox()
            sp.setRange(0, 100)
            sp.setDecimals(1)
            sp.setSingleStep(1)
            sp.setSuffix(" mm")
            sp.setValue(float(mm[i]))
            row, col = divmod(i, 2)
            mg.addWidget(QLabel(label + ":"), row, col * 2)
            mg.addWidget(sp, row, col * 2 + 1)
            self.spins[key] = sp
        mg.addWidget(self.margins_chk, 2, 0, 1, 4)
        self.margins_chk.setChecked(bool(lay["margins_mm"]))
        v.addWidget(m)

        # 方向
        o = QGroupBox(" 纸张方向 ")
        og = QHBoxLayout(o)
        self.orient_c = QComboBox()
        self.orient_c.addItems(list(ORIENT_LABELS.values()))
        self.orient_c.setCurrentText(ORIENT_LABELS.get(lay["orientation"], "保持原样"))
        og.addWidget(self.orient_c)
        og.addStretch()
        v.addWidget(o)

        # Excel
        x = QGroupBox(" Excel 排版 ")
        xg = QGridLayout(x)
        xg.addWidget(QLabel("缩放方式:"), 0, 0)
        self.fit_c = QComboBox()
        self.fit_c.addItems(list(FIT_LABELS.values()))
        self.fit_c.setCurrentText(FIT_LABELS.get(lay["excel_fit"], FIT_LABELS[""]))
        self.fit_c.currentTextChanged.connect(self._sync_enabled)
        xg.addWidget(self.fit_c, 0, 1)
        xg.addWidget(QLabel("缩放比例:"), 1, 0)
        self.scale_sp = QSpinBox()
        self.scale_sp.setRange(10, 400)
        self.scale_sp.setSuffix(" %")
        self.scale_sp.setValue(int(lay["excel_scale"] or 100))
        xg.addWidget(self.scale_sp, 1, 1)
        self.center_chk = QCheckBox("水平居中")
        self.center_chk.setChecked(bool(lay["center"]))
        xg.addWidget(self.center_chk, 2, 0, 1, 2)
        self.grid_chk = QCheckBox("打印网格线")
        self.grid_chk.setChecked(bool(lay["gridlines"]))
        xg.addWidget(self.grid_chk, 3, 0, 1, 2)
        v.addWidget(x)

        self.office_chk = QCheckBox("Office 文件用静默打印 (不弹程序窗口)")
        self.office_chk.setChecked(bool(office_com))
        self.office_chk.setToolTip("关闭后回到系统关联程序打印: 会弹程序窗口, 且无法控制"
                                   "页边距/缩放, 也不保证打到所选打印机")
        v.addWidget(self.office_chk)

        tip = QLabel("提示: 这些设置只对打印生效, 不会修改原文件。\n"
                     "Word 支持页边距/方向; Excel 还支持缩放、居中、网格线。")
        tip.setStyleSheet("color:#888;font-size:11px;")
        tip.setWordWrap(True)
        v.addWidget(tip)

        hb = QHBoxLayout()
        reset = QPushButton("恢复默认")
        ok = QPushButton("确定")
        ok.setDefault(True)
        cancel = QPushButton("取消")
        hb.addWidget(reset)
        hb.addStretch()
        hb.addWidget(ok)
        hb.addWidget(cancel)
        v.addLayout(hb)

        self.margins_chk.toggled.connect(self._sync_enabled)
        reset.clicked.connect(self._reset)
        ok.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)
        self._sync_enabled()

    def _sync_enabled(self, *_):
        on = self.margins_chk.isChecked()
        for sp in self.spins.values():
            sp.setEnabled(on)
        self.scale_sp.setEnabled(self.fit_c.currentText() == FIT_LABELS["percent"])

    def _reset(self):
        self.margins_chk.setChecked(False)
        for sp in self.spins.values():
            sp.setValue(20.0)
        self.orient_c.setCurrentText(ORIENT_LABELS[""])
        self.fit_c.setCurrentText(FIT_LABELS[""])
        self.scale_sp.setValue(100)
        self.center_chk.setChecked(False)
        self.grid_chk.setChecked(False)
        self.office_chk.setChecked(True)
        self._sync_enabled()

    def result_cfg(self):
        return {
            "margins_on": self.margins_chk.isChecked(),
            "top": self.spins["top"].value(),
            "bottom": self.spins["bottom"].value(),
            "left": self.spins["left"].value(),
            "right": self.spins["right"].value(),
            "orientation": ORIENT_LABEL2KEY.get(self.orient_c.currentText(), ""),
            "excel_fit": FIT_LABEL2KEY.get(self.fit_c.currentText(), ""),
            "excel_scale": self.scale_sp.value(),
            "center": self.center_chk.isChecked(),
            "gridlines": self.grid_chk.isChecked(),
        }


class BatchPrinterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"批量打印助手 v{APP_VERSION} - Batch Printer")
        self.resize(1360, 800)
        scr = QApplication.primaryScreen()
        if scr is not None:                       # 小屏幕下别超出桌面
            av = scr.availableGeometry()
            self.resize(min(self.width(), max(1000, av.width() - 60)),
                        min(self.height(), max(620, av.height() - 80)))

        self.cfg = load_config()
        self.items = []
        self.temp_roots = set()
        self._seq = 0
        self._added_seq = 0
        self._last_sort_label = "添加顺序"
        self.print_thread = None
        self.cancel_event = threading.Event()
        self._preview_row = -1
        self.sig = WorkerSignals()
        self.sig.log.connect(self._log)
        self.sig.status.connect(self._set_status)
        self.sig.done.connect(self._on_done)

        self._build_ui()
        self._init_paper_ui()
        self.refresh_printers()
        self._refresh_sumatra_label()
        self._sync_layout_ui()
        self._refresh_engine_label()
        self.setAcceptDrops(True)

        # 单击列表某行 -> 延迟 120ms 再生成预览 (连续点选时只预览最后一个)
        self._preview_timer = QTimer(self)
        self._preview_timer.setSingleShot(True)
        self._preview_timer.setInterval(120)
        self._preview_timer.timeout.connect(self._preview_current)
        self.table.currentCellChanged.connect(self._on_current_changed)

    # ------------------------------------------------------------ UI 构建

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)

        # 第一行: 添加/移除/排序
        bar = QHBoxLayout()
        btns = [
            ("📁 添加文件", self.add_files),
            ("📂 添加文件夹", self.add_folder),
            ("🗜 添加压缩包", self.add_archives),
            ("移除选中", self.remove_selected),
            ("清空列表", self.clear_list),
            ("🔗 扩展名映射", self.edit_extmap),
        ]
        for i, (t, fn) in enumerate(btns):
            b = QPushButton(t)
            b.clicked.connect(fn)
            bar.addWidget(b)
            if i == 2:
                bar.addSpacing(14)
        bar.addStretch(1)
        bar.addWidget(QLabel("排序:"))
        self.sort_combo = QComboBox()
        self.sort_combo.addItems(list(SORT_MODES.keys()))
        self.sort_combo.setMinimumWidth(130)
        self.sort_combo.setToolTip("对整个列表重排 (列表顺序 = 打印顺序)")
        self.sort_combo.currentTextChanged.connect(self._on_sort_changed)
        bar.addWidget(self.sort_combo)
        root.addLayout(bar)

        # 第二行: 打印机 + 列表顺序/批量设置 + 预览开关
        prow = QHBoxLayout()
        prow.addWidget(QLabel("打印机:"))
        self.printer_combo = QComboBox()
        self.printer_combo.setMinimumWidth(180)
        self.printer_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        prow.addWidget(self.printer_combo, 1)
        refresh = QPushButton("刷新")
        refresh.setFixedWidth(56)
        refresh.clicked.connect(self.refresh_printers)
        prow.addWidget(refresh)
        self.sumatra_lbl = QLabel("")
        prow.addWidget(self.sumatra_lbl)
        prow.addSpacing(10)
        for t, fn, tip in (("↑ 上移", lambda: self.move_selected(-1), "把选中行上移一位"),
                           ("↓ 下移", lambda: self.move_selected(1), "把选中行下移一位"),
                           ("⚙ 批量设置", lambda: self.batch_settings(),
                            "对选中的多个文件统一设置份数/色彩/双面/页码范围 (Ctrl/Shift 多选)")):
            b = QPushButton(t)
            b.setToolTip(tip)
            b.clicked.connect(fn)
            prow.addWidget(b)
        # 预览面板开关 (单击文件会自动预览)
        self.preview_btn = QPushButton("👁 打印预览")
        self.preview_btn.setCheckable(True)
        self.preview_btn.setToolTip("显示/隐藏预览面板 (单击列表中的文件即自动预览)\n"
                                    "预览与真正打印完全一致, 不耗纸")
        self.preview_btn.toggled.connect(self.toggle_preview_window)
        prow.addWidget(self.preview_btn)
        root.addLayout(prow)

        # 主体: 表格 + 设置
        splitter = QSplitter(Qt.Horizontal)
        root.addWidget(splitter, 1)

        self.table = ReorderTable(0, len(COLS), self._apply_reorder)
        self.table.setHorizontalHeaderLabels(COLS)
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(C_NAME, QHeaderView.Stretch)
        for c in (C_STATUS, C_COPIES, C_TYPE, C_SEQ):
            hh.setSectionResizeMode(c, QHeaderView.ResizeToContents)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.doubleClicked.connect(self._on_double_click)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._table_menu)
        # 拖动行调整打印顺序
        self.table.setDragEnabled(True)
        self.table.setAcceptDrops(True)
        self.table.setDropIndicatorShown(True)
        self.table.setDragDropMode(QAbstractItemView.InternalMove)
        self.table.setDefaultDropAction(Qt.MoveAction)
        self.table.setDragDropOverwriteMode(False)
        left = QWidget()
        lh = QHBoxLayout(left)
        lh.setContentsMargins(0, 0, 0, 0)
        lh.setSpacing(6)

        table_area = QWidget()
        lv0 = QVBoxLayout(table_area)
        lv0.setContentsMargins(0, 0, 0, 0)
        lv0.setSpacing(4)
        lv0.addWidget(self.table, 1)
        self.drag_hint = QLabel("💡 单击文件即在右侧预览; 列表顺序 = 打印顺序 (可拖动行调整); "
                                "双击设置单个文件")
        self.drag_hint.setWordWrap(True)
        self.drag_hint.setStyleSheet("color:#666;font-size:11px;")
        lv0.addWidget(self.drag_hint)
        lh.addWidget(table_area, 1)

        # ---- 预览面板 (紧挨列表右侧, 固定宽度, 可整块收起)
        self._build_preview_pane()
        vline = QFrame()
        vline.setFrameShape(QFrame.VLine)
        vline.setFrameShadow(QFrame.Sunken)
        lh.addWidget(vline)
        lh.addWidget(self.preview_win, 0)

        splitter.addWidget(left)

        right = QWidget()
        right.setMinimumWidth(310)
        rv = QVBoxLayout(right)
        rv.setContentsMargins(0, 0, 0, 0)
        rv.setSpacing(6)

        # 设置项放进滚动区: 小屏幕 / 高 DPI 下按钮永远够得着, 不会被裁掉
        self.settings_scroll = QScrollArea()
        self.settings_scroll.setWidgetResizable(True)
        self.settings_scroll.setFrameShape(QFrame.NoFrame)
        self.settings_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.settings_scroll.setStyleSheet("QScrollArea{background:transparent;}")
        settings_host = QWidget()
        rv_inner = QVBoxLayout(settings_host)
        rv_inner.setContentsMargins(0, 0, 2, 0)
        rv_inner.setSpacing(6)
        self.settings_scroll.setWidget(settings_host)

        # ---- 打印参数
        g1 = QGroupBox(" 打印参数 (作用于整个列表) ")
        g = QGridLayout(g1)
        g.setVerticalSpacing(6)

        g.addWidget(QLabel("份数:"), 0, 0)
        self.g_copies = QSpinBox()
        self.g_copies.setRange(1, 999)
        self.g_copies.setValue(int(self.cfg["global"].get("copies", 1)))
        g.addWidget(self.g_copies, 0, 1)

        g.addWidget(QLabel("色彩模式:"), 1, 0)
        self.g_color = QComboBox()
        self.g_color.addItems(list(COLOR_LABELS.values()))
        self.g_color.setCurrentText(COLOR_LABELS.get(self.cfg["global"].get("color", ""), "跟随打印机"))
        g.addWidget(self.g_color, 1, 1)

        g.addWidget(QLabel("双面/单面:"), 2, 0)
        self.g_duplex = QComboBox()
        self.g_duplex.addItems(list(DUPLEX_LABELS.values()))
        self.g_duplex.setCurrentText(DUPLEX_LABELS.get(self.cfg["global"].get("duplex", ""), "跟随打印机"))
        g.addWidget(self.g_duplex, 2, 1)

        g.addWidget(QLabel("页码范围:"), 3, 0)
        self.g_pages = QLineEdit(self.cfg["global"].get("pages", ""))
        self.g_pages.setPlaceholderText("如 1-3,5 留空=全部")
        g.addWidget(self.g_pages, 3, 1)

        # 纸型 (按纸型缩放) / 纸张来源 (进纸盒), 选项来自打印机驱动
        g.addWidget(QLabel("纸型:"), 4, 0)
        self.g_paper = QComboBox()
        self.g_paper.setToolTip("按纸型缩放: 内容自动缩放到该纸型的可打印区域\n"
                                "例如 A4 的文档选 A3 → 放大铺满 A3; 选 A5 → 缩到 A5\n"
                                "选项由打印机驱动提供, 留空/跟随 = 用打印机当前纸型")
        g.addWidget(self.g_paper, 4, 1)

        g.addWidget(QLabel("纸张来源:"), 5, 0)
        self.g_source = QComboBox()
        self.g_source.setToolTip("从哪个进纸盒/手送盘取纸 (纸盘1、纸盘2、手送纸盘…)\n"
                                 "选项由打印机驱动提供, 跟随打印机 = 用驱动当前设置")
        g.addWidget(self.g_source, 5, 1)

        # 逐页打印: 每页一个独立作业, 页与页之间可停一会儿
        self.g_per_page = QCheckBox("逐页打印")
        self.g_per_page.setToolTip(
            "把每一页单独作为一个打印作业发送, 而不是整份一次发完。\n"
            "适合: 手动放纸/翻面、逐页核对、卡纸后按页补打、想让打印机一页一页出。\n"
            "代价: 页数多时打印速度略慢 (每页都要和打印机交互一次)")
        g.addWidget(self.g_per_page, 6, 0)

        self.g_gap = QDoubleSpinBox()
        self.g_gap.setRange(0.0, 60.0)
        self.g_gap.setSingleStep(0.5)
        self.g_gap.setDecimals(1)
        self.g_gap.setSuffix(" 秒/页")
        self.g_gap.setToolTip("逐页打印时, 每页之间等待多久再发下一页\n"
                              "0 = 不等待 (连续发); 手动翻纸时可设 3~5 秒")
        g.addWidget(self.g_gap, 6, 1)
        rv_inner.addWidget(g1)

        # ---- 版面设置 (只影响打印, 不改原文件)
        g2 = QGroupBox(" 版面设置 (Word / Excel) ")
        lv2 = QVBoxLayout(g2)
        lv2.setSpacing(5)
        self.layout_lbl = QLabel("")
        self.layout_lbl.setWordWrap(True)
        self.layout_lbl.setStyleSheet("color:#555;font-size:11px;")
        lv2.addWidget(self.layout_lbl)
        self.fit_width_chk = QCheckBox("Excel: 所有列缩到一页宽")
        self.fit_width_chk.setToolTip("打开后, 每张工作表都按纸宽缩放到 1 页 (常用)")
        self.fit_width_chk.toggled.connect(self._on_fit_width_toggled)
        lv2.addWidget(self.fit_width_chk)
        layout_btn = QPushButton("📐 更多版面设置 (页边距/方向/比例)…")
        layout_btn.clicked.connect(self.edit_layout)
        lv2.addWidget(layout_btn)
        self.engine_lbl = QLabel("")
        self.engine_lbl.setStyleSheet("color:#888;font-size:11px;")
        self.engine_lbl.setWordWrap(True)
        lv2.addWidget(self.engine_lbl)
        rv_inner.addWidget(g2)

        # ---- 其他
        g3 = QGroupBox(" 其他 ")
        g3g = QGridLayout(g3)
        g3g.setVerticalSpacing(6)
        self.subfolder_chk = QCheckBox("添加文件夹时包含子文件夹")
        self.subfolder_chk.setChecked(bool(self.cfg.get("subfolders", True)))
        g3g.addWidget(self.subfolder_chk, 0, 0, 1, 2)

        g3g.addWidget(QLabel("打印间隔(秒):"), 1, 0)
        self.interval_edit = QLineEdit(str(self.cfg.get("interval", 2.0)))
        g3g.addWidget(self.interval_edit, 1, 1)

        g3g.addWidget(QLabel("SumatraPDF:"), 2, 0)
        sw = QHBoxLayout()
        self.sumatra_edit = QLineEdit(self.cfg.get("sumatra_path", ""))
        sw.addWidget(self.sumatra_edit)
        browse = QPushButton("…")
        browse.setFixedWidth(28)
        browse.clicked.connect(self._browse_sumatra)
        sw.addWidget(browse)
        g3g.addLayout(sw, 2, 1)
        hint = QLabel("推荐安装, 留空自动搜索")
        hint.setStyleSheet("color:#888;font-size:11px;")
        g3g.addWidget(hint, 3, 0, 1, 2)
        rv_inner.addWidget(g3)

        rv_inner.addStretch(1)
        rv.addWidget(self.settings_scroll, 1)
        self.print_btn = QPushButton("🖨 立即打印全部")
        self.print_btn.setMinimumHeight(40)
        self.print_btn.setStyleSheet("font-weight:bold;font-size:14px;")
        self.print_btn.clicked.connect(self.start_print)
        rv.addWidget(self.print_btn)
        self.stop_btn = QPushButton("停止")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_print)
        rv.addWidget(self.stop_btn)
        splitter.addWidget(right)
        self.splitter = splitter
        splitter.setStretchFactor(0, 1)     # 表格区吃掉多余宽度, 设置栏保持固定
        splitter.setStretchFactor(1, 0)
        splitter.setChildrenCollapsible(False)
        splitter.setCollapsible(0, False)
        splitter.setCollapsible(1, False)
        splitter.setHandleWidth(5)
        self._apply_preview_width()

        self.preview_btn.blockSignals(True)
        self.preview_btn.setChecked(bool(self.cfg.get("preview_show", True)))
        self.preview_btn.blockSignals(False)

        # 日志
        logg = QGroupBox(" 日志 ")
        lv = QVBoxLayout(logg)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(100)
        self.log_text.setStyleSheet("font-family:Consolas,monospace;font-size:12px;")
        lv.addWidget(self.log_text)
        root.addWidget(logg)

    # ------------------------------------------------------------ 拖拽

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        n0 = len(self.items)
        for url in e.mimeData().urls():
            p = url.toLocalFile()
            if p:
                self.enqueue_path(p, recursive=self.subfolder_chk.isChecked())
        self._log(f"拖入并添加 {len(self.items) - n0} 个文件")

    # ------------------------------------------------------------ 配置

    def save_config(self):
        self.cfg["printer"] = self.current_printer()
        self.cfg["global"] = {
            "copies": self.g_copies.value(),
            "color": COLOR_LABEL2VAL.get(self.g_color.currentText(), ""),
            "duplex": DUPLEX_LABEL2VAL.get(self.g_duplex.currentText(), ""),
            "pages": self.g_pages.text().strip(),
            "paper": self.current_paper()[0],
            "source": self.current_source()[0],
            "per_page": self.g_per_page.isChecked(),
            "page_gap": round(float(self.g_gap.value()), 2),
        }
        self.cfg["subfolders"] = self.subfolder_chk.isChecked()
        self.cfg["layout"] = dict(self.cfg.get("layout") or DEFAULT_LAYOUT_CFG)
        try:
            self.cfg["interval"] = max(0.0, float(self.interval_edit.text()))
        except ValueError:
            self.cfg["interval"] = 2.0
        self.cfg["sumatra_path"] = self.sumatra_edit.text().strip()
        pw = getattr(self, "preview_win", None)
        if pw is not None and pw.isVisible() and pw.width() > 60:
            self.cfg["preview_width"] = pw.width()     # 记住用户拖出来的宽度
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(self.cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    @staticmethod
    def _to_float(s, d):
        try:
            return float(s)
        except Exception:
            return d

    def _log(self, msg):
        self.log_text.append(msg)

    def _refresh_sumatra_label(self):
        exe = find_sumatra(self.sumatra_edit.text().strip())
        if exe:
            self.sumatra_lbl.setText("✔ SumatraPDF")
            self.sumatra_lbl.setToolTip("已检测到 SumatraPDF, PDF 静默打印效果最佳")
            self.sumatra_lbl.setStyleSheet("color:#1c7c2e;")
        else:
            self.sumatra_lbl.setText("PDF: 内置引擎")
            self.sumatra_lbl.setToolTip("未安装 SumatraPDF, 使用程序内置渲染引擎\n"
                                        "安装后可获得更精确的 PDF 打印控制")
            self.sumatra_lbl.setStyleSheet("color:#0a7d32;")

    def _browse_sumatra(self):
        p, _ = QFileDialog.getOpenFileName(self, "选择 SumatraPDF.exe", "",
                                           "SumatraPDF.exe (SumatraPDF.exe);;所有文件 (*.*)")
        if p:
            self.sumatra_edit.setText(p)
            self._refresh_sumatra_label()

    # ------------------------------------------------------------ 版面设置

    def layout_opts(self):
        """配置里的版面设置 -> office_core 用的 layout dict"""
        l = self.cfg.get("layout") or {}
        lay = {
            "margins_mm": None,
            "orientation": l.get("orientation", ""),
            "excel_fit": l.get("excel_fit", ""),
            "excel_scale": l.get("excel_scale", 100),
            "center": bool(l.get("center", False)),
            "gridlines": bool(l.get("gridlines", False)),
        }
        if l.get("margins_on"):
            lay["margins_mm"] = [l.get("top", 20), l.get("bottom", 20),
                                 l.get("left", 20), l.get("right", 20)]
        return lay

    def _sync_layout_ui(self):
        lay = self.cfg.setdefault("layout", dict(DEFAULT_LAYOUT_CFG))
        summary = office_core.layout_summary(self.layout_opts())
        self.layout_lbl.setText("当前: " + (summary or "保持原文件版面 (不改边距/缩放)"))
        self.fit_width_chk.blockSignals(True)
        self.fit_width_chk.setChecked(lay.get("excel_fit") == "width")
        self.fit_width_chk.blockSignals(False)

    def _on_fit_width_toggled(self, checked):
        lay = self.cfg.setdefault("layout", dict(DEFAULT_LAYOUT_CFG))
        if checked:
            lay["excel_fit"] = "width"
        elif lay.get("excel_fit") == "width":
            lay["excel_fit"] = ""
        self._sync_layout_ui()
        self.save_config()
        if hasattr(self, "preview_win"):
            self.preview_win.refresh()          # 版面变了, 预览立即跟着变

    def _refresh_engine_label(self):
        label, note = office_core.detect_engine()
        if label:
            self.engine_lbl.setText(f"Office 引擎: {label} ({note})")
        else:
            self.engine_lbl.setText("未检测到 Word/Excel, Office 文件将用系统关联程序打印")

    def edit_layout(self):
        d = LayoutDialog(self, self.layout_opts(), self.cfg.get("office_com", True))
        if d.exec() == QDialog.Accepted:
            self.cfg["layout"] = d.result_cfg()
            self.cfg["office_com"] = d.office_chk.isChecked()
            self._sync_layout_ui()
            self.save_config()
            self._log("版面设置: " + (office_core.layout_summary(self.layout_opts())
                                      or "保持原文件版面"))
            if hasattr(self, "preview_win"):
                self.preview_win.refresh()      # 面板改完立刻在预览面板看效果

    # ------------------------------------------------------------ 打印预览面板

    def _build_preview_pane(self):
        """预览面板: 嵌在列表与设置之间, 单击文件即刷新, 可整块收起"""
        try:
            prune_cache()                   # 清掉上次遗留的预览缓存, 防止临时目录膨胀
        except Exception:
            pass
        self.preview_win = PreviewPanel(self)
        self.preview_win.print_handler = self._print_from_preview

    def _preview_width_cfg(self):
        try:
            w = int(self.cfg.get("preview_width", PREVIEW_DEFAULT_W) or PREVIEW_DEFAULT_W)
        except (TypeError, ValueError):
            w = PREVIEW_DEFAULT_W
        return max(PREVIEW_MIN_W, min(PREVIEW_MAX_W, w))

    def _apply_preview_width(self):
        """按配置设置预览列宽度与显隐 (宽度固定, 不参与拉伸)"""
        pw = self.preview_win
        pw.setFixedWidth(self._preview_width_cfg())
        pw.setVisible(bool(self.cfg.get("preview_show", True)))

    def _expand_pane(self):
        """展开预览列 (只调宽度, 不触发转换)"""
        pw = self.preview_win
        pw.setFixedWidth(self._preview_width_cfg())
        pw.setVisible(True)

    def _collapse_pane(self):
        pw = self.preview_win
        if pw.width() > 60:
            self.cfg["preview_width"] = pw.width()
        pw.setVisible(False)

    def toggle_preview_window(self, on):
        """工具栏按钮: 展开/收起预览面板"""
        pw = getattr(self, "preview_win", None)
        if pw is None:
            return
        if on:
            self._expand_pane()
            row = self._preview_row
            if not (0 <= row < len(self.items)):
                rows = self.selected_rows()
                row = rows[0] if rows else -1
            if 0 <= row < len(self.items):
                self._preview_row_at(row)
            else:
                pw.clear()
        else:
            self._collapse_pane()
        self.cfg["preview_show"] = bool(on)
        self.save_config()

    def _on_current_changed(self, row, *_):
        """单击/键盘切换当前行 -> 延迟一点点再预览 (连点只预览最后一个)"""
        if row is None or row < 0:
            return
        self._preview_row = int(row)
        if not self.preview_win.isVisible():    # 面板收起时不做无用转换
            return
        if self._printing():
            return
        self._preview_timer.start()

    def _preview_current(self):
        row = self._preview_row
        if 0 <= row < len(self.items):
            self._preview_row_at(row)

    def _preview_row_at(self, row, force=False):
        if not (0 <= row < len(self.items)):
            return
        if self._printing():
            return
        pw = self.preview_win
        if not pw.isVisible():
            if not force:                   # 收起状态就不偷偷展开
                return
            self.cfg["preview_show"] = True
            self.preview_btn.blockSignals(True)
            self.preview_btn.setChecked(True)
            self.preview_btn.blockSignals(False)
            self._expand_pane()
        item = self.items[row]
        pw.set_file(item["path"], self._opts_for_item(item), self._ext_map(),
                    self.current_printer())

    def _opts_for_item(self, item):
        g = self.cfg["global"]
        paper, _paper_txt = self.current_paper()
        source, _src_txt = self.current_source()
        opts = {
            "copies": item["copies"] if item["copies"] is not None else (g["copies"] or None),
            "color": item["color"] or g["color"] or None,
            "duplex": item["duplex"] or g["duplex"] or None,
            "pages": item["pages"] or g["pages"] or None,
            "layout": self.layout_opts(),
            "office_com": bool(self.cfg.get("office_com", True)),
            "paper": paper,
            "source": source,
        }
        opts.update(self.print_extras())
        return opts

    def _print_from_preview(self, path, opts):
        """预览面板里点「打印此文件」"""
        if self._printing():
            QMessageBox.information(self, "提示", "已有打印任务正在进行, 请稍后再试")
            return False
        printer = self.current_printer()
        if not printer:
            QMessageBox.warning(self, "提示", "请先在主界面选择打印机")
            return False
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            with default_printer_guard(printer, self._log):
                print_file(path, printer, opts, self._ext_map(),
                           find_sumatra(self.cfg.get("sumatra_path", "")), self._log)
        except Exception as e:
            QMessageBox.warning(self, "打印失败", str(e))
            return False
        finally:
            office_core.close_session()     # 关掉主线程里临时启动的 Word/Excel
            QApplication.restoreOverrideCursor()
        self._log(f"已从预览发送: {os.path.basename(path)}")
        return True

    def preview_selected(self, *_):
        """菜单/右键: 展开预览面板并预览当前选中的文件"""
        rows = self.selected_rows()
        if not rows:
            QMessageBox.information(self, "提示", "请先在列表中选中要预览的文件")
            return
        if self._printing():
            QMessageBox.information(self, "提示", "打印进行中, 请等打印结束后再预览")
            return
        self._preview_row_at(rows[0], force=True)

    def _open_preview(self, row):
        self._preview_row_at(row, force=True)

    def refresh_printers(self):
        ps = list_printers()
        default = get_default_printer()
        self._filling_printers = True          # 重建列表期间别触发纸型重新枚举
        try:
            self.printer_combo.clear()
            for p in ps:
                self.printer_combo.addItem(p + ("  (默认)" if p == default else ""))
            saved = self.cfg.get("printer", "")
            for p in ps:
                if p == saved:
                    self.printer_combo.setCurrentText(p + ("  (默认)" if p == default else ""))
                    break
            else:
                if default:
                    self.printer_combo.setCurrentText(default + "  (默认)")
        finally:
            self._filling_printers = False
        self.refresh_paper_choices()

    # ------------------------------------------------------------ 纸型 / 纸张来源

    def _init_paper_ui(self):
        g = self.cfg["global"]
        self.g_per_page.setChecked(bool(g.get("per_page", False)))
        self.g_gap.setValue(max(0.0, float(g.get("page_gap") or 0.0)))
        self.g_per_page.toggled.connect(self._sync_per_page)
        self.printer_combo.currentTextChanged.connect(self._on_printer_changed)
        self._sync_per_page()

    def _sync_per_page(self, *_):
        self.g_gap.setEnabled(self.g_per_page.isChecked())

    def _printer_caps(self):
        """当前打印机支持的 (纸型列表, 纸盒列表), 按打印机缓存 (枚举一次要几十毫秒)"""
        printer = self.current_printer()
        cache = getattr(self, "_caps_cache", None)
        if cache is None:
            cache = self._caps_cache = {}
        if printer not in cache:
            cache[printer] = (list_paper_sizes(printer), list_paper_bins(printer))
        return cache[printer]

    @staticmethod
    def _combo_data(combo):
        try:
            return int(combo.currentData() or 0)
        except (TypeError, ValueError):
            return 0

    @staticmethod
    def _paper_text(name, w, h):
        """纸型显示名: 驱动自带的名字里没有尺寸时补上 宽×高"""
        name = (name or "").strip()
        if not name:
            return "未命名纸型"
        low = name.lower()
        if not w or not h or "mm" in low or "毫米" in name or "英寸" in name or "×" in name:
            return name
        return f"{name} ({w}×{h}mm)"

    def refresh_paper_choices(self):
        """按当前打印机重建 纸型/纸张来源 下拉, 尽量保留已选的值"""
        if not hasattr(self, "g_paper"):
            return
        papers, bins = self._printer_caps()
        g = self.cfg["global"]
        want_paper = self._combo_data(self.g_paper) or int(g.get("paper") or 0)
        want_src = self._combo_data(self.g_source) or int(g.get("source") or 0)

        self.g_paper.blockSignals(True)
        self.g_paper.clear()
        self.g_paper.addItem("跟随打印机", 0)
        for pid, name, w, h in papers:
            self.g_paper.addItem(self._paper_text(name, w, h), pid)
        idx = self.g_paper.findData(want_paper)
        self.g_paper.setCurrentIndex(idx if idx >= 0 else 0)
        self.g_paper.blockSignals(False)
        self.g_paper.setEnabled(self.g_paper.count() > 1)

        self.g_source.blockSignals(True)
        self.g_source.clear()
        self.g_source.addItem("跟随打印机", 0)
        for bid, name in bins:
            self.g_source.addItem(name, bid)
        idx = self.g_source.findData(want_src)
        self.g_source.setCurrentIndex(idx if idx >= 0 else 0)
        self.g_source.blockSignals(False)
        self.g_source.setEnabled(self.g_source.count() > 1)

    def _on_printer_changed(self, *_):
        if getattr(self, "_filling_printers", False):
            return
        self.refresh_paper_choices()

    def current_paper(self):
        """(纸型ID, 纸型显示名); ID 0 = 跟随打印机"""
        return self._combo_data(self.g_paper), self.g_paper.currentText()

    def current_source(self):
        return self._combo_data(self.g_source), self.g_source.currentText()

    def print_extras(self):
        """逐页打印相关设置"""
        return {"per_page": self.g_per_page.isChecked(),
                "page_gap": float(self.g_gap.value()) if self.g_per_page.isChecked() else 0.0}

    def current_printer(self):
        return self.printer_combo.currentText().replace("  (默认)", "").strip()

    def _ext_map(self):
        return {k.lower().lstrip("."): v.lower().lstrip(".")
                for k, v in self.cfg.get("ext_map", {}).items()}

    # ------------------------------------------------------------ 添加文件

    def add_files(self):
        exts = sorted(all_known_exts(self._ext_map()))
        patt = "*." + " *.".join(e.lstrip(".") for e in exts)
        paths, _ = QFileDialog.getOpenFileNames(self, "选择要打印的文件", "",
                                                f"可打印文件 ({patt});;所有文件 (*.*)")
        for p in paths:
            self.enqueue_path(p, recursive=False)

    def add_folder(self):
        d = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if not d:
            return
        n0 = len(self.items)
        self.enqueue_path(d, recursive=self.subfolder_chk.isChecked())
        self._log(f"已从文件夹添加 {len(self.items) - n0} 个文件: {d}")

    def add_archives(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "选择压缩包", "",
            "压缩包 (*.zip *.7z *.rar *.tar *.tar.gz *.tgz *.tar.bz2 *.tar.xz);;所有文件 (*.*)")
        for p in paths:
            self.enqueue_path(p, recursive=False)

    def enqueue_path(self, path, recursive=True):
        ext_map = self._ext_map()
        known = all_known_exts(ext_map)
        if os.path.isdir(path):
            if recursive:
                for root, dirs, files in os.walk(path):
                    # 子文件夹按资源管理器顺序遍历 (原地排序决定 os.walk 顺序)
                    dirs[:] = sort_names(dirs)
                    for f in sort_names(files):
                        self.enqueue_path(os.path.join(root, f), recursive=False)
            else:
                for f in sort_names(os.listdir(path)):
                    full = os.path.join(path, f)
                    if os.path.isfile(full):
                        self.enqueue_path(full, recursive=False)
            return
        if not os.path.isfile(path):
            return
        ext = os.path.splitext(path)[1].lower().lstrip(".")
        if is_archive(path):
            self._extract_and_enqueue(path)
            return
        if ("." + ext) not in known:
            self._log(f"跳过不支持的文件: {os.path.basename(path)}"
                      f" (.{'.' + ext if ext else '无扩展名'}; 可在扩展名映射中添加别名)")
            return
        cat = resolve_category(ext, ext_map) or "?"
        self._add_item(path, cat)

    def _add_item(self, path, cat, temp_root=None):
        self._seq += 1
        self._added_seq += 1
        item = {"path": path, "type": cat, "copies": None, "color": None,
                "duplex": None, "pages": None, "status": STATUS_QUEUED,
                "temp_root": temp_root, "added": self._added_seq}
        self.items.append(item)
        r = self.table.rowCount()
        self.table.insertRow(r)
        vals = [str(self._seq), os.path.basename(path), cat, "全局", "全局", "全局", "全局", STATUS_QUEUED]
        for c, v in enumerate(vals):
            it = QTableWidgetItem(v)
            if c == C_COPIES or c == C_PAGES:
                it.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(r, c, it)

    def _extract_and_enqueue(self, arc_path):
        try:
            dest, files = extract_archive(arc_path, tempfile.gettempdir(), self._log)
        except Exception as e:
            self._log(f"❌ 解压失败 {os.path.basename(arc_path)}: {e}")
            return
        self.temp_roots.add(dest)
        ext_map = self._ext_map()
        known = all_known_exts(ext_map)
        added = 0
        for f in sort_names(files):
            ext = os.path.splitext(f)[1].lower().lstrip(".")
            if ("." + ext) in known:
                self._add_item(f, resolve_category(ext, ext_map) or "?", temp_root=dest)
                added += 1
        self._log(f"压缩包 {os.path.basename(arc_path)}: 解压出 {len(files)} 个文件, "
                  f"其中 {added} 个加入打印队列")

    # ------------------------------------------------------------ 顺序调整

    def _printing(self):
        return bool(self.print_thread and self.print_thread.is_alive())

    def _apply_reorder(self, src_rows, dst_row):
        """把 src_rows 这些行整体移动到 dst_row 位置 (拖动/上移下移共用)"""
        if self._printing():
            self._log("⚠ 打印进行中, 暂不能调整顺序 (停止后可调)")
            return
        n = len(self.items)
        src = [r for r in src_rows if 0 <= r < n]
        if not src or n < 2:
            return
        moving = [self.items[r] for r in src]
        src_set = set(src)
        rest = [it for i, it in enumerate(self.items) if i not in src_set]
        insert = dst_row - sum(1 for r in src if r < dst_row)
        insert = max(0, min(insert, len(rest)))
        new_items = rest[:insert] + moving + rest[insert:]
        if [id(x) for x in new_items] == [id(x) for x in self.items]:
            return                                    # 位置没变
        self.items = new_items
        self._rebuild_table()
        self.table.clearSelection()
        for r in range(insert, insert + len(moving)):
            self.table.setRangeSelected(
                QTableWidgetSelectionRange(r, 0, r, len(COLS) - 1), True)
        self.table.scrollToItem(self.table.item(insert, C_NAME))
        self._log(f"顺序已调整: {len(moving)} 个文件移到第 {insert + 1} 位")

    def move_selected(self, delta):
        rows = sorted({i.row() for i in self.table.selectedIndexes()})
        if not rows:
            self._log("请先选中要移动的行")
            self.drag_hint.setText("⚠ 请先在列表中选中要移动的行, 再点 上移/下移")
            return
        if delta < 0:
            if rows[0] == 0:
                return
            self._apply_reorder(rows, rows[0] - 1)
        else:
            if rows[-1] >= len(self.items) - 1:
                return
            self._apply_reorder(rows, rows[-1] + 2)

    def sort_by_name(self):
        self.apply_sort("name")

    def sort_by_added(self):
        self.apply_sort("added")

    def _on_sort_changed(self, text):
        mode = SORT_MODES.get(text, "added")
        if not self.apply_sort(mode):
            # 未能生效 (打印中/列表不足), 回到上一个选择
            self.sort_combo.blockSignals(True)
            self.sort_combo.setCurrentText(self._last_sort_label)
            self.sort_combo.blockSignals(False)

    def apply_sort(self, mode):
        """按指定方式重排整个列表, 返回是否生效"""
        if self._printing():
            self._log("⚠ 打印进行中, 暂不能调整顺序")
            return False
        if mode == "added":
            self._last_sort_label = "添加顺序"
        elif mode == "name":
            self._last_sort_label = "文件名 (A→Z)"
        elif mode == "mtime_asc":
            self._last_sort_label = "修改时间 (早→晚)"
        else:
            self._last_sort_label = "修改时间 (晚→早)"
        if len(self.items) < 2:
            return True
        if mode == "added":
            self.items.sort(key=lambda it: it.get("added", 0))
            tip = "按添加顺序"
        elif mode == "name":
            self.items.sort(key=cmp_to_key(
                lambda a, b: natural_cmp(os.path.basename(a["path"]),
                                         os.path.basename(b["path"]))))
            tip = "按文件名 (资源管理器顺序)"
        elif mode == "mtime_asc":
            self.items.sort(key=lambda it: (file_mtime(it["path"]),
                                            os.path.basename(it["path"]).lower()))
            tip = "按修改时间 早→晚"
        else:
            self.items.sort(key=lambda it: (file_mtime(it["path"]),
                                            os.path.basename(it["path"]).lower()),
                            reverse=True)
            tip = "按修改时间 晚→早"
        self._rebuild_table()
        self._log(f"已{tip}重排列表 ({len(self.items)} 项)")
        return True

    def _rebuild_table(self):
        """按 self.items 重绘整个表格 (行号重排为 1..n)"""
        self.table.setRowCount(0)
        for i, item in enumerate(self.items):
            self.table.insertRow(i)
            vals = [
                str(i + 1),
                os.path.basename(item["path"]),
                item["type"],
                "全局" if item["copies"] is None else str(item["copies"]),
                "全局" if not item["color"] else COLOR_LABELS.get(item["color"], item["color"]),
                "全局" if not item["duplex"] else DUPLEX_LABELS.get(item["duplex"], item["duplex"]),
                "全局" if not item["pages"] else item["pages"],
                item["status"],
            ]
            for c, v in enumerate(vals):
                it = QTableWidgetItem(str(v))
                if c in (C_COPIES, C_PAGES):
                    it.setTextAlignment(Qt.AlignCenter)
                if c == C_STATUS:
                    col = STATUS_COLORS.get(item["status"])
                    if col:
                        it.setForeground(QColor(col))
                self.table.setItem(i, c, it)
        self._seq = len(self.items)

    def remove_selected(self):
        rows = sorted({i.row() for i in self.table.selectedIndexes()}, reverse=True)
        if not rows:
            return
        for r in rows:
            if 0 <= r < len(self.items):
                self.items.pop(r)
        self._rebuild_table()
        self._log(f"已移除 {len(rows)} 项")

    def clear_list(self):
        self.table.setRowCount(0)
        self.items.clear()
        self._seq = 0
        self._added_seq = 0
        self._preview_row = -1
        if hasattr(self, "preview_win"):
            self.preview_win.clear()

    # ------------------------------------------------------------ 个体/批量设置

    def selected_rows(self):
        return sorted({i.row() for i in self.table.selectedIndexes()})

    def _on_double_click(self, *_):
        self.batch_settings(self.selected_rows())

    def _table_menu(self, pos):
        rows = self.selected_rows()
        m = QMenu(self)
        n = len(rows)
        act_preview = m.addAction("👁 在预览面板中查看" if n <= 1 else "👁 预览第一个选中的")
        act_preview.triggered.connect(self.preview_selected)
        act_preview.setEnabled(bool(rows) and n >= 1)
        act_set = m.addAction(f"⚙ 批量设置选中的 {n} 个文件" if n > 1 else "⚙ 设置此文件")
        act_set.triggered.connect(lambda: self.batch_settings(rows))
        m.addSeparator()
        act_up = m.addAction("↑ 上移")
        act_up.triggered.connect(lambda: self.move_selected(-1))
        act_up.setEnabled(bool(rows))
        act_dn = m.addAction("↓ 下移")
        act_dn.triggered.connect(lambda: self.move_selected(1))
        act_dn.setEnabled(bool(rows))
        m.addSeparator()
        sm = m.addMenu("排序")
        for label, key in SORT_MODES.items():
            a = sm.addAction(label)
            a.triggered.connect(lambda _=False, k=key: self.apply_sort(k))
        m.addSeparator()
        act_del = m.addAction(f"移除选中的 {n} 项" if n > 1 else "移除此文件")
        act_del.triggered.connect(self.remove_selected)
        act_del.setEnabled(bool(rows))
        m.exec(self.table.viewport().mapToGlobal(pos))

    def batch_settings(self, rows=None, *args):
        """对多个选中的文件统一设置打印参数 (留空/选"不修改"则保持原值)"""
        # 按钮的 clicked 信号会附带 checked(bool) 参数, 这里统一过滤掉
        if rows is not None and not isinstance(rows, (list, tuple, set)):
            rows = None
        rows = self.selected_rows() if rows is None else sorted(rows)
        rows = [r for r in rows if 0 <= r < len(self.items)]
        if not rows:
            QMessageBox.information(
                self, "提示",
                "请先在列表中选中要设置的文件。\n可按住 Ctrl 逐个点选, 或 Shift 选择连续多行。")
            return
        single = len(rows) == 1
        d = QDialog(self)
        d.setWindowTitle("单文件打印设置" if single else f"批量设置 ({len(rows)} 个文件)")
        d.setModal(True)
        d.setMinimumWidth(430)
        g = QGridLayout(d)
        g.setVerticalSpacing(8)

        head = QLabel(os.path.basename(self.items[rows[0]]["path"]) if single
                      else f"将对选中的 {len(rows)} 个文件统一应用以下设置")
        head.setStyleSheet("font-weight:bold;")
        head.setWordWrap(True)
        g.addWidget(head, 0, 0, 1, 2)

        def common(key):
            """所有选中项取值一致时返回该值, 否则返回 None"""
            vals = {self.items[r][key] for r in rows}
            return vals.pop() if len(vals) == 1 else None

        def describe(key, labels, default):
            """当前选中的所有文件在该项上的取值摘要"""
            vals = {self.items[r][key] for r in rows}
            if len(vals) > 1:
                return "各不相同"
            v = vals.pop()
            if key == "copies":
                return default if v is None else f"{v} 份"
            if key == "pages":
                return default if not v else v
            return labels.get(v or "", default)

        if not single:
            lab = QLabel("这些文件当前: " + ", ".join([
                "份数=" + describe("copies", COLOR_LABELS, "跟随全局"),
                "色彩=" + describe("color", COLOR_LABELS, "跟随打印机"),
                "双面=" + describe("duplex", DUPLEX_LABELS, "跟随打印机"),
                "页码=" + describe("pages", COLOR_LABELS, "全局"),
            ]))
            lab.setStyleSheet("color:#666;font-size:11px;")
            lab.setWordWrap(True)
            g.addWidget(lab, 1, 0, 1, 2)

        base = 2 if not single else 1

        # 份数 (多选时一律留空=不修改)
        g.addWidget(QLabel("份数:"), base, 0)
        c_val = common("copies") if single else None
        copies_e = QLineEdit("" if c_val is None else str(c_val))
        copies_e.setPlaceholderText("不修改" if not single else "留空=跟随全局")
        g.addWidget(copies_e, base, 1)

        # 色彩
        g.addWidget(QLabel("色彩模式:"), base + 1, 0)
        color_labels = ["不修改"] + list(COLOR_LABELS.values())
        color_c = QComboBox()
        color_c.addItems(color_labels)
        if single:
            color_c.removeItem(0)
            color_c.setCurrentText(COLOR_LABELS.get(self.items[rows[0]]["color"] or "",
                                                    "跟随打印机"))
        g.addWidget(color_c, base + 1, 1)

        # 双面
        g.addWidget(QLabel("双面/单面:"), base + 2, 0)
        duplex_labels = ["不修改"] + list(DUPLEX_LABELS.values())
        duplex_c = QComboBox()
        duplex_c.addItems(duplex_labels)
        if single:
            duplex_c.removeItem(0)
            duplex_c.setCurrentText(DUPLEX_LABELS.get(self.items[rows[0]]["duplex"] or "",
                                                      "跟随打印机"))
        g.addWidget(duplex_c, base + 2, 1)

        # 页码范围
        g.addWidget(QLabel("页码范围:"), base + 3, 0)
        p_val = common("pages") if single else None
        pages_e = QLineEdit("" if p_val is None else p_val)
        pages_e.setPlaceholderText("不修改, 如 1-3,5" if not single else "留空=跟随全局")
        g.addWidget(pages_e, base + 3, 1)

        reset_chk = None
        if not single:
            reset_chk = QCheckBox("先清除这些文件已有的单独设置 (全部改为跟随全局)")
            g.addWidget(reset_chk, base + 4, 0, 1, 2)
        tip = QLabel("留空或选「不修改」= 保持文件原有设置; 选「跟随打印机」= 用打印机默认参数")
        tip.setStyleSheet("color:#888;font-size:11px;")
        tip.setWordWrap(True)
        g.addWidget(tip, base + 5, 0, 1, 2)

        box = QHBoxLayout()
        ok_b = QPushButton("确定")
        ok_b.setDefault(True)
        cancel_b = QPushButton("取消")
        box.addStretch()
        box.addWidget(ok_b)
        box.addWidget(cancel_b)
        g.addLayout(box, base + 6, 0, 1, 2)
        cancel_b.clicked.connect(d.reject)

        def save():
            copies_txt = copies_e.text().strip()
            copies = None
            if copies_txt:
                try:
                    copies = int(float(copies_txt))
                    if copies < 1:
                        raise ValueError
                except ValueError:
                    QMessageBox.warning(d, "份数无效", "份数请填 1 以上的整数, 或留空不修改")
                    return
            pages_txt = pages_e.text().strip()
            if pages_txt:
                try:
                    if not parse_page_range(pages_txt):
                        raise ValueError
                except Exception:
                    QMessageBox.warning(d, "页码范围无效",
                                        "页码范围格式如 1-3,5,8。\n留空则不修改。")
                    return
            color_txt = color_c.currentText()
            duplex_txt = duplex_c.currentText()
            changed = 0
            for r in rows:
                it = self.items[r]
                if reset_chk is not None and reset_chk.isChecked():
                    it.update({"copies": None, "color": "", "duplex": "",
                               "pages": None})
                if copies is not None:
                    it["copies"] = copies
                if color_txt != "不修改":
                    it["color"] = COLOR_LABEL2VAL.get(color_txt, "")
                if duplex_txt != "不修改":
                    it["duplex"] = DUPLEX_LABEL2VAL.get(duplex_txt, "")
                if pages_txt:
                    it["pages"] = pages_txt
                changed += 1
            self._rebuild_table()
            for r in rows:
                self.table.setRangeSelected(
                    QTableWidgetSelectionRange(r, 0, r, len(COLS) - 1), True)
            self._log(f"已设置 {changed} 个文件: "
                      f"份数={'不修改' if copies is None else copies}, "
                      f"色彩={color_txt}, 双面={duplex_txt}, "
                      f"页码={pages_txt or '不修改'}")
            d.accept()

        ok_b.clicked.connect(save)
        d.exec()

    # ------------------------------------------------------------ 扩展名映射

    def edit_extmap(self):
        d = QDialog(self)
        d.setWindowTitle("文件扩展名映射")
        d.setModal(True)
        d.resize(430, 400)
        v = QVBoxLayout(d)
        tip = QLabel("把非标准扩展名映射到受支持的格式, 例如:\n"
                     "  dat → txt  (.dat 当文本文件打印)\n"
                     "  emb → pdf  (.emb 当 PDF 打印)")
        v.addWidget(tip)

        tv = QTableWidget(0, 2)
        tv.setHorizontalHeaderLabels(["别名扩展名", "映射到"])
        tv.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        v.addWidget(tv)

        def reload_rows():
            tv.setRowCount(0)
            for k, val in sorted(self.cfg.get("ext_map", {}).items()):
                r = tv.rowCount()
                tv.insertRow(r)
                tv.setItem(r, 0, QTableWidgetItem(k))
                tv.setItem(r, 1, QTableWidgetItem(val))

        reload_rows()

        def add_row():
            sub = QDialog(d)
            sub.setWindowTitle("添加映射")
            sub.setModal(True)
            gg = QGridLayout(sub)
            gg.addWidget(QLabel("别名 (如 dat):"), 0, 0)
            e1 = QLineEdit()
            e1.setPlaceholderText("不含点, 如 dat")
            gg.addWidget(e1, 0, 1)
            gg.addWidget(QLabel("映射到:"), 1, 0)
            e2 = QComboBox()
            e2.addItems(["pdf", "jpg", "png", "txt", "docx", "xlsx", "pptx", "prn"])
            gg.addWidget(e2, 1, 1)

            def ok():
                a = e1.text().strip().lstrip(".").lower()
                t = e2.currentText().strip().lower()
                if a and t:
                    self.cfg.setdefault("ext_map", {})[a] = t
                    reload_rows()
                    sub.accept()

            ob = QPushButton("确定")
            ob.clicked.connect(ok)
            gg.addWidget(ob, 2, 0, 1, 2)
            sub.exec()

        def del_row():
            rows = sorted({i.row() for i in tv.selectedIndexes()}, reverse=True)
            for r in rows:
                self.cfg.get("ext_map", {}).pop(tv.item(r, 0).text(), None)
            reload_rows()

        hb = QHBoxLayout()
        ab = QPushButton("添加映射")
        ab.clicked.connect(add_row)
        db = QPushButton("删除选中")
        db.clicked.connect(del_row)
        cb = QPushButton("关闭")
        cb.clicked.connect(d.accept)
        hb.addWidget(ab)
        hb.addWidget(db)
        hb.addStretch()
        hb.addWidget(cb)
        v.addLayout(hb)
        d.exec()
        self._save_and_refresh()

    def _save_and_refresh(self):
        self.save_config()

    # ------------------------------------------------------------ 打印

    def start_print(self):
        if not self.items:
            QMessageBox.information(self, "提示", "请先添加要打印的文件")
            return
        printer = self.current_printer()
        if not printer:
            QMessageBox.warning(self, "提示", "请先选择打印机")
            return
        if self.print_thread and self.print_thread.is_alive():
            QMessageBox.warning(self, "提示", "已有打印任务正在进行")
            return
        self.save_config()
        self._refresh_sumatra_label()

        g = self.cfg["global"]
        sumatra = find_sumatra(self.cfg.get("sumatra_path", ""))
        interval = self._to_float(self.interval_edit.text(), 2.0)
        ext_map = self._ext_map()
        items_snapshot = list(self.items)
        layout = self.layout_opts()
        office_com = bool(self.cfg.get("office_com", True))
        paper, paper_txt = self.current_paper()
        source, source_txt = self.current_source()
        extras = self.print_extras()

        self.cancel_event.clear()
        self.print_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self._log(f"=== 开始打印, 共 {len(items_snapshot)} 项, 打印机: {printer} ===")
        lay_txt = office_core.layout_summary(layout)
        if lay_txt:
            self._log(f"版面设置: {lay_txt}")
        if paper:
            self._log(f"纸型: {paper_txt} (内容按纸型缩放)")
        if source:
            self._log(f"纸张来源: {source_txt}")
        if extras["per_page"]:
            self._log(f"逐页打印: 每页一个作业, 页间等待 {extras['page_gap']:g} 秒")
        if office_com:
            self._log("Office 文件使用静默打印 (COM 转 PDF 后发送, 不会弹出程序窗口)")

        def work():
            sig = self.sig
            # 整批打印期间把系统默认打印机切到所选打印机, 结束后恢复。
            # 没有这一步, Office/未知格式的 Shell 打印永远只会打到默认打印机。
            guard_log = lambda m: sig.log.emit(m)  # noqa: E731
            try:
                with default_printer_guard(printer, guard_log):
                    for idx, item in enumerate(items_snapshot):
                        if self.cancel_event.is_set():
                            sig.status.emit(idx, "已取消")
                            continue
                        sig.status.emit(idx, "打印中")
                        opts = {
                            "copies": item["copies"] if item["copies"] is not None else (g["copies"] or None),
                            "color": item["color"] or g["color"] or None,
                            "duplex": item["duplex"] or g["duplex"] or None,
                            "pages": item["pages"] or g["pages"] or None,
                            "layout": layout,
                            "office_com": office_com,
                            "paper": paper,
                            "source": source,
                        }
                        opts.update(extras)
                        try:
                            print_file(item["path"], printer, opts, ext_map, sumatra,
                                       lambda m: sig.log.emit("  " + m))
                            sig.status.emit(idx, "已发送")
                        except Exception as e:
                            sig.log.emit(f"❌ {os.path.basename(item['path'])}: {e}")
                            sig.status.emit(idx, "失败")
                        if interval > 0:
                            self.cancel_event.wait(interval)
            finally:
                # 关闭本线程里复用的 Word/Excel 进程 (COM 进程不关会残留在后台)
                office_core.close_session()
                sig.done.emit()

        self.print_thread = threading.Thread(target=work, daemon=True)
        self.print_thread.start()

    def stop_print(self):
        self.cancel_event.set()
        self._log("⏹ 已请求停止, 等待当前文件处理完成...")

    def _set_status(self, idx, status):
        if 0 <= idx < len(self.items) and idx < self.table.rowCount():
            self.items[idx]["status"] = status
            it = self.table.item(idx, C_STATUS)
            it.setText(status)
            color = STATUS_COLORS.get(status)
            if color:
                it.setForeground(QColor(color))
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum())

    def _on_done(self):
        self.print_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        ok = sum(1 for it in self.items if it["status"] == "已发送")
        fail = sum(1 for it in self.items if it["status"] == "失败")
        self._log(f"=== 打印结束: 成功 {ok}, 失败 {fail}, 共 {len(self.items)} ===")
        self._cleanup_temp()
        self.save_config()

    def _cleanup_temp(self):
        if self.cancel_event.is_set():
            return
        for t in list(self.temp_roots):
            cleanup_dir(t)
            self.temp_roots.discard(t)

    def closeEvent(self, e):
        if self.print_thread and self.print_thread.is_alive():
            if QMessageBox.question(self, "退出", "打印任务进行中, 确定退出吗?") != QMessageBox.Yes:
                e.ignore()
                return
            self.cancel_event.set()
        pw = getattr(self, "preview_win", None)
        self.save_config()
        if pw is not None:
            pw.shutdown()                    # 停掉后台预览线程, 关掉临时 Word/Excel
        for t in self.temp_roots:
            cleanup_dir(t)
        e.accept()


def _office_selftest():
    """打包后自检: 带环境变量 BP_SELFTEST=<Office 文件> 启动时, 导出 PDF 并写结果.

    结果写入 %TEMP%\\bp_selftest.txt (窗口程序没有控制台, 只能落文件)。
    仅用于验证打包后的 COM 链路是否可用, 正常使用不会触发。
    """
    src = os.environ.get("BP_SELFTEST")
    if not src:
        return False
    out = os.path.join(tempfile.gettempdir(), "bp_selftest.txt")
    text = ""
    try:
        pdf, work = office_core.export_pdf(
            src, {"margins_mm": [15, 15, 15, 15], "excel_fit": "width"})
        import pypdfium2 as pdfium
        d = pdfium.PdfDocument(pdf)
        n = len(d)
        d.close()
        shutil.rmtree(work, ignore_errors=True)
        text = f"OK pages={n} pdf={pdf}\n"
    except Exception as e:
        text = f"FAIL {type(e).__name__}: {e}\n"
    finally:
        office_core.close_session()
    try:
        with open(out, "w", encoding="utf-8") as f:
            f.write(text)
    except Exception:
        pass
    return True


def main():
    if _office_selftest():
        return
    app = QApplication(sys.argv)
    win = BatchPrinterApp()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
