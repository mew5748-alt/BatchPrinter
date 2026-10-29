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
import sys
import json
import queue
import threading
import tempfile
import shutil

from PySide6.QtCore import Qt, QObject, Signal
from PySide6.QtGui import QColor, QAction
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QComboBox, QLabel, QTableWidget, QTableWidgetItem, QSpinBox,
    QCheckBox, QLineEdit, QTextEdit, QFileDialog, QMessageBox, QDialog,
    QHeaderView, QAbstractItemView, QGroupBox, QSplitter, QToolBar, QStyle,
)

from printer_core import (
    list_printers, get_default_printer, find_sumatra, print_file,
    resolve_category, all_known_exts,
)
from archive_util import extract_archive, cleanup_dir, is_archive

APP_DIR = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
    else os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "config.json")

DEFAULT_CONFIG = {
    "printer": "",
    "global": {"copies": 1, "color": "", "duplex": "", "pages": ""},
    "subfolders": True,
    "interval": 2.0,
    "sumatra_path": "",
    "ext_map": {},
}

COLOR_LABELS = {"": "跟随打印机", "color": "彩色", "mono": "黑白"}
DUPLEX_LABELS = {"": "跟随打印机", "simplex": "单面",
                 "duplex_long": "双面(长边)", "duplex_short": "双面(短边)"}
COLOR_LABEL2VAL = {v: k for k, v in COLOR_LABELS.items()}
DUPLEX_LABEL2VAL = {v: k for k, v in DUPLEX_LABELS.items()}

STATUS_QUEUED = "等待中"
C_SEQ, C_NAME, C_TYPE, C_COPIES, C_COLOR, C_DUPLEX, C_PAGES, C_STATUS = range(8)
COLS = ["#", "文件名", "类型", "份数", "色彩", "双面", "页码范围", "状态"]
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
                else:
                    cfg[k] = v
        except Exception:
            pass
    return cfg


class WorkerSignals(QObject):
    log = Signal(str)
    status = Signal(int, str)
    done = Signal()


class BatchPrinterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("批量打印助手 - Batch Printer")
        self.resize(1180, 780)

        self.cfg = load_config()
        self.items = []
        self.temp_roots = set()
        self._seq = 0
        self.print_thread = None
        self.cancel_event = threading.Event()
        self.sig = WorkerSignals()
        self.sig.log.connect(self._log)
        self.sig.status.connect(self._set_status)
        self.sig.done.connect(self._on_done)

        self._build_ui()
        self.refresh_printers()
        self._refresh_sumatra_label()
        self.setAcceptDrops(True)

    # ------------------------------------------------------------ UI 构建

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)

        # 工具栏
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
                bar.addSpacing(18)
        bar.addStretch()
        root.addLayout(bar)

        # 打印机行
        prow = QHBoxLayout()
        prow.addWidget(QLabel("打印机:"))
        self.printer_combo = QComboBox()
        self.printer_combo.setMinimumWidth(420)
        prow.addWidget(self.printer_combo)
        refresh = QPushButton("刷新")
        refresh.setFixedWidth(60)
        refresh.clicked.connect(self.refresh_printers)
        prow.addWidget(refresh)
        self.sumatra_lbl = QLabel("")
        prow.addWidget(self.sumatra_lbl)
        prow.addStretch()
        root.addLayout(prow)

        # 主体: 表格 + 设置
        splitter = QSplitter(Qt.Horizontal)
        root.addWidget(splitter, 1)

        self.table = QTableWidget(0, len(COLS))
        self.table.setHorizontalHeaderLabels(COLS)
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(C_NAME, QHeaderView.Stretch)
        for c in (C_STATUS, C_COPIES, C_TYPE, C_SEQ):
            hh.setSectionResizeMode(c, QHeaderView.ResizeToContents)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.doubleClicked.connect(self.edit_item_settings)
        splitter.addWidget(self.table)

        right = QGroupBox(" 全局设置 (作用于整个列表) ")
        right.setMinimumWidth(250)
        g = QGridLayout(right)
        g.setVerticalSpacing(8)

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

        self.subfolder_chk = QCheckBox("添加文件夹时包含子文件夹")
        self.subfolder_chk.setChecked(bool(self.cfg.get("subfolders", True)))
        g.addWidget(self.subfolder_chk, 4, 0, 1, 2)

        g.addWidget(QLabel("打印间隔(秒):"), 5, 0)
        self.interval_edit = QLineEdit(str(self.cfg.get("interval", 2.0)))
        g.addWidget(self.interval_edit, 5, 1)

        g.addWidget(QLabel("SumatraPDF:"), 6, 0)
        sw = QHBoxLayout()
        self.sumatra_edit = QLineEdit(self.cfg.get("sumatra_path", ""))
        sw.addWidget(self.sumatra_edit)
        browse = QPushButton("…")
        browse.setFixedWidth(28)
        browse.clicked.connect(self._browse_sumatra)
        sw.addWidget(browse)
        g.addLayout(sw, 6, 1)
        hint = QLabel("PDF 静默打印推荐安装,\n留空自动搜索")
        hint.setStyleSheet("color:#888;font-size:11px;")
        g.addWidget(hint, 7, 0, 1, 2)

        self.print_btn = QPushButton("🖨 立即打印全部")
        self.print_btn.setMinimumHeight(40)
        self.print_btn.setStyleSheet("font-weight:bold;font-size:14px;")
        self.print_btn.clicked.connect(self.start_print)
        g.addWidget(self.print_btn, 8, 0, 1, 2)
        self.stop_btn = QPushButton("停止")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_print)
        g.addWidget(self.stop_btn, 9, 0, 1, 2)
        g.setRowStretch(10, 1)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 1)

        # 日志
        logg = QGroupBox(" 日志 ")
        lv = QVBoxLayout(logg)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
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
        }
        self.cfg["subfolders"] = self.subfolder_chk.isChecked()
        try:
            self.cfg["interval"] = max(0.0, float(self.interval_edit.text()))
        except ValueError:
            self.cfg["interval"] = 2.0
        self.cfg["sumatra_path"] = self.sumatra_edit.text().strip()
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
            self.sumatra_lbl.setText("✔ 检测到 SumatraPDF (PDF 静默打印最优)")
            self.sumatra_lbl.setStyleSheet("color:#1c7c2e;")
        else:
            self.sumatra_lbl.setText("PDF 使用内置渲染引擎 (安装 SumatraPDF 效果更佳)")
            self.sumatra_lbl.setStyleSheet("color:#0a7d32;")

    def _browse_sumatra(self):
        p, _ = QFileDialog.getOpenFileName(self, "选择 SumatraPDF.exe", "",
                                           "SumatraPDF.exe (SumatraPDF.exe);;所有文件 (*.*)")
        if p:
            self.sumatra_edit.setText(p)
            self._refresh_sumatra_label()

    def refresh_printers(self):
        ps = list_printers()
        default = get_default_printer()
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
                for root, _dirs, files in os.walk(path):
                    for f in sorted(files):
                        self.enqueue_path(os.path.join(root, f), recursive=False)
            else:
                for f in sorted(os.listdir(path)):
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
        item = {"path": path, "type": cat, "copies": None, "color": None,
                "duplex": None, "pages": None, "status": STATUS_QUEUED, "temp_root": temp_root}
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
        for f in sorted(files):
            ext = os.path.splitext(f)[1].lower().lstrip(".")
            if ("." + ext) in known:
                self._add_item(f, resolve_category(ext, ext_map) or "?", temp_root=dest)
                added += 1
        self._log(f"压缩包 {os.path.basename(arc_path)}: 解压出 {len(files)} 个文件, "
                  f"其中 {added} 个加入打印队列")

    def remove_selected(self):
        rows = sorted({i.row() for i in self.table.selectedIndexes()}, reverse=True)
        for r in rows:
            self.table.removeRow(r)
            self.items.pop(r)

    def clear_list(self):
        self.table.setRowCount(0)
        self.items.clear()

    # ------------------------------------------------------------ 个体设置

    def edit_item_settings(self):
        idx = self.table.currentRow()
        if idx < 0 or idx >= len(self.items):
            return
        item = self.items[idx]
        d = QDialog(self)
        d.setWindowTitle("单文件打印设置")
        d.setModal(True)
        g = QGridLayout(d)
        g.setVerticalSpacing(8)

        name = QLabel(os.path.basename(item["path"]))
        name.setStyleSheet("font-weight:bold;")
        g.addWidget(name, 0, 0, 1, 2)

        g.addWidget(QLabel("份数 (留空=跟随全局):"), 1, 0)
        copies_e = QLineEdit("" if item["copies"] is None else str(item["copies"]))
        g.addWidget(copies_e, 1, 1)

        g.addWidget(QLabel("色彩模式:"), 2, 0)
        color_c = QComboBox()
        color_c.addItems(["跟随全局"] + list(COLOR_LABELS.values()))
        color_c.setCurrentText(COLOR_LABELS.get(item["color"] or "", "跟随全局"))
        g.addWidget(color_c, 2, 1)

        g.addWidget(QLabel("双面/单面:"), 3, 0)
        duplex_c = QComboBox()
        duplex_c.addItems(["跟随全局"] + list(DUPLEX_LABELS.values()))
        duplex_c.setCurrentText(DUPLEX_LABELS.get(item["duplex"] or "", "跟随全局"))
        g.addWidget(duplex_c, 3, 1)

        g.addWidget(QLabel("页码范围 (留空=跟随全局):"), 4, 0)
        pages_e = QLineEdit(item["pages"] or "")
        g.addWidget(pages_e, 4, 1)

        def save():
            c = copies_e.text().strip()
            try:
                item["copies"] = int(float(c)) if c else None
            except ValueError:
                item["copies"] = None
            item["color"] = "" if color_c.currentText() == "跟随全局" \
                else COLOR_LABEL2VAL.get(color_c.currentText(), None)
            item["duplex"] = "" if duplex_c.currentText() == "跟随全局" \
                else DUPLEX_LABEL2VAL.get(duplex_c.currentText(), None)
            item["pages"] = pages_e.text().strip() or None
            self.table.item(idx, C_COPIES).setText("全局" if item["copies"] is None else str(item["copies"]))
            self.table.item(idx, C_COLOR).setText("全局" if not item["color"] else COLOR_LABELS[item["color"]])
            self.table.item(idx, C_DUPLEX).setText("全局" if not item["duplex"] else DUPLEX_LABELS[item["duplex"]])
            self.table.item(idx, C_PAGES).setText("全局" if not item["pages"] else item["pages"])
            d.accept()

        save_b = QPushButton("保存")
        save_b.clicked.connect(save)
        g.addWidget(save_b, 5, 0, 1, 2)
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

        self.cancel_event.clear()
        self.print_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self._log(f"=== 开始打印, 共 {len(items_snapshot)} 项, 打印机: {printer} ===")

        def work():
            sig = self.sig
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
                }
                try:
                    print_file(item["path"], printer, opts, ext_map, sumatra,
                               lambda m: sig.log.emit("  " + m))
                    sig.status.emit(idx, "已发送")
                except Exception as e:
                    sig.log.emit(f"❌ {os.path.basename(item['path'])}: {e}")
                    sig.status.emit(idx, "失败")
                if interval > 0:
                    self.cancel_event.wait(interval)
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
        self.save_config()
        for t in self.temp_roots:
            cleanup_dir(t)
        e.accept()


def main():
    app = QApplication(sys.argv)
    win = BatchPrinterApp()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
