# -*- coding: utf-8 -*-
"""生效参数显示自测 (不需要打印机, 不出纸)

覆盖:
  1. 列表里显示的是"实际会用的值", 不是"全局"两个字
  2. 改右侧全局参数 -> 各行显示立刻跟着变 (无需保存/重开)
  3. 灰色=跟随全局, 黑色加粗=本文件单独设置; 单独设置的项不再被全局覆盖
  4. 顶部摘要条包含份数/色彩/双面/页码/纸型/进纸
  5. 单文件打印参数 (_opts_for_item) 取界面实时值, 不依赖上次保存的 config

运行: python test_effective.py
"""

import os
import sys
import time
import tempfile

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")   # 无桌面环境也能跑

from PySide6.QtWidgets import QApplication                     # noqa: E402

import main as M                                               # noqa: E402

PASS = FAIL = 0


def check(name, cond, extra=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print("  PASS", name)
    else:
        FAIL += 1
        print("  FAIL", name, extra)


def pump(app, ms=250):
    """跑一小会儿事件循环, 让 80ms 节流定时器有机会触发"""
    t0 = time.time()
    while (time.time() - t0) * 1000 < ms:
        app.processEvents()
        time.sleep(0.01)


def cell(w, r, c):
    it = w.table.item(r, c)
    return None if it is None else it.text()


def main():
    app = QApplication.instance() or QApplication(sys.argv)
    tmp = tempfile.mkdtemp(prefix="bp_test_eff_")

    # 配置指向临时目录, 避免动到用户的 config.json
    M.CONFIG_PATH = os.path.join(tmp, "config.json")

    files = []
    for name in ("a.docx", "b.pdf", "c.txt"):
        p = os.path.join(tmp, name)
        with open(p, "wb") as f:
            f.write(b"x")
        files.append(p)

    w = M.BatchPrinterApp()
    w.resize(1360, 800)

    print("\n[1] 初始状态")
    for p in files:
        w._add_item(p, M.resolve_category(p.rsplit(".", 1)[1], {}))
    pump(app, 120)
    check("3 行已加入", w.table.rowCount() == 3, str(w.table.rowCount()))
    check("份数列显示全局值 1", cell(w, 0, M.C_COPIES) == "1", cell(w, 0, M.C_COPIES))
    check("不再显示「全局」占位", "全局" not in (cell(w, 0, M.C_COPIES) or ""),
          str(cell(w, 0, M.C_COPIES)))
    check("色彩列显示 驱动", cell(w, 0, M.C_COLOR) == "驱动", cell(w, 0, M.C_COLOR))
    check("页码列显示 全部", cell(w, 0, M.C_PAGES) == "全部", cell(w, 0, M.C_PAGES))
    f0 = w.table.item(0, M.C_COPIES).font()
    check("跟随全局的数值不加粗", not f0.bold())
    check("跟随全局的数值是灰色",
          w.table.item(0, M.C_COPIES).foreground().color().name() == M.INHERIT_FG.lower(),
          w.table.item(0, M.C_COPIES).foreground().color().name())
    check("顶部摘要条有内容", "生效参数" in w.eff_lbl.text(), w.eff_lbl.text())

    print("\n[2] 改全局份数 -> 各行立刻变")
    w.g_copies.setValue(7)
    pump(app)
    check("第1行份数变成 7", cell(w, 0, M.C_COPIES) == "7", cell(w, 0, M.C_COPIES))
    check("第3行份数变成 7", cell(w, 2, M.C_COPIES) == "7", cell(w, 2, M.C_COPIES))
    check("摘要条显示 7 份", "7 份" in w.eff_lbl.text(), w.eff_lbl.text())

    print("\n[3] 改全局色彩/双面/页码 -> 各行立刻变")
    w.g_color.setCurrentText("黑白")
    w.g_duplex.setCurrentText("双面(长边)")
    w.g_pages.setText("1-3")
    pump(app)
    check("色彩列变成 黑白", cell(w, 1, M.C_COLOR) == "黑白", cell(w, 1, M.C_COLOR))
    check("双面列变成 双面长", cell(w, 1, M.C_DUPLEX) == "双面长", cell(w, 1, M.C_DUPLEX))
    check("页码列变成 1-3", cell(w, 1, M.C_PAGES) == "1-3", cell(w, 1, M.C_PAGES))
    check("摘要条含 黑白/双面/页码",
          all(s in w.eff_lbl.text() for s in ("黑白", "双面(长边)", "页码 1-3")),
          w.eff_lbl.text())

    print("\n[4] 单个文件单独设置 -> 黑色加粗, 且不再跟随全局")
    w.items[0]["copies"] = 2
    w.items[0]["color"] = "color"
    w._refresh_effective_now()
    check("单独设置的份数显示 2", cell(w, 0, M.C_COPIES) == "2", cell(w, 0, M.C_COPIES))
    check("单独设置的行加粗", w.table.item(0, M.C_COPIES).font().bold())
    check("单独设置的颜色为黑色",
          w.table.item(0, M.C_COPIES).foreground().color().name() == M.OVERRIDE_FG.lower(),
          w.table.item(0, M.C_COPIES).foreground().color().name())
    check("未设置的行仍是灰色",
          w.table.item(1, M.C_COPIES).foreground().color().name() == M.INHERIT_FG.lower())
    check("摘要条统计单独设置数量", "1 个有自己的单独设置" in w.eff_lbl.text(),
          w.eff_lbl.text())
    opts = w._opts_for_item(w.items[0])
    check("单文件参数: 份数用 2 而非全局 7", opts["copies"] == 2, str(opts.get("copies")))
    check("单文件参数: 色彩用 彩色 而非全局 黑白", opts["color"] == "color", str(opts.get("color")))

    print("\n[5] 单文件「跟随打印机」不被全局覆盖")
    w.items[1]["color"] = ""            # 明确选「跟随打印机」
    w.items[1]["duplex"] = ""
    w._refresh_effective_now()
    check("显示 驱动", cell(w, 1, M.C_COLOR) == "驱动", cell(w, 1, M.C_COLOR))
    opts1 = w._opts_for_item(w.items[1])
    check("传给打印层的色彩为空 (走驱动默认)", not opts1["color"], str(opts1.get("color")))
    check("传给打印层的双面为空 (走驱动默认)", not opts1["duplex"], str(opts1.get("duplex")))

    print("\n[6] 单文件参数取界面实时值 (不经过保存)")
    w.g_copies.setValue(9)
    pump(app, 120)
    opts2 = w._opts_for_item(w.items[2])
    check("未保存也立刻用 9 份", opts2["copies"] == 9, str(opts2.get("copies")))
    check("沿用的仍是未单独设置状态", w.items[2]["copies"] is None)

    print("\n[7] 纸型/逐页进摘要条")
    if w.g_per_page.isEnabled():
        w.g_per_page.setChecked(True)
        w.g_gap.setValue(3.0)
        pump(app, 120)
        check("摘要条出现逐页", "逐页 3 秒/页" in w.eff_lbl.text(), w.eff_lbl.text())
        check("单文件参数带逐页", w._opts_for_item(w.items[2])["per_page"] is True)
        w.g_per_page.setChecked(False)
        pump(app, 120)
        check("关掉后摘要条不再显示逐页", "逐页" not in w.eff_lbl.text(), w.eff_lbl.text())

    print("\n[8] 清空列表后摘要条复位")
    w.clear_list()
    pump(app, 120)
    check("队列为空提示", "队列为空" in w.eff_lbl.text(), w.eff_lbl.text())

    print("\n[9] 每次启动重置为默认设置")
    # 模拟上次会话留下的一堆非默认设置
    with open(M.CONFIG_PATH, "w", encoding="utf-8") as f:
        import json as _json
        _json.dump({
            "printer": "某台上次选的打印机",
            "global": {"copies": 9, "color": "mono", "duplex": "duplex_long",
                       "pages": "2-4", "paper": 8, "source": 3,
                       "per_page": True, "page_gap": 5.0},
            "layout": {"word_margins_mm": [5, 5, 5, 5], "excel_fit_width": True},
            "interval": 9.0,
            # 以下应当被保留
            "sumatra_path": "C:/tools/SumatraPDF.exe",
            "ext_map": {"dat": "txt"},
            "office_com": False,
            "preview_show": False,
            "preview_width": 420,
        }, f)
    cfg = M.load_config()
    check("打印机回到默认 (跟随系统)", cfg["printer"] == "", str(cfg.get("printer")))
    check("份数回到 1", cfg["global"]["copies"] == 1, str(cfg["global"]["copies"]))
    check("色彩回到空 (跟随打印机)", cfg["global"]["color"] == "", str(cfg["global"]["color"]))
    check("双面回到空", cfg["global"]["duplex"] == "", str(cfg["global"]["duplex"]))
    check("页码范围回到空", cfg["global"]["pages"] == "", str(cfg["global"]["pages"]))
    check("纸型回到 0 (跟随打印机)", cfg["global"]["paper"] == 0, str(cfg["global"]["paper"]))
    check("纸张来源回到 0", cfg["global"]["source"] == 0, str(cfg["global"]["source"]))
    check("逐页关闭", cfg["global"]["per_page"] is False, str(cfg["global"]["per_page"]))
    check("版面回到默认", cfg["layout"] == M.DEFAULT_LAYOUT_CFG, str(cfg["layout"]))
    check("文件间隔回到 2.0", cfg["interval"] == 2.0, str(cfg["interval"]))
    check("Sumatra 路径保留", cfg["sumatra_path"] == "C:/tools/SumatraPDF.exe",
          str(cfg.get("sumatra_path")))
    check("扩展名映射保留", cfg["ext_map"] == {"dat": "txt"}, str(cfg.get("ext_map")))
    check("Office 静默开关保留", cfg["office_com"] is False, str(cfg.get("office_com")))
    check("预览面板状态保留", cfg["preview_show"] is False and cfg["preview_width"] == 420,
          str((cfg.get("preview_show"), cfg.get("preview_width"))))

    w.deleteLater()
    print(f"\n=== 通过 {PASS} 项, 失败 {FAIL} 项 ===")
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
