# -*- coding: utf-8 -*-
"""预览面板自测 (不需要打印机, 不需要 Office)

覆盖:
  1. 缓存键: 同一文件+同一版面稳定; 版面或修改时间变化则重新生成
  2. 预览分派: PDF / 图片 / 不支持的格式
  3. 面板显示: 图片预览能渲染, 换文件能刷新, clear() 能复位

运行: python test_preview.py
"""

import os
import sys
import time
import tempfile

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")   # 无桌面环境也能跑

from PySide6.QtWidgets import QApplication                     # noqa: E402
from PIL import Image                                          # noqa: E402

import preview_ui as P                                         # noqa: E402

PASS = FAIL = 0


def check(name, cond, extra=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print("  PASS", name)
    else:
        FAIL += 1
        print("  FAIL", name, extra)


def main():
    tmp = tempfile.mkdtemp(prefix="bp_test_preview_")
    img_path = os.path.join(tmp, "pic.png")
    Image.new("RGB", (600, 800), (240, 240, 250)).save(img_path)
    txt_path = os.path.join(tmp, "note.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("hello")

    print("== 1) 缓存键 ==")
    lay1 = {"margins_mm": [10, 10, 10, 10], "excel_fit": "width"}
    lay2 = {"margins_mm": [10, 10, 10, 10], "excel_fit": "page"}
    a = P._cache_path(img_path, lay1, "excel")
    b = P._cache_path(img_path, dict(lay1), "excel")
    c = P._cache_path(img_path, lay2, "excel")
    check("同文件同版面 -> 同一缓存", a == b)
    check("版面变化 -> 缓存不同", a != c)
    check("缓存目录在临时目录下", os.path.dirname(a) == P.PREVIEW_CACHE_DIR)
    old = os.path.getmtime(img_path)
    os.utime(img_path, (old + 5, old + 5))
    d = P._cache_path(img_path, lay1, "excel")
    check("修改时间变化 -> 缓存不同", a != d)

    print("== 2) 预览分派 ==")
    kind, payload, note = P.PreviewLoader._prepare(img_path, {}, {})
    check("图片 -> image 类型", kind == "image" and payload == img_path, (kind, payload))
    try:
        P.PreviewLoader._prepare(txt_path, {}, {})
        check("文本 -> 明确报不支持", False)
    except RuntimeError as e:
        check("文本 -> 明确报不支持", "不支持预览" in str(e), str(e))

    print("== 3) 面板显示 ==")
    app = QApplication.instance() or QApplication(sys.argv)
    win = P.PreviewPanel()
    seen = {}
    win.set_file(img_path, {"copies": 2, "layout": lay1}, {}, "测试打印机")
    t0 = time.time()
    while win._pages == 0 and time.time() - t0 < 20:
        app.processEvents()
        time.sleep(0.02)
    pm = win.canvas.pixmap()
    check("图片预览已渲染", win._pages == 1 and pm is not None and not pm.isNull())
    check("页码显示 1/1", win.page_lbl.text() == "1/1", win.page_lbl.text())
    check("信息行含页数与份数",
          "1 页" in win.info.text() and "2 份" in win.info.text(), win.info.text())
    check("打印机名在提示里", "测试打印机" in win.info.toolTip(), win.info.toolTip())

    # 换成不支持的格式 -> 显示提示而不是崩
    win.set_file(txt_path, {}, {}, "测试打印机")
    t0 = time.time()
    while time.time() - t0 < 10 and "不支持" not in win.info.text():
        app.processEvents()
        time.sleep(0.02)
    check("不支持格式给出提示", "不支持预览" in win.info.text(), win.info.text())
    check("失败时页码复位", win.page_lbl.text() == "—", win.page_lbl.text())

    win.clear()
    check("clear() 复位标题", win.title_lbl.text() == "打印预览", win.title_lbl.text())
    check("clear() 复位信息", win.info.text() == "", win.info.text())

    # 打印按钮回调
    called = {}
    win.print_handler = lambda p, o: called.update(path=p, opts=o)
    win.set_file(img_path, {"copies": 1}, {}, "测试打印机")
    t0 = time.time()
    while win._pages == 0 and time.time() - t0 < 20:
        app.processEvents()
        time.sleep(0.02)
    win._print()
    check("打印按钮回传文件与参数",
          called.get("path") == img_path and called.get("opts") == {"copies": 1}, str(called))

    win.shutdown()
    print("\n结果: %d 通过, %d 失败" % (PASS, FAIL))
    return 0 if FAIL == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
