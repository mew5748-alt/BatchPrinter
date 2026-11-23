# -*- coding: utf-8 -*-
"""多图拼版(双面)功能的自检 —— 全程不真正打印, 只验证页数/参数计算.

运行: python test_merge.py
"""

import os
import tempfile

import win32con
import win32print
import win32ui

from PIL import Image

import printer_core as pc


def _make_images(dirpath):
    paths = []
    # 5 张不同尺寸/方向的图片
    specs = [("a.jpg", (1600, 1200), (220, 60, 60)),
             ("b.jpg", (1200, 1600), (60, 140, 220)),
             ("c.png", (900, 900), (60, 170, 90)),
             ("d.jpg", (3000, 800), (200, 170, 40)),
             ("e.png", (800, 2000), (140, 70, 180))]
    for name, size, color in specs:
        p = os.path.join(dirpath, name)
        Image.new("RGB", size, color).save(p)
        paths.append(p)
    # 多帧 tif: 1 个文件 = 3 页
    tif = os.path.join(dirpath, "multi.tif")
    frames = [Image.new("RGB", (400, 300), (i * 60, 80, 160)) for i in range(3)]
    frames[0].save(tif, save_all=True, append_images=frames[1:])
    return paths, tif


def test_frame_count(paths, tif):
    assert pc.image_page_count(paths[0]) == 1, "单帧图片应为 1 页"
    n = pc.image_page_count(tif)
    assert n == 3, "多帧 TIFF 应为 3 页, 实际 %d" % n
    print("✓ 页数统计: 单帧=1, 多帧 TIFF=%d" % n)


def test_grid_shape():
    assert pc._grid_shape(1) == (1, 1)
    assert pc._grid_shape(2) == (1, 2)
    assert pc._grid_shape(4) == (2, 2)
    print("✓ 每面排布: 1张→1x1, 2张→1x2, 4张→2x2")


def test_page_split(paths, tif):
    """5 张单帧 + 1 个 3 帧 tif = 8 帧 -> 页数/纸数"""
    frames = sum(pc.image_page_count(p) for p in paths) + pc.image_page_count(tif)
    assert frames == 8, frames
    for per_side, want_pages in ((1, 8), (2, 4), (4, 2)):
        pages = (frames + per_side - 1) // per_side
        assert pages == want_pages, (per_side, pages)
        sheets = (pages + 1) // 2
        print(f"✓ 每面 {per_side} 张: {frames} 张图 → {pages} 页 → 双面约 {sheets} 张纸")


def test_devmode_duplex():
    for key, want in (("duplex_long", pc.DMDUP_VERTICAL),
                      ("duplex_short", pc.DMDUP_HORIZONTAL),
                      ("simplex", pc.DMDUP_SIMPLEX)):
        _c, _col, dup = pc.settings_to_devmode_values(1, "mono", key)
        assert dup == want, (key, dup)
    print("✓ 设置映射: 双面长边/短边/单面 → DEVMODE 值正确")


def test_printer_dc():
    """真实打印机 DC 上验证可打印区与卡片尺寸 (只创建 DC, 不发送作业)"""
    printer = pc.get_default_printer_name()
    if not printer:
        print("· 跳过 DC 检查: 没有默认打印机")
        return
    print(f"· 当前默认打印机: {printer}")
    dm = win32print.GetPrinter(win32print.OpenPrinter(printer), 2)["pDevMode"]
    pc._modify_devmode(dm, copies=1, duplex=pc.DMDUP_VERTICAL,
                       paper_size=9 if hasattr(pc, "_paper_size_to_int") else None)
    print("   DEVMODE.Duplex =", int(dm.Duplex), "(2=长边双面)")
    assert int(dm.Duplex) == pc.DMDUP_VERTICAL

    dc = pc._open_printer_dc(printer, copies=1, duplex=pc.DMDUP_VERTICAL)
    try:
        dpi = dc.GetDeviceCaps(win32con.LOGPIXELSX) or 600
        pw = dc.GetDeviceCaps(win32con.HORZRES)
        ph = dc.GetDeviceCaps(win32con.VERTRES)
        print(f"   可打印区: {pw}x{ph}px @ {dpi}dpi "
              f"({pw * 25.4 / dpi:.0f}x{ph * 25.4 / dpi:.0f}mm)")
        for per_side in (1, 2, 4):
            cols, rows = pc._grid_shape(per_side)
            margin = int(dpi * 4 / 25.4)
            gap = int(dpi * 3 / 25.4)
            cw = (pw - margin * 2 - gap * (cols - 1)) // cols
            ch = (ph - margin * 2 - gap * (rows - 1)) // rows
            assert cw > 0 and ch > 0, (per_side, cw, ch)
            print(f"   每面 {per_side} 张: 单元格 {cw}x{ch}px "
                  f"({cw * 25.4 / dpi:.0f}x{ch * 25.4 / dpi:.0f}mm)")
    finally:
        dc.DeleteDC()

    sup = pc.printer_supports_duplex(printer)
    print("   自动双面支持:", {True: "支持", False: "不支持", None: "驱动未报告"}[sup])


def main():
    with tempfile.TemporaryDirectory() as td:
        paths, tif = _make_images(td)
        test_frame_count(paths, tif)
        test_grid_shape()
        test_page_split(paths, tif)
        test_devmode_duplex()
        test_printer_dc()
    print("\n全部检查通过 (未真正打印, 不耗纸)")


if __name__ == "__main__":
    main()
