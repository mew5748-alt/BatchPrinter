# -*- coding: utf-8 -*-
"""图片方向 (纵向/横向/自动) 自检 —— 只建 DC 读尺寸, 不发送打印作业.

运行: python test_orient.py
"""

import os
import tempfile

from PIL import Image
import win32con

import printer_core as pc


def _mk(dirpath):
    land = os.path.join(dirpath, "wide.jpg")
    Image.new("RGB", (1600, 900), (200, 120, 80)).save(land)

    port = os.path.join(dirpath, "tall.jpg")
    Image.new("RGB", (900, 1600), (80, 160, 200)).save(port)

    # 存成 1600x900 但 EXIF Orientation=6 (顺时针 90°) -> 实际是竖图
    exif = Image.Exif()
    exif[0x0112] = 6
    rotated = os.path.join(dirpath, "rotated.jpg")
    Image.new("RGB", (1600, 900), (120, 200, 120)).save(rotated, exif=exif)
    return land, port, rotated


def test_orientation_of(land, port, rotated):
    assert pc.image_orientation(land) == "landscape", pc.image_orientation(land)
    assert pc.image_orientation(port) == "portrait", pc.image_orientation(port)
    got = pc.image_orientation(rotated)
    assert got == "portrait", "EXIF 旋转 90° 的图应判为竖图, 实际 %s" % got
    print("✓ 朝向识别: 横图=landscape, 竖图=portrait, EXIF旋转图=portrait")


def test_resolve(land, port):
    assert pc.resolve_orientation("auto", [land]) == pc.DMORIENT_LANDSCAPE
    assert pc.resolve_orientation("auto", [port]) == pc.DMORIENT_PORTRAIT
    assert pc.resolve_orientation("portrait", [land]) == pc.DMORIENT_PORTRAIT
    assert pc.resolve_orientation("landscape", [port]) == pc.DMORIENT_LANDSCAPE
    assert pc.resolve_orientation("auto", []) is None
    print("✓ 方向解析: 自动跟图片 / 固定纵向 / 固定横向 / 空列表=跟随驱动")


def test_devmode(land):
    """DEVMODE 里写入方向, 并确认 DC 的可打印区真的翻转了"""
    printer = pc.get_default_printer_name()
    if not printer:
        print("· 跳过 DC 检查: 没有默认打印机")
        return

    sizes = {}
    for label, orient in (("纵向", pc.DMORIENT_PORTRAIT),
                          ("横向", pc.DMORIENT_LANDSCAPE)):
        dc = pc._open_printer_dc(printer, copies=1, orientation=orient)
        try:
            w = dc.GetDeviceCaps(win32con.HORZRES)
            h = dc.GetDeviceCaps(win32con.VERTRES)
            sizes[label] = (w, h)
            print(f"  {label}: 可打印区 {w}x{h}px {'(宽>高 ✓)' if (w > h) == (orient == pc.DMORIENT_LANDSCAPE) else '(未翻转 ✗)'}")
        finally:
            dc.DeleteDC()

    pw, ph = sizes["纵向"]
    lw, lh = sizes["横向"]
    assert pw < ph, "纵向纸应 宽<高, 实际 %dx%d" % (pw, ph)
    assert lw > lh, "横向纸应 宽>高, 实际 %dx%d" % (lw, lh)
    print("✓ 方向真实生效: 纵向 %dx%d / 横向 %dx%d" % (pw, ph, lw, lh))


def test_mixed_split():
    """横竖混排时应切成两段 (一个作业内纸张方向是固定的)"""
    seq = ["landscape", "landscape", "portrait", "portrait", "landscape"]
    segs, cur, key = [], [], None
    for k in seq:
        if key is None or k == key:
            key = k
            cur.append(k)
        else:
            segs.append(cur)
            cur, key = [k], k
    if cur:
        segs.append(cur)
    assert [len(s) for s in segs] == [2, 2, 1], segs
    print("✓ 横竖混排切段: %s -> %s" % (seq, [len(s) for s in segs]))


def main():
    with tempfile.TemporaryDirectory() as td:
        land, port, rotated = _mk(td)
        test_orientation_of(land, port, rotated)
        test_resolve(land, port)
        test_devmode(land)
        test_mixed_split()
    print("\n全部检查通过 (未发送打印作业, 不耗纸)")


if __name__ == "__main__":
    main()
