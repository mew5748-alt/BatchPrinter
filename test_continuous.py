# -*- coding: utf-8 -*-
"""连续打印 (跨文件接续) 自检 —— 只验证分页/渲染, 不发送任何打印作业.

运行: python test_continuous.py
"""

import os
import tempfile

from PIL import Image

import printer_core as pc


def _mk(dirpath):
    paths = {}
    # 3 张单页图片
    ps = []
    for i, n in enumerate(("a.jpg", "b.png", "c.jpg")):
        p = os.path.join(dirpath, n)
        Image.new("RGB", (1200, 900), (i * 70, 90, 160)).save(p)
        ps.append(p)
    paths["images"] = ps
    # 3 帧 tif
    tif = os.path.join(dirpath, "multi.tif")
    fr = [Image.new("RGB", (400, 300), (40, i * 70, 120)) for i in range(3)]
    fr[0].save(tif, save_all=True, append_images=fr[1:])
    paths["tif"] = tif
    # 文本 (够排 2 页)
    txt = os.path.join(dirpath, "note.txt")
    with open(txt, "w", encoding="utf-8") as f:
        f.write("\n".join("line %03d" % i for i in range(400)))
    paths["txt"] = txt
    # 3 页 PDF
    pdf = os.path.join(dirpath, "doc.pdf")
    pages = [Image.new("RGB", (1240, 1754), (250, 250 - i * 40, 200)) for i in range(3)]
    pages[0].save(pdf, save_all=True, append_images=pages[1:], resolution=150)
    paths["pdf"] = pdf
    return paths


def count(job, pw=1654, ph=2339, dpi=200):
    n = 0
    for _ in pc._iter_job_pages(job, pw, ph, dpi, lambda m: None):
        n += 1
    return n


def test_pages(paths):
    log = lambda m: None
    n1 = count({"path": paths["images"][0], "cat": pc.CAT_IMAGE, "pages": None})
    assert n1 == 1, "单图应为 1 页"
    print("✓ 单张图片: 1 页")

    n2 = count({"path": paths["tif"], "cat": pc.CAT_IMAGE, "pages": None})
    assert n2 == 3, "3 帧 TIFF 应为 3 页, 实际 %d" % n2
    print("✓ 多帧 TIFF: 3 页")

    n3 = count({"path": paths["txt"], "cat": pc.CAT_TEXT, "pages": None})
    assert n3 >= 2, "长文本应分成多页, 实际 %d" % n3
    print("✓ 文本: %d 页" % n3)

    n4 = count({"path": paths["pdf"], "cat": pc.CAT_PDF, "pages": None})
    assert n4 == 3, "3 页 PDF 应为 3 页, 实际 %d" % n4
    print("✓ PDF: 3 页")


def test_page_range(paths):
    n = count({"path": paths["pdf"], "cat": pc.CAT_PDF, "pages": {2}})
    assert n == 1, "页码范围只留第2页, 实际 %d" % n
    n = count({"path": paths["tif"], "cat": pc.CAT_IMAGE, "pages": {1, 3}})
    assert n == 2
    print("✓ 页码范围过滤生效 (PDF 只打第2页 / TIFF 只打 1、3 帧)")


def test_text_image_pages(paths):
    """文本页图像应该是整页尺寸 (不是打印机 600dpi 的近亿像素)"""
    job = {"path": paths["txt"], "cat": pc.CAT_TEXT, "pages": {1}}
    for im in pc._iter_job_pages(job, 1654, 2339, 200, lambda m: None):
        assert im.size == (1654, 2339), im.size
        assert im.mode == "RGB"
        break
    print("✓ 文本页尺寸 = 渲染尺寸 1654x2339 (不是 600dpi 巨图)")


def test_flow(paths):
    """模拟接续: 1 页图 + 1 页图 + 1 页图 = 3 页, 双面 2 张纸 (单独打则要 3 张)"""
    jobs = [{"path": p, "cat": pc.CAT_IMAGE, "pages": None} for p in paths["images"]]
    total = sum(count(j) for j in jobs)
    assert total == 3
    sheets = (total + 1) // 2
    print(f"✓ 接续 3 个单页文件 = {total} 页 → 双面 {sheets} 张纸 "
          f"(各打各的需 {len(jobs)} 张纸, 省 {len(jobs) - sheets} 张)")

    # 每面 2 张拼版
    per_side = 2
    sides = (total + per_side - 1) // per_side
    print(f"✓ 每面 2 张拼版 = {sides} 页 → 双面 {(sides + 1) // 2} 张纸")


def test_mixed_order(paths):
    """混合格式接续: 图片 + PDF + 文本, 页序必须与列表顺序一致"""
    jobs = [{"path": paths["images"][0], "cat": pc.CAT_IMAGE, "pages": None},
            {"path": paths["pdf"], "cat": pc.CAT_PDF, "pages": None},
            {"path": paths["txt"], "cat": pc.CAT_TEXT, "pages": {1}}]
    seq = []
    for j in jobs:
        seq.append((os.path.basename(j["path"]),
                    sum(1 for _ in pc._iter_job_pages(j, 1654, 2339, 200, lambda m: None))))
    assert seq[0][1] == 1 and seq[1][1] == 3 and seq[2][1] == 1, seq
    print("✓ 混合接续顺序正确:", " + ".join("%s(%d页)" % s for s in seq))


def main():
    with tempfile.TemporaryDirectory() as td:
        paths = _mk(td)
        test_pages(paths)
        test_page_range(paths)
        test_text_image_pages(paths)
        test_flow(paths)
        test_mixed_order(paths)
    print("\n全部检查通过 (未发送打印作业, 不耗纸)")


if __name__ == "__main__":
    main()
