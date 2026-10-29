# -*- coding: utf-8 -*-
"""test_core.py - 核心逻辑自测 (不实际发送打印任务)"""
import os
import sys
import zipfile
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

ok = 0
fail = 0


def check(name, cond, detail=""):
    global ok, fail
    if cond:
        ok += 1
        print(f"  PASS  {name}")
    else:
        fail += 1
        print(f"  FAIL  {name} {detail}")


# 1. 模块导入
print("[1] 模块导入")
import printer_core as pc
import archive_util as au
import main as mn
check("printer_core 导入", True)
check("archive_util 导入", True)
check("main (GUI) 导入", True)

# 2. 页码范围解析
print("[2] 页码范围解析")
check("1-3,5,8-10", pc.parse_page_range("1-3,5,8-10") == {1, 2, 3, 5, 8, 9, 10})
check("空 -> None", pc.parse_page_range("") is None)
check("None -> None", pc.parse_page_range(None) is None)
check("中文逗号", pc.parse_page_range("2，4") == {2, 4})

# 3. 类别解析 + 扩展名映射
print("[3] 类别解析与扩展名映射")
check("pdf", pc.resolve_category("pdf", {}) == pc.CAT_PDF)
check("docx", pc.resolve_category("docx", {}) == pc.CAT_OFFICE)
check("jpg", pc.resolve_category("jpg", {}) == pc.CAT_IMAGE)
check("txt", pc.resolve_category("txt", {}) == pc.CAT_TEXT)
check("prn", pc.resolve_category("prn", {}) == pc.CAT_RAW)
check("dat 别名->txt", pc.resolve_category("dat", {"dat": "txt"}) == pc.CAT_TEXT)
check("emb 别名->pdf", pc.resolve_category("emb", {"emb": "pdf"}) == pc.CAT_PDF)
check("未知类别", pc.resolve_category("xyz", {}) is None)
exts = pc.all_known_exts({"dat": "txt"})
check("已知扩展名包含映射别名", ".dat" in exts and ".zip" in exts)

# 4. 打印机枚举
print("[4] 打印机枚举")
ps = pc.list_printers()
check("枚举到打印机", len(ps) > 0, str(ps))
print(f"    打印机: {ps}")
dp = pc.get_default_printer()
print(f"    默认打印机: {dp}")

# 5. DEVMODE 可写性 (不 SetPrinter, 只验证字段存在)
print("[5] DEVMODE 字段")
if ps:
    h = pc.win32print.OpenPrinter(ps[0])
    dm = pc.win32print.GetPrinter(h, 2)["pDevMode"]
    dm.Copies = 3
    dm.Fields |= pc.DM_COPIES
    check("DEVMODE Copies 可写", dm.Copies == 3)
    check("默认打印机 DEVMODE 获取", True)
    pc.win32print.ClosePrinter(h)

# 6. 压缩包解压
print("[6] 压缩包解压")
td = tempfile.mkdtemp()
src = os.path.join(td, "src")
os.makedirs(src)
with open(os.path.join(src, "a.txt"), "w") as f:
    f.write("hello print\n" * 20)
with open(os.path.join(src, "b.pdf"), "wb") as f:
    f.write(b"%PDF-1.4 fake")
with open(os.path.join(src, "c.dat"), "w") as f:
    f.write("mapped file")

zpath = os.path.join(td, "t.zip")
with zipfile.ZipFile(zpath, "w") as z:
    for fn in os.listdir(src):
        z.write(os.path.join(src, fn), "sub/" + fn)

logmsg = []
dest, files = au.extract_archive(zpath, td, lambda m: logmsg.append(m))
check("zip 解压出 3 个文件", len(files) == 3, str(files))
check("zip 保留子目录结构", any(f.replace("/", os.sep).endswith(os.path.join("sub", "a.txt")) for f in files))

# 7z
try:
    import py7zr
    z7 = os.path.join(td, "t.7z")
    with py7zr.SevenZipFile(z7, "w") as z:
        z.writeall(src, "base")
    dest7, files7 = au.extract_archive(z7, td, lambda m: logmsg.append(m))
    check("7z 解压出 3 个文件", len(files7) == 3, str(len(files7)))
except ImportError:
    print("  SKIP  py7zr 未安装")

# tar.gz
import tarfile
tgz = os.path.join(td, "t.tar.gz")
with tarfile.open(tgz, "w:gz") as t:
    t.add(src, arcname="tg")
destt, filest = au.extract_archive(tgz, td, lambda m: logmsg.append(m))
check("tar.gz 解压", len(filest) == 3, str(len(filest)))

# rar 提示 (无 unrar 时应给出友好错误)
rpath = os.path.join(td, "fake.rar")
open(rpath, "wb").write(b"not a real rar")
try:
    au.extract_archive(rpath, td, lambda m: None)
    check("rar 错误处理", True)
except au.ArchiveError as e:
    check("rar 错误处理 (友好报错)", "unrar" in str(e) or "7z" in str(e) or "失败" in str(e), str(e))
except Exception as e:
    check("rar 错误处理", False, repr(e))

au.cleanup_dir(dest)
check("临时目录清理", not os.path.isdir(dest))

# 7. Sumatra 检测
print("[7] SumatraPDF 检测")
print(f"    找到: {pc.find_sumatra('')}")

# 8. 模拟一次 enqueue 逻辑 (不开 GUI)
print("[8] 文件归类统计")
app_cfg = {"dat": "txt"}
check("c.dat 映射为 text", pc.resolve_category("dat", app_cfg) == pc.CAT_TEXT)

print()
print(f"结果: {ok} 通过, {fail} 失败")
sys.exit(1 if fail else 0)
