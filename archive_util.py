# -*- coding: utf-8 -*-
"""archive_util.py - 压缩包自动解压 (zip/7z/rar/tar.gz 等), 无需事先手动解压"""

import os
import tarfile
import zipfile
import shutil
import subprocess

ARCHIVE_EXTS = {".zip", ".7z", ".rar", ".tar", ".gz", ".tgz", ".bz2", ".xz"}


class ArchiveError(Exception):
    pass


def _find_7z_exe():
    candidates = [
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe",
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return shutil.which("7z.exe") or shutil.which("7z")


def _safe_join(base, name):
    """防 zip-slip 路径穿越"""
    p = os.path.normpath(os.path.join(base, name))
    if not (p == os.path.normpath(base) or p.startswith(os.path.normpath(base) + os.sep)):
        raise ArchiveError(f"压缩包内路径非法: {name}")
    return p


def extract_archive(path, dest_root, log):
    """
    解压压缩包到 dest_root 下的独立子目录, 返回解压出的文件列表.
    支持: zip / 7z / rar / tar / tar.gz / tar.bz2 / tar.xz
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".gz" and not path.lower().endswith(".tar.gz"):
        raise ArchiveError("仅支持 .tar.gz, 不支持单独的 .gz")
    # 目标目录用完整文件名 (含扩展名) + pid, 避免同名不同格式压缩包互相覆盖
    arcname = os.path.basename(path)
    dest = os.path.join(dest_root, f"__extract__{arcname}_{os.getpid()}")
    os.makedirs(dest, exist_ok=True)

    if ext == ".zip":
        _extract_zip(path, dest, log)
    elif ext == ".7z":
        import py7zr
        try:
            with py7zr.SevenZipFile(path, mode="r") as z:
                z.extractall(dest)
        except py7zr.Bad7zFile as e:
            raise ArchiveError(f"7z 文件损坏或加密: {e}")
        except py7zr.exceptions.PasswordRequired:
            raise ArchiveError("该 7z 压缩包已加密, 请先手动解压")
        log(f"已解压 7z: {os.path.basename(path)} -> {dest}")
    elif ext == ".rar":
        _extract_rar(path, dest, log)
    elif ext in (".tar", ".gz", ".tgz", ".bz2", ".xz"):
        try:
            with tarfile.open(path, "r:*") as t:
                t.extractall(dest, filter="data")
        except tarfile.TarError as e:
            raise ArchiveError(f"tar 解压失败: {e}")
        log(f"已解压 tar 包: {os.path.basename(path)} -> {dest}")
    else:
        raise ArchiveError(f"不支持的压缩格式: {ext}")

    files = []
    for root, _dirs, fs in os.walk(dest):
        for f in fs:
            files.append(os.path.join(root, f))
    return dest, files


def _extract_zip(path, dest, log):
    try:
        with zipfile.ZipFile(path) as z:
            if z.testzip() is not None:
                pass  # 个别条目 CRC 有误仍尝试解压其余文件
            for info in z.infolist():
                if info.is_dir():
                    continue
                if info.flag_bits & 0x1:
                    log(f"跳过加密文件: {info.filename}")
                    continue
                target = _safe_join(dest, info.filename)
                os.makedirs(os.path.dirname(target) or dest, exist_ok=True)
                try:
                    with z.open(info) as src, open(target, "wb") as dst:
                        shutil.copyfileobj(src, dst)
                except RuntimeError:
                    log(f"跳过加密文件: {info.filename}")
    except zipfile.BadZipFile as e:
        raise ArchiveError(f"zip 文件损坏: {e}")
    log(f"已解压 zip: {os.path.basename(path)} -> {dest}")


def _extract_rar(path, dest, log):
    # 优先 rarfile (需要 unrar/unar/bsdtar), 失败则尝试 7z.exe
    try:
        import rarfile
        try:
            with rarfile.RarFile(path) as r:
                r.extractall(dest)
            log(f"已解压 rar: {os.path.basename(path)} -> {dest}")
            return
        except rarfile.Error:
            pass
    except ImportError:
        pass

    seven_zip = _find_7z_exe()
    if not seven_zip:
        raise ArchiveError(
            "解压 RAR 需要 unrar.exe 或 7-Zip (推荐安装 7-Zip: https://www.7-zip.org)")
    proc = subprocess.run([seven_zip, "x", "-y", f"-o{dest}", path],
                          capture_output=True, timeout=600)
    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout).decode("gbk", "replace").strip()[:300]
        raise ArchiveError(f"7z 解压 RAR 失败: {err}")
    log(f"已解压 rar (经 7z): {os.path.basename(path)} -> {dest}")


def is_archive(path):
    return os.path.splitext(path)[1].lower() in ARCHIVE_EXTS


def cleanup_dir(path):
    """安全删除临时解压目录"""
    if path and os.path.isdir(path) and f"__extract__" in os.path.basename(path):
        shutil.rmtree(path, ignore_errors=True)
