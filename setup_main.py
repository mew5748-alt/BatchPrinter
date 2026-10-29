# -*- coding: utf-8 -*-
"""
setup_main.py - BatchPrinter 安装器 (纯标准库, 由 PyInstaller 打成单文件 setup exe)
- 解压内嵌 payload.zip 到 %LocalAppData%\\Programs\\BatchPrinter
- 创建 开始菜单 + 桌面 快捷方式
- 生成 uninstall.cmd 卸载脚本
"""
import os
import sys
import zipfile
import ctypes


def msg(text, title="BatchPrinter 安装程序", style=0x40):  # 0x40=信息
    ctypes.windll.user32.MessageBoxW(0, text, title, style)


def ps(script):
    """执行 PowerShell 单行命令 (创建快捷方式等)"""
    import subprocess
    return subprocess.run(["powershell", "-NoProfile", "-Command", script],
                          capture_output=True, timeout=60).returncode == 0


def main():
    dest_root = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs")
    dest = os.path.join(dest_root, "BatchPrinter")

    payload = os.path.join(getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__))),
                           "payload.zip")
    if not os.path.isfile(payload):
        msg("安装包损坏: 缺少内嵌文件 payload.zip", style=0x10)
        sys.exit(1)

    # 1. 关闭正在运行的实例
    import subprocess
    subprocess.run(["taskkill", "/IM", "BatchPrinter.exe", "/F"],
                   capture_output=True)
    import time
    time.sleep(1)

    # 2. 解压 (覆盖安装, 保留用户 config.json)
    try:
        os.makedirs(dest, exist_ok=True)
        with zipfile.ZipFile(payload) as z:
            total = len(z.namelist())
            for i, info in enumerate(z.infolist(), 1):
                if info.is_dir():
                    continue
                target = os.path.join(dest, info.filename)
                os.makedirs(os.path.dirname(target) or dest, exist_ok=True)
                with z.open(info) as src, open(target, "wb") as dst:
                    dst.write(src.read())
    except Exception as e:
        msg(f"安装失败: {e}", style=0x10)
        sys.exit(1)

    # 3. 卸载脚本 (ASCII, 保留用户配置的提示写在 README)
    uninstall = r"""@echo off
setlocal
set "DEST=%LocalAppData%\Programs\BatchPrinter"
taskkill /IM BatchPrinter.exe /F >nul 2>&1
timeout /t 1 /nobreak >nul
rmdir /S /Q "%DEST%" >nul 2>&1
del "%AppData%\Microsoft\Windows\Start Menu\Programs\BatchPrinter.lnk" >nul 2>&1
powershell -NoProfile -Command "$d=[Environment]::GetFolderPath('Desktop'); Remove-Item -Force ($d+'\BatchPrinter.lnk') -ErrorAction SilentlyContinue"
echo BatchPrinter uninstalled.
pause
"""
    try:
        with open(os.path.join(dest, "uninstall.cmd"), "w", encoding="ascii", newline="\r\n") as f:
            f.write(uninstall)
    except Exception:
        pass

    # 4. 快捷方式 (当前用户, 无需管理员)
    exe = os.path.join(dest, "BatchPrinter.exe")
    ok1 = ps(f"$ws=New-Object -ComObject WScript.Shell; "
             f"$s=$ws.CreateShortcut($env:AppData+'\\Microsoft\\Windows\\Start Menu\\Programs\\BatchPrinter.lnk'); "
             f"$s.TargetPath='{exe}'; $s.WorkingDirectory='{dest}'; $s.Save()")
    ok2 = ps(f"$ws=New-Object -ComObject WScript.Shell; "
             f"$d=[Environment]::GetFolderPath('Desktop'); "
             f"$s=$ws.CreateShortcut($d+'\\BatchPrinter.lnk'); "
             f"$s.TargetPath='{exe}'; $s.WorkingDirectory='{dest}'; $s.Save()")

    sc = "已创建开始菜单和桌面快捷方式" if (ok1 and ok2) else "快捷方式创建失败, 可手动运行程序目录中的 BatchPrinter.exe"
    msg(f"BatchPrinter 安装完成!\n\n安装位置: {dest}\n{sc}\n\n"
        f"卸载方法: 运行安装目录中的 uninstall.cmd", style=0x40)


if __name__ == "__main__":
    main()
