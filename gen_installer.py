# -*- coding: utf-8 -*-
"""
gen_installer.py - 用 Windows 自带 IExpress 把打包结果封装成安装 exe
产物: BatchPrinter-Setup.exe (免管理员权限, 安装到 %LocalAppData%\\Programs\\BatchPrinter)
"""
import os
import shutil
import subprocess
import sys

APP_DIR = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(APP_DIR, "installer_src")
APP_SRC = os.path.join(APP_DIR, "dist", "BatchPrinter")
APP_DST = os.path.join(SRC, "app")
SED = os.path.join(SRC, "installer.sed")
OUT = os.path.join(APP_DIR, "BatchPrinter-Setup.exe")

INSTALL_CMD = r"""@echo off
setlocal
set "DEST=%LocalAppData%\Programs\BatchPrinter"
echo Installing BatchPrinter to "%DEST%" ...
if not exist "%DEST%" mkdir "%DEST%"
xcopy /E /I /Y /Q "%~dp0app" "%DEST%" >nul
if errorlevel 1 goto :err
rem --- Start Menu shortcut (current user) ---
set "SMENU=%AppData%\Microsoft\Windows\Start Menu\Programs"
if not exist "%SMENU%\BatchPrinter" mkdir "%SMENU%\BatchPrinter"
powershell -NoProfile -Command "$ws=New-Object -ComObject WScript.Shell; $s=$ws.CreateShortcut($env:AppData+'\Microsoft\Windows\Start Menu\Programs\BatchPrinter\BatchPrinter.lnk'); $s.TargetPath=$env:LocalAppData+'\Programs\BatchPrinter\BatchPrinter.exe'; $s.WorkingDirectory=$env:LocalAppData+'\Programs\BatchPrinter'; $s.Save()"
rem --- Desktop shortcut (current user) ---
powershell -NoProfile -Command "$ws=New-Object -ComObject WScript.Shell; $d=[Environment]::GetFolderPath('Desktop'); $s=$ws.CreateShortcut($d+'\BatchPrinter.lnk'); $s.TargetPath=$env:LocalAppData+'\Programs\BatchPrinter\BatchPrinter.exe'; $s.WorkingDirectory=$env:LocalAppData+'\Programs\BatchPrinter'; $s.Save()"
echo Installation complete.
exit /b 0
:err
echo Installation failed. Please check permissions.
exit /b 1
"""

UNINSTALL_CMD = r"""@echo off
setlocal
set "DEST=%LocalAppData%\Programs\BatchPrinter"
taskkill /IM BatchPrinter.exe /F >nul 2>&1
timeout /t 1 /nobreak >nul
rmdir /S /Q "%DEST%" >nul 2>&1
del "%AppData%\Microsoft\Windows\Start Menu\Programs\BatchPrinter\BatchPrinter.lnk" >nul 2>&1
rmdir "%AppData%\Microsoft\Windows\Start Menu\Programs\BatchPrinter" >nul 2>&1
powershell -NoProfile -Command "$d=[Environment]::GetFolderPath('Desktop'); Remove-Item -Force ($d+'\BatchPrinter.lnk') -ErrorAction SilentlyContinue"
echo BatchPrinter has been uninstalled. (config.json was inside the folder and is removed too.)
pause
"""

README_TXT = r"""BatchPrinter - Batch Printing Assistant
=========================================

1. Double-click "BatchPrinter.exe" (or the Desktop/Start Menu shortcut).
2. Add files / folders / archives (zip, 7z, rar, tar are extracted automatically).
3. Pick a printer, adjust global or per-file settings, then click Print.

Global + per-file settings: copies, color mode, duplex, page range.
Extension mapping: map non-standard extensions to supported formats.

Install location : %LocalAppData%\Programs\BatchPrinter
Uninstall        : run "uninstall.cmd" in the install folder.
Config file      : config.json (next to the exe; delete to reset).

Requires Windows 10/11 64-bit. No admin rights needed.
"""


def main():
    # 1. 准备源目录
    if os.path.isdir(SRC):
        shutil.rmtree(SRC)
    os.makedirs(APP_DST, exist_ok=True)
    shutil.copytree(APP_SRC, APP_DST, dirs_exist_ok=True)
    with open(os.path.join(SRC, "install.cmd"), "w", encoding="ascii", newline="\r\n") as f:
        f.write(INSTALL_CMD)
    with open(os.path.join(APP_DST, "uninstall.cmd"), "w", encoding="ascii", newline="\r\n") as f:
        f.write(UNINSTALL_CMD)
    with open(os.path.join(APP_DST, "README.txt"), "w", encoding="ascii", newline="\r\n") as f:
        f.write(README_TXT)

    # 2. 生成 .sed (列出 app 下所有文件的相对路径, IExpress 会保留子目录)
    rels = ["install.cmd"]
    for root, _dirs, files in os.walk(APP_DST):
        for fn in sorted(files):
            p = os.path.join(root, fn)
            rels.append("app\\" + os.path.relpath(p, APP_DST).replace("/", "\\"))
    sed_lines = [
        "[Version]", "Class=IEXPRESS", "SEDVersion=1", "[Options]",
        "PackagePurpose=InstallApp",
        "ShowInstallProgramWindow=0",
        "HideExtractAnimation=1",
        "UseLongFileName=1",
        "InsideCompressed=0",
        "CAB_FixedSize=0",
        "CAB_ResvCodeSigning=0",
        "RebootMode=N",
        "InstallPrompt=",
        "DisplayLicense=",
        "FinishMessage=BatchPrinter has been installed successfully. Look for the shortcut on your desktop.",
        f"TargetName={OUT}",
        "FriendlyName=BatchPrinter 1.0 Setup",
        "AppLaunched=install.cmd",
        "PostInstallCmd=<None>",
        "AdminQuietInstallCmd=<None>",
        "UserQuietInstallCmd=<None>",
        "SourceFiles=SourceFiles",
        "[SourceFiles]",
        f"SourceFiles0={SRC}\\",
    ]
    for i, r in enumerate(rels):
        sed_lines.append(f"%FILE{i}%={r}")
    with open(SED, "w", encoding="ascii", newline="\r\n") as f:
        f.write("\n".join(sed_lines) + "\n")
    print(f"sed 生成完成: {len(rels)} 个文件")

    # 3. 调用 IExpress 静默打包
    iexpress = r"C:\Windows\System32\iexpress.exe"
    r = subprocess.run([iexpress, "/N", "/Q", SED], capture_output=True, text=True)
    print("iexpress rc =", r.returncode)
    if os.path.isfile(OUT):
        print(f"安装包已生成: {OUT} ({os.path.getsize(OUT)//1024//1024} MB)")
    else:
        print("生成失败:", r.stdout[-2000:], r.stderr[-2000:])
        sys.exit(1)


if __name__ == "__main__":
    main()
