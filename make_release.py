# -*- coding: utf-8 -*-
"""make_release.py - 创建/更新 GitHub Release 并上传附件 (走 api.github.com, 避开被拦截的 git 协议)"""
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import pathlib

TOKEN = os.environ.get("GH_TOKEN") or pathlib.Path(
    os.environ.get("TEMP", "/tmp"), "_gh_token.txt").read_text(encoding="utf-8").strip()
REPO = "mew5748-alt/BatchPrinter"
API = f"https://api.github.com/repos/{REPO}"
UPLOADS = f"https://uploads.github.com/repos/{REPO}"
TAG = sys.argv[1] if len(sys.argv) > 1 else "v1.0.2"

HDRS = {
    "Authorization": f"Bearer {TOKEN}",
    "Accept": "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28",
}

BODY = """Windows 批量打印助手 v1.0.2。

## 本次更新

**修复：选择了其他打印机却仍然打到默认打印机**
Office / 未知格式走的 Windows Shell `print` 动词无法指定打印机，现在批次打印期间会把系统默认打印机临时切到你所选的那台，整批结束后自动恢复；切换失败时自动改用 `printto` 动词显式指定打印机（打印机名 + 驱动名 + 端口名）。

同时包含 v1.0.1 的全部改动：
- 新增拖动行排序、上移/下移、按文件名 / 修改时间（早→晚、晚→早）排序
- 新增多选批量设置份数 / 色彩 / 双面 / 页码范围（只改填写的项，其余保持）
- 文件夹与压缩包内文件改为资源管理器自然排序（1、2、10 而不是 1、10、2）
- 修复状态列错位、批量设置按钮点击无反应、删除后序号断号等问题

## 下载

- `BatchPrinter-Setup.exe` — 安装版（免管理员权限，自动创建开始菜单 + 桌面快捷方式，含卸载脚本）
- `BatchPrinter-portable.zip` — 绿色便携版（解压即用）
- `SHA256SUMS.txt` — 校验值

## 打印引擎

- PDF：内置 PDFium 渲染直印（份数 / 色彩 / 双面 / 页码范围全部精确生效）
- 图片 / 文本：GDI 直印，全参数生效
- Office / 其他：系统关联程序打印，期间临时切换默认打印机以确保打到所选设备
- prn / ps / pcl：字节流直送

系统要求：Windows 10 / 11（64 位）
"""

ASSETS = [
    ("dist/BatchPrinter-Setup.exe", "BatchPrinter-Setup.exe"),
    ("BatchPrinter-portable.zip", "BatchPrinter-portable.zip"),
    ("SHA256SUMS.txt", "SHA256SUMS.txt"),
]


def req(method, url, data=None, retries=4):
    body = None
    headers = dict(HDRS)
    if data is not None:
        body = json.dumps(data).encode()
        headers["Content-Type"] = "application/json"
    for i in range(retries):
        try:
            r = urllib.request.Request(url, method=method, data=body, headers=headers)
            with urllib.request.urlopen(r, timeout=90) as resp:
                raw = resp.read()
            return json.loads(raw) if raw else {}
        except urllib.error.HTTPError:
            raise
        except Exception as e:
            print(f"  网络抖动重试 {i+1}/{retries}: {e}")
            time.sleep(2 + 2 * i)
    raise RuntimeError("请求失败")


def upload(url, path, retries=3):
    size = os.path.getsize(path)
    with open(path, "rb") as f:
        data = f.read()
    headers = dict(HDRS)
    headers["Content-Type"] = "application/octet-stream"
    for i in range(retries):
        try:
            r = urllib.request.Request(url, method="POST", data=data, headers=headers)
            with urllib.request.urlopen(r, timeout=580) as resp:
                return json.loads(resp.read())
        except urllib.error.HTTPError as e:
            msg = e.read().decode()[:200]
            if e.code == 422 and "already_exists" in msg:
                print(f"  附件已存在, 跳过: {os.path.basename(path)}")
                return {"state": "exists"}
            print(f"  上传失败({e.code}): {msg}")
        except Exception as e:
            print(f"  网络中断, 重试 {i+1}/{retries}: {e}")
        time.sleep(3 + 3 * i)
    raise RuntimeError(f"上传失败: {path} ({size/1024/1024:.1f} MB)")


def main():
    # 1. 找已存在的 release
    rel = None
    try:
        rel = req("GET", f"{API}/releases/tags/{TAG}")
        print("Release 已存在, id =", rel["id"])
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise
        print("创建 Release", TAG)
        rel = req("POST", f"{API}/releases", {
            "tag_name": TAG,
            "target_commitish": "main",
            "name": f"BatchPrinter {TAG}",
            "body": BODY,
            "draft": False,
            "prerelease": False,
        })
        print("  id =", rel["id"])

    # 2. 上传附件 (已存在的先删掉, 保证内容与本次构建一致)
    existing = {a["name"]: a["id"] for a in rel.get("assets", [])}
    for local, name in ASSETS:
        if name in existing:
            print(f"删除旧附件 {name}")
            try:
                req("DELETE", f"{API}/releases/assets/{existing[name]}")
            except urllib.error.HTTPError as e:
                print("  删除失败:", e.code)
        print(f"上传 {name} ({os.path.getsize(local)/1024/1024:.1f} MB)")
        r = upload(f"{UPLOADS}/releases/{rel['id']}/assets?name={urllib.parse.quote(name)}", local)
        print("  ->", r.get("state"))

    # 3. 校验
    rel = req("GET", f"{API}/releases/{rel['id']}")
    print("\nRelease:", rel["name"], "| tag:", rel["tag_name"])
    print("页面:", rel["html_url"])
    for a in rel["assets"]:
        print(f"  {a['name']}: {a['size']/1024/1024:.1f} MB, 下载数 {a['download_count']}")


if __name__ == "__main__":
    main()
