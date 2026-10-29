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

BODY = """Windows 批量打印助手 v1.4.0。

## 本次更新

### 一、纸型 / 纸张来源 / 逐页打印（新）

- **纸型（按纸型缩放）**：下拉项直接读取打印机驱动能力（实测 Apeos 3561 有 46 种纸型），选中后内容等比缩放到该纸型的可打印区域 —— A4 稿件选 A3 就放大铺满 A3，选 A5 就整体缩到 A5，不裁边、不缺角。实现方式是把 `dmPaperSize` 写入 DEVMODE 再 `CreateDC`，尺寸由驱动保证。
- **纸张来源**：枚举驱动上报的进纸盒（自动 / 纸盘1 / 纸盘2 / 手送盘…），通过 `dmDefaultSource` 指定，不影响打印机自身的默认设置，打完即恢复。
- **逐页打印**：每页单独作为一个打印作业发送（作业名带页码），可设置页与页之间的等待秒数。适合手动放纸/翻面、逐页核对、卡纸后按页补打。配合页码范围可以只打指定页。
- 设置栏改为可滚动，小屏或高 DPI 下"立即打印"按钮不再被裁切（窗口最小高度从 797 降到 356）。

### 二、Word / Excel 静默打印 + 版面控制（v1.1.0）

- Office 文档不再依赖系统关联程序的打印动词：**COM 只读打开 → 改版面 → 导出临时 PDF → 内置 PDF 引擎打印**。全程不弹程序窗口、不用切换系统默认打印机。
- 可调**页边距**（上下左右 mm）、**方向**（纵向/横向）、**Excel 所有列缩到一页宽**、Word 按纸型缩放。
- 只读方式打开，原文件哈希不变。
- 补充：WPS 环境下 `ActivePrinter` / `PrintOut` 的打印机参数会被静默忽略，这是上面改用"导出 PDF"路线的直接原因。

### 三、打印预览（v1.2.0 / v1.3.0）

- 单击列表中的文件即在右侧显示首页预览，后台线程转换 + 缓存，不卡界面。
- 预览区为目标界面**内嵌面板**，可一键收起/展开、宽度可拖；关闭跟随即停止自动预览。
- 工具栏重排为两行，1366×768 的机器也不再超出桌面。

## 下载

- `BatchPrinter-Setup.exe` — 安装版（免管理员权限，自动创建开始菜单 + 桌面快捷方式，含卸载脚本）
- `BatchPrinter-portable.zip` — 绿色便携版（解压即用）
- `SHA256SUMS.txt` — 校验值

## 打印引擎

| 格式 | 引擎 | 份数/色彩/双面/页码 | 纸型/进纸盒 | 逐页 |
|---|---|---|---|---|
| Word / Excel / PPT（docx/xlsx/pptx/wps/et…） | Office COM → 临时 PDF → PDFium + GDI | ✅ | ✅ | ✅ |
| PDF | 内置 PDFium 渲染直印 | ✅ | ✅ | ✅ |
| 图片 / 文本 | GDI 直印 | ✅ | ✅ | ✅ |
| prn / ps / pcl | 字节流直送 | 按内容本身 | — | — |
| 其他（扩展名映射来的格式） | 系统 Shell 打印动词 | 临时写默认值，尽力生效 | 临时写默认值 | — |

系统要求：Windows 10 / 11（64 位）。PDF 引擎已内置，无需另装 SumatraPDF。
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
