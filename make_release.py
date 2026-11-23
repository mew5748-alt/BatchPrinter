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

BODY = """Windows 批量打印助手 v1.4.4。

## 本次更新

### 一、连续打印：文件之间不留空面（v1.4.4）

- **问题**：每个文件单独成作业时，打印机总是从新纸的正面开始 —— 前一个文件页数为奇数（1 张图片、3 页文档都很常见），那张纸的背面就空着，下一个文件又从新纸开始，中间白白浪费一整个面。
- **现在**：勾选「🔗 连续打印」，整批文件按列表顺序接续成**一个打印作业**：上一个文件的末页后面紧接下一个文件的首页，页码连续排下去，双面打印时由打印机自动正反配对。
- 支持 **图片 / PDF / Word·Excel·PPT / 文本混排**；`prn/ps/pcl` 是原始字节流无法接续，仍按原样单独发送。
- 页码设置对每个文件各自生效；份数 / 色彩 / 双面整段统一（取第一个文件）。
- 与「🖼 多图拼成一份双面文档」互斥（连续打印是它的超集），勾一个自动取消另一个。

### 二、图片方向：纵向 / 横向 / 自动（v1.4.4）

- **问题**：打印机默认纵向出纸，横向拍的照片放到竖向 A4 上只能占中间窄窄一条，其余全是白纸，看着"很小"。
- **现在**：新增「图片方向」设置，直接改 DEVMODE 的 `dmOrientation`（不动原文件、不改打印机默认设置）：
  - **自动（推荐）**：按图片宽高判断，横图用横向纸、竖图用纵向纸，各自铺满页面；手机照片先按 EXIF 摆正再判断
  - **纵向 / 横向**：强制统一方向
- 一个作业内纸张方向固定，所以「自动」遇到横竖混排时会在方向变化处分段发送（每段内部仍连续、双面照常配对）。

### 三、拖入大量文件不再卡顿（v1.4.4）

- 以前每加一个文件都要重算整张表的列宽、再遍历一遍队列重算摘要（O(n²)），几百个文件会卡住十几秒。
- 现在添加期间冻结表格重绘、列宽只测一次、摘要只算一次，最后统一刷新。**实测 800 个文件：19.7 秒 → 0.26 秒**。

### 四、多图双面：两张图片共用一张纸（v1.4.3）

- **问题**：图片只有一面，每张图各发一个打印作业时，作业里只有一页 —— 双面设置其实不生效，打印机只往每张纸的正面打一张图，背面全空，纸耗翻倍。
- **现在**：勾选「🖼 多图拼成一份双面文档」后，列表里**连续的图片按顺序合成一份文档**一次发出：图1 → 正面、图2 → 背面、图3 → 下一张纸正面……纸张直接省一半。
- **每面图片**：可选 1 张（正反各一张，不缩小）/ 2 张（一张纸 4 张）/ 4 张（一张纸 8 张，2×2 排布），进一步省纸。
- **只合并连续的图片**：中间的 PDF/Word 不参与拼版，顺序也不打乱（3 图 + 1 PDF + 2 图 → 前 3 张合成一份、PDF 单独打、后 2 张再合成一份）。
- 勾选后自动关闭「逐页打印」（逐页会把文档拆回单页作业，双面就失效了）；选了「单面」或驱动报告不支持双面时，日志里会直接提醒。
- 多帧 TIFF 按帧算页；图片按 EXIF 方向自动摆正，手机竖拍的照片不会躺倒。
- 该选项**默认关闭，勾选状态会记住**，下次打开照旧。

### 五、每次启动都回到默认状态（v1.4.2）

- **打印参数不再跨会话保留**：关掉再打开，打印机跟随系统默认，份数 / 色彩 / 双面 / 页码范围 / 纸型 / 纸张来源 / 逐页打印 / 页间等待、版面（页边距·方向·Excel 缩放）、文件间隔全部恢复初始值，不会带着上一批的参数开始下一批。
- **仍然保留的**：SumatraPDF 路径、扩展名映射、Office 静默开关、预览面板宽度与收起状态、多图双面开关 —— 这些是环境配置或使用习惯，重置要么会丢手工映射，要么纯属多余。

### 六、参数是否生效一眼可见（v1.4.1）

- 表格「份数 / 色彩 / 双面 / 页码范围」四列直接显示**实际会用的值**，不再是「全局」两个字；**灰色 = 跟随右侧全局参数，黑色加粗 = 该文件单独设置**。
- 表格上方新增**生效参数摘要条**（份数·色彩·双面·页码·纸型·进纸·逐页），改任何一个参数立刻同步到摘要和各行。
- 修复：预览面板「打印此文件」以前用的是**上次保存的旧参数**，现在始终取界面当前值 —— 这类"改了没用"的问题从根上解决。
- 修复：单文件设置里「跟随打印机」以前会被全局值覆盖；打开对话框什么都不改直接点确定，会把「跟随全局」悄悄变成「跟随打印机」。现在两种状态严格区分，批量「清除单独设置」也真正回到跟随全局。

### 七、纸型 / 纸张来源 / 逐页打印（v1.4.0）

- **纸型（按纸型缩放）**：下拉项直接读取打印机驱动能力（实测 Apeos 3561 有 46 种纸型），选中后内容等比缩放到该纸型的可打印区域 —— A4 稿件选 A3 就放大铺满 A3，选 A5 就整体缩到 A5，不裁边、不缺角。实现方式是把 `dmPaperSize` 写入 DEVMODE 再 `CreateDC`，尺寸由驱动保证。
- **纸张来源**：枚举驱动上报的进纸盒（自动 / 纸盘1 / 纸盘2 / 手送盘…），通过 `dmDefaultSource` 指定，不影响打印机自身的默认设置，打完即恢复。
- **逐页打印**：每页单独作为一个打印作业发送（作业名带页码），可设置页与页之间的等待秒数。适合手动放纸/翻面、逐页核对、卡纸后按页补打。配合页码范围可以只打指定页。
- 设置栏改为可滚动，小屏或高 DPI 下"立即打印"按钮不再被裁切（窗口最小高度从 797 降到 356）。

### 八、Word / Excel 静默打印 + 版面控制（v1.1.0）

- Office 文档不再依赖系统关联程序的打印动词：**COM 只读打开 → 改版面 → 导出临时 PDF → 内置 PDF 引擎打印**。全程不弹程序窗口、不用切换系统默认打印机。
- 可调**页边距**（上下左右 mm）、**方向**（纵向/横向）、**Excel 所有列缩到一页宽**、Word 按纸型缩放。
- 只读方式打开，原文件哈希不变。
- 补充：WPS 环境下 `ActivePrinter` / `PrintOut` 的打印机参数会被静默忽略，这是上面改用"导出 PDF"路线的直接原因。

### 九、打印预览（v1.2.0 / v1.3.0）

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
| 图片 / 文本 | GDI 直印（多图可合成一份双面拼版作业） | ✅ | ✅ | ✅ |
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
