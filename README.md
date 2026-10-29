# BatchPrinter - 批量打印助手

Windows 批量打印工具：把大量、多种格式的文档一次性发送到打印机，无需逐个打开文件。
添加文件 → 选择打印机 → 点击打印，就这么简单。

![Screenshot](https://img.shields.io/badge/platform-Windows%2010%2F11-blue) ![License](https://img.shields.io/badge/license-MIT-green)

## ✨ 核心功能

- **批量打印**：一次添加大量文件，按队列自动依次发送到打印机
- **多格式支持**：PDF、Word、Excel、PPT、RTF、图片（jpg/png/tiff…）、文本（txt/csv…）、PRN/PS/PCL
- **压缩包直打**：ZIP / 7Z / RAR / TAR.GZ 内的文件自动解压进队列，无需事先手动解压
- **文件夹递归**：添加整个文件夹（可选含子文件夹），自动筛出可打印文件
- **全局 + 个体设置**：整个列表统一设置份数、色彩模式（黑白/彩色）、单双面、页码范围；也可以**双击任意文件单独设置**
- **扩展名映射**：非标准扩展名（如 `.dat`）可映射到受支持格式（如 `txt`）间接打印
- **拖拽支持**：直接把文件/文件夹拖进窗口
- **PDF 内置引擎**：无需安装额外软件，PDF 静默直印且完整支持份数/色彩/双面/页码范围
- **免管理员权限**：安装到用户目录，不弹 UAC

## 📥 下载

| 文件 | 说明 |
|---|---|
| `BatchPrinter-Setup.exe` | 安装版（推荐），自动创建快捷方式，含卸载脚本 |
| `BatchPrinter-portable.zip` | 绿色便携版，解压即用 |

系统要求：Windows 10 / 11（64 位）

SHA256 校验值见 [SHA256SUMS.txt](SHA256SUMS.txt)。

## 🚀 快速上手

1. 运行 `BatchPrinter-Setup.exe`（或解压便携版后运行 `BatchPrinter.exe`）
2. 「添加文件 / 添加文件夹 / 添加压缩包」（或直接拖拽进窗口）
3. 右侧面板设置全局参数，选择打印机
4. 点击「立即打印全部」；双击列表中任意文件可为它单独设置参数

## 🖨 打印引擎说明

| 格式 | 引擎 | 份数/色彩/双面/页码范围 |
|---|---|---|
| PDF | 内置 PDFium 渲染 + GDI | ✅ 全部精确支持 |
| 图片 / 文本 | GDI 直印 | ✅ 全部精确支持 |
| Office (docx/xlsx/pptx/wps…) | 系统 Shell 打印动词（关联程序） | ⚠️ 经打印机默认值临时写入，尽力生效；页码范围不支持 |
| prn/ps/pcl | 字节流直送 | 按内容本身 |

> Office 格式如遇「ShellExecute 31 设备不起作用」错误，通常是默认打开方式为 UWP 应用（如 Edge），改为 WPS/Office 或使用扩展名映射即可。

## 🧰 从源码运行

```bash
pip install pywin32 pypdfium2 py7zr rarfile Pillow PySide6
python main.py
```

打包/构建安装器：

```bash
python -m PyInstaller --noconfirm --clean --windowed --onedir --name BatchPrinter --icon=app.ico main.py
python - <<'EOF'
# 生成 setup_payload/payload.zip 后:
python -m PyInstaller --noconfirm --clean --windowed --onefile --name BatchPrinter-Setup --icon=app.ico --add-data "setup_payload/payload.zip;." setup_main.py
EOF
```

## 📁 项目结构

```
main.py           GUI（PySide6）
printer_core.py   打印引擎（打印机枚举 / DEVMODE / GDI / SumatraPDF / Shell / RAW）
archive_util.py   压缩包自动解压（zip/7z/rar/tar）
setup_main.py     安装器脚本
gen_installer.py  安装包构建辅助
test_core.py      核心逻辑自测（26 项）
```

## ⚠️ 免责声明

本软件按“现状”提供，打印结果受打印机驱动影响，重要文档请先小批量试打。使用本软件产生的任何后果由使用者自行承担。

## 📄 License

[MIT](LICENSE)
