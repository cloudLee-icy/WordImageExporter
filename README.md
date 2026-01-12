# WordImageExporter / Word 图片导出器

将 `.docx` 中插入的图片按文中顺序导出为 PNG（1.png, 2.png...），并可等比缩放到指定宽度（默认 500px）。

Export inline images from a `.docx` file in document order as PNG (1.png, 2.png...), with optional resize to target width (default 500px).

---

## ✅ 下载即用（推荐） / Ready-to-use (Recommended)

请到本仓库的 **Releases** 页面下载可执行文件（Windows）：

- `WordImageExporter.exe`：CLI 单文件版（拖拽 docx 到 exe 上即可导出）
- `WordImageExporterGUI.exe`：GUI 单文件版（双击打开选择文件）
- `WordImageExporterGUI-onedir.zip`：GUI onedir 版本（更稳，解压后运行里面的 exe）

Go to **Releases** to download prebuilt executables for Windows:
- `WordImageExporter.exe` (CLI onefile)
- `WordImageExporterGUI.exe` (GUI onefile)
- `WordImageExporterGUI-onedir.zip` (GUI onedir, recommended for maximum stability)

---

## 功能 / Features
- 按文中顺序导出（1.png, 2.png...）
- 输出 PNG
- 默认输出目录：原文档同目录下 `exported_images`
- 可选缩放到指定宽度（默认 500px，等比缩放，默认只缩小不放大）

- Export in appearance order (1.png, 2.png...)
- Output as PNG
- Default output folder: `exported_images` next to the input docx
- Optional resize to target width (default 500px; downscale only by default)

---

## 源码运行 / Run from source

```bash
pip install -r requirements.txt
python WordImageExporterCLI.py your.docx
python WordImageExporterGUI.py