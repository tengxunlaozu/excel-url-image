<div align="center">

# Excel URL Image Processor / Excel 图片处理器

<p align="center">
  <a href="#en">🇬🇧 English</a> · <a href="#zh">🇨🇳 中文</a>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.13+-blue?logo=python" alt="Python">
  <img src="https://img.shields.io/badge/PySide6-6.11-green?logo=qt" alt="PySide6">
  <img src="https://img.shields.io/badge/license-MIT-orange" alt="License">
  <img src="https://img.shields.io/github/v/release/tengxunlaozu/excel-url-image" alt="Release">
</p>

<p align="center">
  🖼️ Batch convert image URLs in Excel cells into embedded images — with a beautiful PySide6 GUI.<br>
  🖼️ 将 Excel 单元格中的图片链接批量下载并嵌入为图片——附带精美的 PySide6 图形界面。
</p>

<p align="center">
  <a href="https://tengxunlaozu.github.io/excel-url-image/">🌐 Interactive Docs / 交互式文档</a>
</p>

</div>

---

<a name="en"></a>

## 🇬🇧 English

### ✨ Features

- **GUI Operation** — Built with PySide6, no command-line skills needed
- **Batch Processing** — Process all image URLs in a column at once
- **Multi-Image Support** — One cell can contain multiple URLs, automatically arranged side by side
- **Auto Column Width** — Column width adjusts automatically based on image count
- **Original Filenames** — Images saved to `excel_images/` with original filenames
- **Progress & Logs** — Real-time progress bar and detailed processing log

### 📦 Download

Download the latest packaged **`Excel图片处理器.exe`** from the [Releases page](https://github.com/tengxunlaozu/excel-url-image/releases) — no Python installation required.

### 🚀 Quick Start

#### Option 1: Run the Pre-built EXE

1. Download `Excel图片处理器.exe` from [Releases](https://github.com/tengxunlaozu/excel-url-image/releases)
2. Double-click to run
3. Select your Excel file, enter the column letter (e.g., `E`), and click **Start**

#### Option 2: Run from Source

```bash
# Clone the repo
git clone https://github.com/tengxunlaozu/excel-url-image.git
cd excel-url-image

# Install dependencies
pip install -r requirements.txt

# Run the GUI
python excel_image_processor_gui.py
```

### 🖥️ GUI Overview

| Section | Description |
|---------|-------------|
| **File Settings** | Browse for input Excel file; output path auto-generated |
| **Column Setting** | Enter the column letter containing image URLs (e.g., E) |
| **Progress Bar** | Shows current processing progress |
| **Log Panel** | Real-time logs with dark terminal-style display |

### 📁 Output Structure

```
your-project/
├── excel_images/            # Downloaded images (original filenames)
│   ├── admin-check_xxx.jpeg
│   └── ...
└── 行政检查_处理后_20260820_102916.xlsx  # Processed Excel file
```

### ⚙️ Requirements

- Python 3.13+
- PySide6
- openpyxl
- requests

### 📄 License

This project is licensed under the MIT License.

---

<a name="zh"></a>

## 🇨🇳 中文

### ✨ 功能特点

- **图形界面操作** — 基于 PySide6 构建，无需命令行知识
- **批量处理** — 一次性处理整列的图片链接
- **多图支持** — 一个单元格可含多个链接，自动并排排列
- **列宽自适应** — 根据图片数量自动调整列宽
- **保留原始文件名** — 图片下载到 `excel_images/` 目录，使用原始文件名
- **进度与日志** — 实时进度条 + 详细处理日志

### 📦 下载

从 [Releases 页面](https://github.com/tengxunlaozu/excel-url-image/releases) 下载最新打包好的 **`Excel图片处理器.exe`**，无需安装 Python 即可运行。

### 🚀 快速开始

#### 方式一：直接运行 EXE

1. 从 [Releases](https://github.com/tengxunlaozu/excel-url-image/releases) 下载 `Excel图片处理器.exe`
2. 双击运行
3. 选择 Excel 文件，输入列字母（如 `E`），点击「开始处理」

#### 方式二：源码运行

```bash
# 克隆仓库
git clone https://github.com/tengxunlaozu/excel-url-image.git
cd excel-url-image

# 安装依赖
pip install -r requirements.txt

# 运行 GUI
python excel_image_processor_gui.py
```

### 🖥️ 界面说明

| 功能区域 | 说明 |
|---------|------|
| **文件设置** | 浏览选择输入 Excel 文件；输出路径自动生成 |
| **列设置** | 输入图片链接所在的列字母（如 E） |
| **进度条** | 实时显示当前处理进度 |
| **日志面板** | 深色终端风格，实时输出详细处理日志 |

### 📁 输出结构

```
your-project/
├── excel_images/            # 下载的图片（原始文件名）
│   ├── admin-check_xxx.jpeg
│   └── ...
└── 行政检查_处理后_20260820_102916.xlsx  # 处理后的 Excel 文件
```

### ⚙️ 依赖要求

- Python 3.13+
- PySide6
- openpyxl
- requests

### 📄 许可证

本项目基于 MIT 许可证开源。

---

<div align="center">
  Made with ❤️ by <a href="https://github.com/tengxunlaozu">tengxunlaozu</a>
</div>