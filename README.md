[README.md](https://github.com/user-attachments/files/24530839/README.md)


# 📝 PaperSmith (Markdown 智能编辑器)

![Python](https://img.shields.io/badge/Python-3.8%2B-blue) ![PyQt6](https://img.shields.io/badge/GUI-PyQt6-green) ![License](https://img.shields.io/badge/License-MIT-orange) ![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

**PaperSmith** 是一款专为中文写作和文档交付打造的 Markdown 编辑器。它解决了传统编辑器在导出 Word 时格式混乱的痛点，并提供了智能的格式规范功能，让您专注于写作，无需操心排版。

## ✨ 核心特性

### 1. 🚀 智能格式规范 (Smart Formatting)
告别繁琐的手动调整，软件会实时监控并修正您的文档结构：

*   **自动列表隔离**：当列表与普通文本紧挨时，自动插入空行，防止渲染错误。
*   **自动表格隔离**：智能检测表格边界，确保表格前后有正确的间距。
*   **嵌套保护**：智能识别嵌套列表，不会破坏原有的缩进结构。

### 2. 📄 完美的 Word 导出
基于 `Pandoc` 与 `python-docx` 的深度后处理技术：

*   **格式清洗**：自动去除导出 Word 时产生的多余空行。
*   **列表优化**：修复 Word 中列表缩进过大、符号错乱的问题。
*   **样式重置**：自动将任务列表转换为美观的复选框样式，而非乱码符号。

### 3. 👁️ 丝滑的实时预览

*   **无闪烁更新**：采用增量 DOM 更新技术，输入时预览区不会闪烁或跳动。
*   **双重同步**：
    *   **滚动条跟随**：预览区精准跟随编辑区滚动。
    *   **状态记忆**：记住您的同步开关设置。

### 4. 🧮 强大的数学公式

*   内置 **KaTeX** 渲染引擎（使用国内极速 CDN）。
*   完美支持行内公式 `$E=mc^2$` 和块级公式 `$$...$$`。
*   解决了 Markdown 解析器对 LaTeX 下划线 `_` 的误判问题。

## 🛠️ 安装与运行

### 方式一：直接运行 (源码)

1.  **克隆仓库**
    ```
    git clone https://github.com/Ma6302/papersmith.git
    cd papersmith
    ```

2.  **安装依赖**
    ```
    pip install PyQt6 markdown pypandoc python-docx
    ```
    *注意：您还需要在系统路径中安装 `pandoc`，或者让软件首次运行自动安装。*

3.  **运行**
    ```
    python main.py
    ```

### 方式二：下载发行版 (Windows)
前往 [Releases](https://github.com/Ma6302/papersmith-markdown-editor/releases/tag/v16.8) 页面下载最新的 `.exe` 安装包，开箱即用（已内置 Pandoc 安装程序）。

## 📖 使用指南

*   **同步滚动**：点击工具栏的 `🔗 同步` 按钮可开启/关闭预览区跟随。
*   **插入元素**：工具栏提供了 `H` (标题), `B` (粗体), `I` (斜体), `▦` (表格), `🖼` (图片) 等快捷按钮。
*   **导出文档**：点击右上角的 `💾 保存` 按钮，选择 `.docx` 或 `.pdf` 格式即可导出。

## 🔧 技术栈

*   **GUI**: PyQt6 + QWebEngine
*   **Parser**: Python-Markdown (with Custom Extensions)
*   **Converter**: Pandoc + python-docx
*   **Math**: KaTeX (via Staticfile CDN)


## 📄 开源协议

本项目基于 [MIT 协议](LICENSE) 开源。

---
*Created with ❤️ by [Ma6302]*
