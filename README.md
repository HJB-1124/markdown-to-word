# Markdown 转 Word 转换器 (markdown-to-word)

这是一个专为大语言模型（Agent）设计的**输出适配器**与**文档渲染编译器**。它的核心作用是将 Agent 生成的 Markdown（包含复杂的 LaTeX 数学公式）稳定、可靠地转换为结构化且公式原生的 Microsoft Word (`.docx`) 文件。

## 一、核心定位：Agent 的“最后一公里”

**让 Agent 只负责内容和语义，让 Pandoc 负责 Word 文件结构。**

直接让大语言模型使用 Python 代码（如 `python-docx`）操作 Word 的底层 XML 结构不仅容易出错，且难以维护。本工具将“写作问题”与“文档工程问题”彻底解耦：Agent 只需要输出它最擅长的 Markdown 文本，本工具负责将其统一编译为人类工作环境中常要求的 `.docx` 交付物。

## 二、核心特性与实际价值

### 1. 原生 Word 数学公式 (OMML) 支持（最大亮点）
很多普通的 Word 生成方案（甚至简单的 `python-docx` 封装）无法将 LaTeX 公式（如 `\frac{-b\pm\sqrt{b^2-4ac}}{2a}`）转化为真正的 Word 公式，通常只能呈现为纯文本或插入为静态截图。
本工具底层调用 Pandoc，能将 Markdown 中的 LaTeX 数学表达式完美转换为 Word 原生的 **OMML (Office Math Markup Language)**。生成的公式是**可点击、可编辑的**。这对生成数学作业、工程报告、物理论文、算法说明尤为重要。

### 2. 结构化渲染
支持将 Markdown 中的标准结构高质量映射为 Word 原生结构，包括：
- 标题层级
- 字体加粗、斜体
- 有序/无序列表
- 引用块
- 超链接、代码块
- 表格

### 3. 中文排版优化选项
提供了首行缩进选项，一键自动为正文段落附加缩进格式。

## 三、快速开始

本工具通过 `scripts/convert.py` 提供 CLI 接口。在首次运行时，它会自动下载并配置独立的 Pandoc 环境，无需用户事先手动安装系统级依赖。

### 1. 转换直接传入的文本
```bash
uv run scripts/convert.py convert --text "LQR 的性能指标定义为：\n\n\$\$J=\int_0^\infty \left(x^\mathrm{T}Qx+u^\mathrm{T}Ru\right)\,\mathrm{d}t\$\$" --output "lqr_report.docx" --first-line-indent
```

### 2. 转换现有的 Markdown 文件
```bash
uv run scripts/convert.py convert --input-file "notes.md" --output "notes.docx" --first-line-indent
```

**命令行参数说明:**
- `--text`: 直接传入包含 Markdown/LaTeX 的字符串。
- `--input-file`: 传入现有的 `.md` 文件路径（与 `--text` 互斥）。
- `--output`: **(必需)** 指定输出 `.docx` 文件的路径。
- `--first-line-indent`: (可选) 为生成的正文段落添加约 2 字符的首行缩进，更符合中文排版习惯。

## 四、适用场景

- **含公式的学术/技术文档**：将 Agent 生成的推导过程、物理模型等一键编译为公式可编辑的 Word 成果。
- **自动化文档流水线**：将大量 Markdown 笔记、报告批量自动化转换为 Word 文件。

## 五、当前版本的局限性与边界

作为一个轻量级的公式友好型转换器，本工具目前主要解决“内容生成与公式排版”的基础需求。但在复杂的生产级或企业级排版上仍存在以下局限：

1. **环境配置固化且跨平台受限**：目前的自动化安装脚本硬编码了下载 `3.1.13` 版本的 Windows x86-64 Pandoc。这虽然保证了输出的绝对一致性，但既不支持 macOS/Linux 系统的免配置运行，也让用户无法使用系统已有的 Pandoc 或自动获取新版本更新。
2. **未开放高级排版能力**：目前的封装极为精简，没有透传 Pandoc 强大的 `--reference-doc` 接口。用户无法应用自定义 Word 模板、企业规范排版、自定义字体、页眉页脚及目录生成等进阶功能，只能输出默认朴素样式。
3. **首行缩进实现较为粗糙**：目前的缩进逻辑是通过 `python-docx` 后期遍历部分默认样式（如 `Normal`）附加缩进值实现的。这种方式无法覆盖表格内文本或特殊自定义样式，特定情况下也可能发生误判。
