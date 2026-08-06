# Markdown 转 Word 转换器 (markdown-to-word)

这是一个专为大语言模型（Agent）设计的**输出适配器**与**文档渲染编译器**。它的核心作用是将 Agent 生成的 Markdown（包含复杂的 LaTeX 数学公式）稳定、可靠地转换为结构化且公式原生的 Microsoft Word (`.docx`) 文件。

## 一、核心定位：Agent 的“最后一公里”

**让 Agent 只负责内容和语义，让 Pandoc 负责 Word 文件结构。**

直接让大语言模型使用 Python 代码（如 `python-docx`）操作 Word 的底层 XML 结构不仅容易出错，且难以维护。本工具将“写作问题”与“文档工程问题”彻底解耦：Agent 只需要输出它最擅长的 Markdown 文本，本工具负责将其统一编译为人类工作环境中常要求的 `.docx` 交付物。

## 二、核心特性与实际价值

### 1. 原生 Word 数学公式 (OMML) 支持（最大亮点）
很多普通的 Word 生成方案无法将 LaTeX 公式转化为真正的 Word 公式，通常只能呈现为纯文本或插入为静态截图。
本工具底层调用 Pandoc，能将 Markdown 中的 LaTeX 数学表达式完美转换为 Word 原生的 **OMML (Office Math Markup Language)**。生成的公式是**可点击、可编辑的**。这对生成数学作业、工程报告、物理论文尤为重要。

### 2. 高度灵活的排版控制 (`--reference-doc`)
工具现在原生支持传入您配置好的 Word 模板文件 (`.docx`)。只需一个模板，即可彻底接管中英文正文字体、多级标题字号、行距、页面布局、表格样式以及首行缩进等复杂排版规范。

### 3. Agent 工作流深度集成
- **管道流支持 (`--stdin`)**：允许大段内容直接通过系统标准输入送入转换器，避免大量 LaTeX 语法被 Shell 转义破坏或触碰命令行参数长度上限。
- **本地图片支持 (`--resource-path`)**：可显式指定资源路径，确保图片被正确打包入 Word。
- **安全的执行环境**：内置基于 SHA-256 哈希完整性校验的 Pandoc 自动化下载管理，并优先复用系统自带的 Pandoc 环境。

## 三、快速开始

本工具通过 `scripts/convert.py` 提供 CLI 接口。在首次运行时，它会自动下载并配置独立的 Pandoc 环境，无需事先手动安装系统级依赖。

### 1. 转换直接传入的文本
```bash
uv run scripts/convert.py convert --text "LQR 性能指标定义：\$\$\int_0^\infty (x^T Q x + u^T R u) dt\$\$" --output "lqr.docx"
```

### 2. 通过标准输入转换（推荐 Agent 使用）
```bash
echo "Hello **World**!" | uv run scripts/convert.py convert --stdin --output "hello.docx"
```

### 3. 应用 Word 模板
```bash
uv run scripts/convert.py convert --input-file "notes.md" --output "notes.docx" --reference-doc "my_template.docx"
```

**命令行参数说明:**
- `--text`: 直接传入包含 Markdown/LaTeX 的字符串。
- `--input-file`: 传入现有的 `.md` 文件路径。
- `--stdin`: 从标准输入读取纯文本或 Markdown（与以上两者互斥）。
- `--output`: **(必需)** 指定输出 `.docx` 文件的路径。
- `--reference-doc`: (可选) 传入自定义的 Word 模板文件（`.docx`），以控制生成文档的所有排版样式。
- `--first-line-indent`: (可选) 后处理模式，为正文和表格内的正文段落强制添加首行缩进。**注意：如果您使用了 `--reference-doc` 模板，该参数将被自动忽略，全权由模板控制缩进。**
- `--resource-path`: (可选) 指定 Markdown 中引用的本地图片资源所在目录。
- `--pandoc-path`: (可选) 指定自定义的 Pandoc 可执行文件路径。

## 四、适用场景

- **含公式的学术/技术文档**：将 Agent 生成的推导过程、物理模型等一键编译为公式可编辑的 Word 成果。
- **自动化文档流水线**：结合模板文件，将大量 Markdown 笔记自动化转换为统一排版规范的 Word 文件。

## 五、当前版本的局限性与边界

作为一个轻量级的公式友好型转换器，本工具目前主要解决“内容生成与公式排版”的基础需求。在某些特定场景下仍有以下局限：

1. **自动化下载仅限 Windows x86-64**：目前的**免配置自动化安装脚本**仍硬编码下载了 Windows x86-64 版本的 Pandoc。不过，macOS 和 Linux 用户现在可以通过手动将 Pandoc 添加入系统 `PATH` 或使用 `--pandoc-path` 参数来完美绕过此限制。
2. **后处理缩进逻辑的妥协性**：当您未使用 `--reference-doc` 模板，却强行启用 `--first-line-indent` 时，脚本会采用后处理方式为 `Normal`、`正文` 等常规样式追加 21 磅的硬编码缩进。这种方式虽然现在能穿透表格，但仍属于经验值补丁，无法完美适配所有非常规字号。**强烈建议通过提供 Word 模板来解决严格的排版问题。**
