# Markdown 转 Word 转换器 (markdown-to-word)

这是一个专为大语言模型（Agent）设计的**输出适配器**与**文档渲染编译器**。它的核心作用是将 Agent 生成的 Markdown（包含复杂的 LaTeX 数学公式）稳定、可靠地转换为结构化且公式原生的 Microsoft Word (`.docx`) 文件。

## 一、核心定位：Agent 的“最后一公里”

**让 Agent 只负责内容和语义，让 Pandoc 负责 Word 文件结构。**

直接让大语言模型使用 Python 代码（如 `python-docx`）操作 Word 的底层 XML 结构不仅容易出错，且难以维护。本工具将“写作问题”与“文档工程问题”彻底解耦：Agent 只需要输出它最擅长的 Markdown 文本，本工具负责将其统一编译为人类工作环境中常要求的 `.docx` 交付物。

## 二、核心特性与实际价值

### 1. 原生 Word 数学公式 (OMML) 支持（最大亮点）
很多普通的 Word 生成方案无法将 LaTeX 公式转化为真正的 Word 公式，通常只能呈现为纯文本或插入为静态截图。
本工具底层调用 Pandoc，能将 Markdown 中的 LaTeX 数学表达式完美转换为 Word 原生的 **OMML (Office Math Markup Language)**。生成的公式是**可点击、可编辑的**。这对生成数学作业、工程报告、物理论文尤为重要。

### 2. 默认 Word 排版标准

未指定 `--reference-doc` 时，转换器自动应用以下格式，无需额外参数：

| 项目 | 默认标准 | 适用范围与说明 |
| --- | --- | --- |
| 纸张与方向 | **A4 纵向（210 × 297 毫米）** | 宽表格也保持纵向；需要时调整列宽和换行。 |
| 中文字体 | **宋体** | 公式外的中文文字。 |
| 英文和数字字体 | **Times New Roman** | 公式外的英文字母和数字。 |
| 数学公式字体 | **Cambria Math** | 行内公式与独立公式均使用原生、可编辑的 Word 公式。 |
| 正文字号 | **五号（10.5 磅）** | 常规正文段落与表格正文。 |
| 正文首行缩进 | **2 个字符** | 表格外的正文，包括标题后的首段和含行内公式的正文段落。 |
| 文字颜色 | **黑色（`#000000`）** | 全文文字，包括标题、超链接、代码高亮和公式。 |
| 表格内容对齐 | **水平居中、垂直居中** | 表头、正文和嵌套表格均适用；单元格内段落不缩进，独立公式也居中。 |

正文缩进使用 Word 原生的“首行缩进 2 字符”属性，并提供 21 磅的兼容值，不在文本前插入空格。标题、独立公式、列表和代码块保留各自的对齐方式；标题保留 Pandoc 默认的层级字号和加粗等强调格式。

### 3. 自定义 Word 模板 (`--reference-doc`)

工具支持传入您配置好的 Word 模板文件 (`.docx`)。**模板优先于内置默认格式**：使用 `--reference-doc` 时，转换器跳过上述默认排版处理，由 Pandoc 按模板样式和布局生成文档。

`--first-line-indent` 仅为兼容旧命令而保留。默认转换已经自动缩进正文，不需要该参数；使用模板时，该参数会被忽略。

### 4. Agent 工作流深度集成
- **管道流支持 (`--stdin`)**：允许大段内容直接通过系统标准输入送入转换器，避免大量 LaTeX 语法被 Shell 转义破坏或触碰命令行参数长度上限。
- **本地图片支持 (`--resource-path`)**：可显式指定资源路径，确保图片被正确打包入 Word。
- **安全的执行环境**：内置基于 SHA-256 哈希完整性校验的 Pandoc 自动化下载管理，并优先复用系统自带的 Pandoc 环境。

## 三、快速开始

本工具通过 `scripts/convert.py` 提供 CLI 接口。先安装 `uv` 并确保它可在命令行中使用，Python 依赖由 `uv run` 自动管理。

Pandoc 按以下顺序查找：`--pandoc-path` 指定的路径、系统 `PATH`、本地缓存。均未找到时，脚本会下载 Windows x86-64 版本并校验 SHA-256；macOS 和 Linux 请先安装 Pandoc 或使用 `--pandoc-path`。

以下命令在 skill 目录运行，示例使用 PowerShell。LaTeX 行内公式使用 `$...$`，独立公式使用 `$$...$$`。

### 1. 转换直接传入的文本
```powershell
uv run scripts/convert.py convert --text '质能方程：$E=mc^2$。' --output "formula.docx"
```

### 2. 通过标准输入转换（推荐 Agent 使用）
```powershell
@'
# LQR 学习笔记

性能指标定义如下：

$$
J = \int_0^\infty (x^T Q x + u^T R u)\,dt
$$
'@ | uv run scripts/convert.py convert --stdin --output "lqr.docx"
```

### 3. 转换已有 Markdown 文件

```powershell
uv run scripts/convert.py convert --input-file "notes.md" --output "notes.docx"
```

此命令自动应用上表中的默认排版。Markdown 中的本地图片默认相对于输入文件所在目录解析。

### 4. 应用 Word 模板

```powershell
uv run scripts/convert.py convert --input-file "notes.md" --output "notes.docx" --reference-doc "my_template.docx"
```

**命令行参数说明:**

- `--text`: 直接传入包含 Markdown/LaTeX 的字符串。
- `--input-file`: 传入现有的 `.md` 文件路径。
- `--stdin`: 从标准输入读取纯文本或 Markdown（与以上两者互斥）。
- `--output`: **(必需)** 指定输出 `.docx` 文件的路径。
- `--reference-doc`: (可选) 传入自定义的 Word 模板文件（`.docx`），跳过内置默认排版处理，由 Pandoc 应用模板的样式和布局。
- `--first-line-indent`: (兼容旧命令) 正文默认已首行缩进 2 个字符，无需额外传入该参数；表格内容仍保持居中且不缩进。使用 `--reference-doc` 时忽略该参数，按模板生成文档。
- `--resource-path`: (可选) 指定 Markdown 中引用的本地图片资源所在目录。
- `--pandoc-path`: (可选) 指定自定义的 Pandoc 可执行文件路径。

`--text`、`--input-file` 和 `--stdin` 三种输入方式必须且只能选择一种。

## 四、适用场景

- **含公式的学术/技术文档**：将 Agent 生成的推导过程、物理模型等一键编译为公式可编辑的 Word 成果。
- **自动化文档流水线**：结合模板文件，将大量 Markdown 笔记自动化转换为统一排版规范的 Word 文件。

## 五、当前版本的局限性与边界

作为一个轻量级的公式友好型转换器，本工具目前主要解决“内容生成与公式排版”的基础需求。在某些特定场景下仍有以下局限：

1. **自动化下载仅限 Windows x86-64**：macOS 和 Linux 用户需自行安装 Pandoc，并通过系统 `PATH` 或 `--pandoc-path` 提供可执行文件。
2. **字体显示依赖本机字体**：查看或渲染文档的设备应安装宋体、Times New Roman 和 Cambria Math；缺少字体时，应用程序可能使用替代字体。
3. **复杂页面布局仍需检查**：默认保持 A4 纵向；宽表格和较长公式可能需要调整列宽、换行或使用模板。转换器不会自动将页面改为横向。
