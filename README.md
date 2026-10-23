# Markdown 转 Word 转换器

将 Markdown 和 Pandoc 支持的 LaTeX 数学公式转换为 Microsoft Word（`.docx`）。公式使用 Word 原生 OMML 格式，可在 Word/WPS 中继续编辑。

默认使用浅灰表头和完整浅灰网格线；Markdown 章节分隔线默认不输出，可用 `--keep-horizontal-rules` 保留。自定义 Word 模板可覆盖默认排版。

仓库同时提供供 Agent 使用的 [SKILL.md](SKILL.md)、[转换脚本](scripts/convert.py)和[回归测试](scripts/test_convert.py)。

## 快速开始

需要安装 `uv`，Python 依赖由 `uv run` 自动管理。以下命令在仓库目录执行，示例使用 PowerShell。

```powershell
git clone https://github.com/HJB-1124/markdown-to-word.git
cd markdown-to-word
uv run scripts/convert.py convert --input-file "notes.md" --output "notes.docx"
```

输入文件使用 UTF-8 编码，支持 `.md` 和包含 Markdown 的 `.txt` 文件。脚本显式按 Markdown 解析输入，文件扩展名不会改变公式解析方式。本地图片默认相对于输入文件所在目录查找。

Pandoc 按以下顺序查找：

1. `--pandoc-path` 指定的可执行文件。
2. 本 skill 目录下 `.tools/pandoc/` 中的缓存版本。
3. 系统 `PATH` 中的 Pandoc。
4. 均未找到时，下载内置的 Windows x86-64 版本到本 skill 的 `.tools/pandoc/`，并校验 SHA-256。

依赖目录按转换脚本的真实位置解析，与当前工作目录无关；通过目录联接调用时也使用同一份缓存，不依赖 Antigravity 配置目录。

`.tools/` 是脚本管理的本地运行缓存，已由 `.gitignore` 排除。克隆仓库后可按需安装，无需将 Pandoc 可执行文件提交到仓库。

macOS 和 Linux 请预先安装 Pandoc，或通过 `--pandoc-path` 指定路径。

## 公式写法与编号

同时支持两组 LaTeX 分隔符：

| 类型 | 支持的写法 |
| --- | --- |
| 行内公式 | `\(...\)`、`$...$` |
| 独立公式 | `\[...\]`、`$$...$$` |

例如，以下 Markdown 会生成 1 个行内公式和 2 个独立公式：

```markdown
模型参数为 \(\theta\)。

\[
\hat{y}=f_\theta(x),\qquad x\in X \tag{1}
\]

$$
\begin{cases}
x'=x+\delta\\
y'=f_\theta(x')
\end{cases}
$$
```

`\tag{n}` 保留为带括号的编号，`\tag*{label}` 保留为不加括号的标签。标签作为可编辑的公式文字显示，不是 Word 自动编号或交叉引用字段。

代码块和行内代码保持字面内容。普通文字或加粗文字中的表达式不会自动变成公式；需要原生 Word 公式时，应先在源文中添加上述分隔符。LaTeX 公式环境也应放在数学分隔符内部。

## 转换校验与失败处理

转换器先解析 Pandoc JSON 文档，再生成并检查 DOCX：

- 显式启用反斜杠与美元符号公式分隔符。
- 保留公式标签，并检查正文、脚注和尾注中的原生公式数量、独立公式数量、非空公式对象及简单显式标签。
- 显示 Pandoc 的诊断信息；遇到转换警告、不支持的公式或会在 Word 中丢失的原始 TeX 时，以非零退出码结束。
- 先生成临时文件，通过格式处理、ZIP 完整性检查和公式校验后，才替换目标文件。失败时保留已有目标文件。

成功时会输出校验摘要，例如：

```text
Validated 3 native Word equations (2 display, 1 inline); preserved 1 equation labels.
Success! Document successfully written to: ...
```

公式数量来自已解析的正文及实际输出的标题、作者、摘要等元数据。结构校验用于发现公式丢失、回退为文本等转换问题；原文的数学正确性和最终页面布局仍需按文档用途检查。

## 清理粘贴文本中的来源文件名

从聊天或论文译文复制的内容中，可能残留如下来源标记：

```text
(2024TPAMI)Paper_Name…
(2026IF)Multi-modal Attack Surv…
```

脚本默认检测并提示匹配到的标记。需要生成清理后的阅读文档时，使用：

```powershell
uv run scripts/convert.py convert --input-file "translation.txt" --output "translation.docx" --strip-source-markers
```

`--strip-source-markers` 仅删除匹配规则的来源文件名标记，并报告数量；代码示例、公式和 `[1]` 等文献引用编号保持原样。

## 默认 Word 排版

未指定 `--reference-doc` 时自动应用以下格式：

| 项目 | 默认格式 |
| --- | --- |
| 页面 | A4 纵向，210 × 297 毫米 |
| 中文正文 | 宋体，五号（10.5 磅） |
| 英文与数字 | Times New Roman，五号（10.5 磅） |
| 原生公式 | Cambria Math，可编辑 |
| 正文样式 | 内置 Normal，在中文 Word/WPS 中显示为“正文” |
| 首行缩进 | 在 Normal 样式中定义为 2 个字符，并提供 21 磅兼容值 |
| 列表 | 保留编号、项目符号及原有列表缩进；紧凑列表使用 Compact |
| 标题 | 保留 Pandoc 的标题层级、字号和强调 |
| 文字颜色 | 全文黑色，包括超链接、代码高亮和公式 |
| 表格 | 所有单元格内容水平、垂直居中；取消单元格内段落缩进 |
| 表头底色 | 浅灰色 `#F2F2F2`，正文单元格为白色 |
| 表格边框 | 完整浅灰网格线 `#D9D9D9`，0.5 磅，包含外框与内部横竖线 |

标题、独立公式、列表和代码块保留各自的对齐方式。正文通过样式继承首行缩进，重新应用 Normal 或新建正文段落时仍可保留。查看文档的设备需要安装相应字体；宽表格与长公式可能需要调整列宽或换行。

### 章节分隔线

默认忽略 Markdown 中独立的 `---` 等章节分隔线。过滤在 Pandoc 解析后执行，只移除分隔线节点；代码中的字面内容、YAML 元数据、Setext 标题下划线和表格分隔行保持原有含义。表格边框与公式框不受影响。

需要保留章节分隔线时，显式使用：

```powershell
uv run scripts/convert.py convert --input-file "notes.md" --output "notes.docx" --keep-horizontal-rules
```

这项内容过滤同样适用于自定义模板；模板仍优先决定表格外观。

### 使用 Word 模板

```powershell
uv run scripts/convert.py convert --input-file "notes.md" --output "notes.docx" --reference-doc "my_template.docx"
```

提供模板时，脚本跳过内置排版，由 Pandoc 使用模板样式和布局；公式解析与完整性校验仍然执行。`--first-line-indent` 仅保留用于兼容旧命令，默认正文已缩进；使用模板时该参数被忽略。

## 输入方式与资源路径

### 短文本

```powershell
uv run scripts/convert.py convert --text '质能方程：$E=mc^2$。' --output "formula.docx"
```

### 标准输入

标准输入按 UTF-8 读取。通过 PowerShell 管道传递中文时，需使用 UTF-8 编码；以下示例在当前会话设置管道编码，并使用单引号 Here-String 保留公式中的美元符号和反斜杠：

```powershell
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
@'
# LQR 学习笔记

性能指标：

$$
J = \int_0^\infty (x^T Q x + u^T R u)\,dt
$$
'@ | uv run scripts/convert.py convert --stdin --output "lqr.docx"
```

### 本地图片

使用 `--input-file` 时，图片默认相对于源文件目录解析。使用 `--stdin`、`--text`，或需要从其他目录读取图片时，可指定：

```powershell
uv run scripts/convert.py convert --input-file "notes.md" --output "notes.docx" --resource-path ".\assets"
```

## 命令行参数

| 参数 | 说明 |
| --- | --- |
| `--input-file PATH` | 读取 UTF-8 Markdown 或文本文件 |
| `--stdin` | 从标准输入读取 UTF-8 Markdown |
| `--text TEXT` | 直接传入短文本或 Markdown |
| `--output PATH` | 必需，输出文件路径须以 `.docx` 结尾，父目录须已存在 |
| `--reference-doc PATH` | 使用 Word 模板，替代内置排版 |
| `--resource-path PATH` | 指定本地图片等资源所在目录 |
| `--pandoc-path PATH` | 指定 Pandoc 可执行文件 |
| `--keep-horizontal-rules` | 保留 Markdown 章节分隔线；默认忽略 |
| `--strip-source-markers` | 删除匹配的来源文件名残留；默认仅检测提示 |
| `--first-line-indent` | 兼容参数；默认正文已有两字符首行缩进 |

`--input-file`、`--stdin` 和 `--text` 必须且只能选择一种。

## 回归测试

修改转换器后运行：

```powershell
uv run scripts/test_convert.py
```

行为回归测试覆盖混合公式分隔符、希腊字母、上下标、分段公式、公式标签、转换失败时保护旧文件、引用残留清理、代码保护、UTF-8 中文输入、Word 模板、脚注、列表、浅灰表头与网格线、章节分隔线过滤及其保留开关、相对路径图片，以及 Pandoc 迁移和安装。测试产物写入临时目录；需要可用的 Pandoc，未找到时测试会跳过。

## 范围与限制

- LaTeX 支持范围取决于 Pandoc；出现不支持的命令时，应根据诊断修改公式。
- 输入中的加粗表达式、未明确标记的数学文字和不完整的分隔符需要先整理为有效 Markdown 数学语法。
- 公式编号以可编辑文字保留，不提供自动编号或交叉引用更新。
- 页面布局、字体替换和很长的表格或公式，需要在目标 Word/WPS 环境中检查。
