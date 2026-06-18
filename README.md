# Markdown 转 Word 转换器

## 简介
本工具将纯文本或 Markdown（包含复杂的 LaTeX 数学公式）转换为原生的 Microsoft Word (`.docx`) 文件。它底层使用 `Pandoc`，确保 LaTeX 公式能够被正确地转换为 Word 的原生 Office Math Markup Language (OMML)，从而避免了常见的简单 Python 库所生成的无格式文本输出。本工具在首次使用时会自动下载并安装一个持久化版本的 Pandoc。

## 快速开始
如需使用本工具，请运行辅助脚本 `convert.py`。

```bash
uv run scripts/convert.py convert --text "这是一个公式: \$\$E=mc^2\$\$" --output "output.docx"
```

## 实用脚本
**`scripts/convert.py`**
这是一个命令行界面 (CLI) 脚本，负责下载 Pandoc（如果缺失）、写入临时的 Markdown 文件，并执行转换过程。

**子命令:**
- `convert`: 将 Markdown/文本 转换为 Word (.docx) 文件。

**`convert` 参数:**
- `--text`: 包含需要转换的 LaTeX 公式的纯文本或 Markdown。
- `--input-file`: 要转换的现有 `.md` 文件的路径（与 `--text` 互斥）。
- `--output`: (必需) 输出 `.docx` 文件的目标路径。

**示例:**
使用直接文本输入:
```bash
uv run scripts/convert.py convert --text "你好 **世界**!" --output "hello.docx"
```

使用现有文件:
```bash
uv run scripts/convert.py convert --input-file "notes.md" --output "notes.docx"
```
