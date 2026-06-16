# Markdown to Word Converter

## Overview
This tool converts raw text or Markdown (including complex LaTeX mathematical formulas) into a native Microsoft Word (`.docx`) file. It uses `Pandoc` under the hood to ensure LaTeX equations are properly converted into Word's native Office Math Markup Language (OMML), avoiding the unformatted text output commonly seen with simple Python libraries. The tool automatically downloads and installs a persistent copy of Pandoc upon first use.

## Quick Start
To use the tool, run the helper script `convert.py`.

```bash
uv run scripts/convert.py convert --text "This is an equation: \$\$E=mc^2\$\$" --output "output.docx"
```

## Utility Scripts
**`scripts/convert.py`**
A CLI script that handles downloading Pandoc (if missing), writing temporary Markdown files, and executing the conversion process.

**Subcommands:**
- `convert`: Converts Markdown/Text to Word (.docx).

**Arguments for `convert`:**
- `--text`: The raw text or markdown containing LaTeX formulas to be converted.
- `--input-file`: The path to an existing `.md` file to convert (mutually exclusive with `--text`).
- `--output`: (Required) Destination path for the `.docx` file.

**Examples:**
Using direct text input:
```bash
uv run scripts/convert.py convert --text "Hello **World**!" --output "hello.docx"
```

Using an existing file:
```bash
uv run scripts/convert.py convert --input-file "notes.md" --output "notes.docx"
```
