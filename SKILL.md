---
name: markdown-to-word
description: >-
  Converts a text string, Markdown file, or standard input containing text and LaTeX mathematical
  formulas into a structured Microsoft Word (.docx) file with native math support using Pandoc.
  Supports Word templates, custom Pandoc paths, and automatic secure installation.
---

# Markdown to Word Converter

## Overview
This skill converts Markdown (including Pandoc-supported LaTeX mathematical formulas) into a native Microsoft Word (`.docx`) file. It uses `Pandoc` under the hood to ensure LaTeX equations are properly converted into Word's native Office Math Markup Language (OMML), making them editable inside Word. It decouples content generation from document formatting, allowing you to focus on Markdown semantics while producing standard Word files.

If Pandoc is not provided or found in the system `PATH`, the script safely downloads and caches a verified version (using SHA-256 integrity checks) to `~/.gemini/tools/pandoc/`.

## Dependencies
- `uv` must be installed and available in the system path.
- Python dependencies (such as `python-docx`) are managed automatically by `uv`.
- Pandoc is discovered from a custom path, the system `PATH`, or the local cache; on Windows x86-64 it can be downloaded automatically when missing.

## Agent Behavior
When invoking this skill, agents should adhere to the following rules:
1. **Prefer `--stdin`** for generated or long Markdown content (avoids shell escaping/length issues).
2. **Prefer `--input-file`** when the Markdown already exists on disk.
3. **Use `--reference-doc`** when the user requests specific Word formatting or a custom template.
4. **Do not combine** `--reference-doc` with `--first-line-indent` (the script ignores it anyway).
5. **Supply `--resource-path`** when the Markdown contains local images.
6. **Use `--text`** only for short, simple content.
7. After conversion, verify that the output file exists and report its path to the user.

## Quick Start
To use the skill, run the helper script `convert.py`.

For Agents, using standard input (`--stdin`) is **highly recommended**. Here is a PowerShell example:

```powershell
# Using a Here-String in PowerShell
@"
# Title

`$`$E=mc^2`$`$
"@ | uv run ~/.gemini/config/skills/markdown-to-word/scripts/convert.py convert --stdin --output "output.docx"

# Or simply piping an existing file:
Get-Content ".\report.md" -Raw | uv run ~/.gemini/config/skills/markdown-to-word/scripts/convert.py convert --stdin --output "output.docx"
```

## Utility Scripts
**`scripts/convert.py`**
A CLI script that handles executing the conversion process, discovering or downloading Pandoc, and managing temporary files.

**Subcommands:**
- `convert`: Converts Markdown/Text to Word (.docx).

**Arguments for `convert`:**
- `--stdin`: (Recommended) Read raw text or markdown from standard input.
- `--input-file`: The path to an existing `.md` file to convert.
- `--text`: The raw text or markdown to be converted via command line argument.
  *(Note: You must provide exactly ONE of `--stdin`, `--input-file`, or `--text`)*
- `--output`: (Required) Destination path for the `.docx` file.
- `--reference-doc`: (Optional) Path to a Word template (`.docx`) to control the main output styles and document layout.
- `--resource-path`: (Optional) Directory relative to which Markdown image paths should be resolved.
- `--first-line-indent`: (Optional) A heuristic post-processing patch that applies an approx. 2-character indent to standard paragraphs and table cells. **Note: Automatically disabled if `--reference-doc` is used.**
- `--pandoc-path`: (Optional) Path to a custom pandoc executable. If valid, this takes priority over the system `PATH` and automated downloads.

## Examples

**Using a template and an input file:**
```bash
uv run ~/.gemini/config/skills/markdown-to-word/scripts/convert.py convert --input-file "report.md" --output "final.docx" --reference-doc "corporate_template.docx" --resource-path "."
```

**Using direct text input with fallback indentation:**
```bash
uv run ~/.gemini/config/skills/markdown-to-word/scripts/convert.py convert --text "Hello **World**!" --output "hello.docx" --first-line-indent
```

## Pandoc Resolution Order
The script locates Pandoc in the following order:
1. Valid custom path provided via `--pandoc-path`.
2. Existing installation in the system `PATH`.
3. Previously cached automated download.
4. If none are found, securely downloads the Windows x86-64 binary, verifies its SHA-256 hash, and installs it.

## Common Mistakes
- **Omitting Required Output**: The `--output` argument is strictly required.
- **Providing Multiple Inputs**: Supplying more than one of `--text`, `--input-file`, or `--stdin` will result in an error. Choose exactly one.
- **Malformed LaTeX**: If the provided LaTeX is severely malformed, Pandoc will fail to compile it and return a non-zero exit code. The script will surface stderr to help identify the parse error.
- **Overusing `--text`**: Do not use `--text` for large reports. Use `--stdin` or write the text to a temporary `.md` file and use `--input-file` instead.
- **Incorrect `--resource-path`**: Ensure the resource path aligns with how images are referenced in the Markdown. (e.g., if markdown has `![img](images/chart.png)`, the resource path should usually be `.`, not `./images`).
