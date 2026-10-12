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

## Default Word Formatting
When `--reference-doc` is omitted, the script applies these defaults:
- Page: **A4 portrait (210 × 297 mm)**. Keep portrait orientation when adjusting wide tables; adjust column widths and wrapping instead.
- Chinese text: **宋体 (SimSun)**.
- English letters and digits outside formulas: **Times New Roman**.
- Native Word formulas: **Cambria Math**, explicitly set in the document's math settings and equation runs; formulas remain editable.
- Body text: **五号, 10.5 pt**, including standard paragraphs and table text.
- Body paragraph style: **Normal**, displayed as **正文** in Chinese Word/WPS. Normalize Pandoc's `FirstParagraph` and `BodyText` paragraphs to this built-in style while retaining their paragraph spacing.
- Body paragraphs outside tables: **two-character first-line indent**, defined in the **Normal style itself** using Word's native character units (`firstLineChars=200`, with a 21 pt compatibility fallback). Body paragraphs inherit this setting, including the first paragraph after headings and paragraphs containing inline formulas. Reapplying Normal or creating another body paragraph retains the indent. Titles, headings, standalone equations, lists and code blocks keep their own alignment; table cells remain centered with no indentation.
- Bullet and numbered lists: preserve their list styles, numbering and indentation. Pandoc's tight lists keep **Compact**, with the same body fonts, size and black color; do not normalize lists to Normal or apply the body first-line indent.
- All text: **black (`#000000`)**, including headings, hyperlinks, code highlighting and formulas.
- Table contents: **centered horizontally and vertically** in every cell, including header rows and nested tables. Clear paragraph indentation inside cells so content remains centered; center display equations inside cells too.
- Headings retain Pandoc's heading sizes and emphasis.

These defaults are applied automatically; no additional flags or template are needed. A supplied `--reference-doc` takes precedence: the script skips these built-in defaults and lets Pandoc apply the template's styles and layout.

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
5. **Supply `--resource-path`** only when using `--stdin` or `--text` with local images, or when image paths should be resolved relative to a directory other than the Markdown file's directory. (When using `--input-file`, the script automatically defaults to the input file's directory).
6. **Use `--text`** only for short, simple content.
7. After conversion, verify that the process exited successfully (code 0) and that the output file exists and is non-empty. Finally, report the absolute path to the user.

## Quick Start
To use the skill, run the helper script `convert.py`.

**For Agent-generated content:** Using standard input (`--stdin`) is **highly recommended**.
```powershell
# Using a Here-String in PowerShell
@"
# Title

`$`$E=mc^2`$`$
"@ | uv run "$HOME\.gemini\config\skills\markdown-to-word\scripts\convert.py" convert --stdin --output "output.docx"
```

**For existing files:** Use `--input-file` to automatically handle resources and avoid PowerShell pipe parsing.
```powershell
uv run "$HOME\.gemini\config\skills\markdown-to-word\scripts\convert.py" convert --input-file ".\report.md" --output ".\output.docx"
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
- `--reference-doc`: (Optional) Path to a Word template (`.docx`) to control the main output styles and document layout, overriding the built-in font, size and color defaults.
- `--resource-path`: (Optional) Directory relative to which Markdown image paths should be resolved.
- `--first-line-indent`: (Optional, compatibility flag) Body paragraphs already receive a two-character first-line indent by default; no flag is needed. Table contents remain centered with no paragraph indentation. **Note: Ignored if `--reference-doc` is used.**
- `--pandoc-path`: (Optional) Path to a custom pandoc executable. If valid, this takes priority over the system `PATH` and automated downloads.

## Examples

**Using a template and an input file (images in a separate folder):**
```powershell
uv run "$HOME\.gemini\config\skills\markdown-to-word\scripts\convert.py" convert --input-file "report.md" --output "final.docx" --reference-doc "corporate_template.docx" --resource-path ".\assets\images"
```

**Using direct text input with default formatting:**
```powershell
uv run "$HOME\.gemini\config\skills\markdown-to-word\scripts\convert.py" convert --text "Hello **World**!" --output "hello.docx"
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
- **Overusing `--text`**: Do not use `--text` for large reports. Use `--stdin` or `--input-file` instead.
- **Incorrect `--resource-path`**: Ensure the resource path aligns with how images are referenced in the Markdown. If `--input-file` is used, you often don't need this parameter at all.
