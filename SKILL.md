---
name: markdown-to-word
description: >-
  Converts a text string or Markdown file containing text and LaTeX mathematical
  formulas into a perfectly formatted native Microsoft Word (.docx) file using
  Pandoc. Automatically handles Pandoc installation if missing.
---

# Markdown to Word Converter

## Overview
This skill converts raw text or Markdown (including complex LaTeX mathematical formulas) into a native Microsoft Word (`.docx`) file. It uses `Pandoc` under the hood to ensure LaTeX equations are properly converted into Word's native Office Math Markup Language (OMML), avoiding the unformatted text output commonly seen with simple Python libraries. The skill automatically downloads and installs a persistent copy of Pandoc to `~/.gemini/tools/pandoc/` upon first use.

## Dependencies
None.

## Quick Start
To use the skill, run the helper script `convert.py`.

```bash
uv run ~/.gemini/config/skills/markdown-to-word/scripts/convert.py convert --text "This is an equation: \$\$E=mc^2\$\$" --output "output.docx"
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
uv run ~/.gemini/config/skills/markdown-to-word/scripts/convert.py convert --text "Hello **World**!" --output "hello.docx"
```

Using an existing file:
```bash
uv run ~/.gemini/config/skills/markdown-to-word/scripts/convert.py convert --input-file "notes.md" --output "notes.docx"
```

## Rate Limiting
No specific rate limits apply to the usage of the conversion script. However, the one-time automated download of Pandoc interacts with the GitHub API. The script implements exponential backoff with retries to handle potential HTTP 429 rate limits or transient server errors from GitHub.

## Common Mistakes
- **Omitting Required Output**: The `--output` argument is strictly required. Forgetting it will cause an argument parsing error.
- **Providing Both Text and File**: Supplying both `--text` and `--input-file` will result in an error. Choose one.
- **Malformed LaTeX**: If the provided LaTeX is severely malformed, Pandoc might fail to compile it and will return a non-zero exit code. The script will surface the Pandoc stderr so you can see exactly where the parse error occurred.
