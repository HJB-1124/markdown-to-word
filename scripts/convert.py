# /// script
# requires-python = ">=3.9"
# dependencies = [
#     "python-docx",
# ]
# ///
import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request
import urllib.error
import zipfile
from pathlib import Path

try:
    import docx
    from docx.shared import Pt
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# Paths
PANDOC_TOOL_DIR = Path.home() / ".gemini" / "tools" / "pandoc"
PANDOC_VERSION = "3.1.13"
PANDOC_ZIP_URL = f"https://github.com/jgm/pandoc/releases/download/{PANDOC_VERSION}/pandoc-{PANDOC_VERSION}-windows-x86_64.zip"
PANDOC_EXPECTED_SHA256 = "347b250e855d0cb4bf9e7dcda3b55a2b5bdae588e637a363810008b75be5ca81"
PANDOC_EXE_DIR = PANDOC_TOOL_DIR / f"pandoc-{PANDOC_VERSION}"
PANDOC_EXE = PANDOC_EXE_DIR / "pandoc.exe"

class RateLimitError(Exception):
    pass

def download_with_retry(url, dest_path, max_retries=3):
    """Download file with exponential backoff for GitHub releases."""
    for attempt in range(max_retries):
        try:
            print(f"Downloading Pandoc from {url} (Attempt {attempt+1}/{max_retries})...", file=sys.stderr)
            urllib.request.urlretrieve(url, dest_path)
            return
        except urllib.error.HTTPError as e:
            if e.code == 429:
                delay = (2 ** attempt) * 2
                print(f"Rate limited by GitHub. Waiting {delay}s...", file=sys.stderr)
                time.sleep(delay)
                if attempt == max_retries - 1:
                    if os.path.exists(dest_path): os.remove(dest_path)
                    raise RateLimitError(f"Rate limit exceeded fetching {url}. Try again later.")
            elif e.code in (500, 502, 503, 504):
                delay = (2 ** attempt)
                print(f"Server error {e.code}. Retrying in {delay}s...", file=sys.stderr)
                time.sleep(delay)
                if attempt == max_retries - 1:
                    if os.path.exists(dest_path): os.remove(dest_path)
                    sys.exit(1)
            else:
                body = e.read().decode('utf-8', errors='ignore')
                print(f"Error {e.code} downloading Pandoc: {body}", file=sys.stderr)
                if os.path.exists(dest_path): os.remove(dest_path)
                sys.exit(1)
        except urllib.error.URLError as e:
            delay = (2 ** attempt)
            print(f"Network error: {e.reason}. Retrying in {delay}s...", file=sys.stderr)
            time.sleep(delay)
            if attempt == max_retries - 1:
                print("Max retries reached. Network error.", file=sys.stderr)
                if os.path.exists(dest_path): os.remove(dest_path)
                sys.exit(1)

def ensure_pandoc_installed(pandoc_path=None):
    """Ensure pandoc is installed in the global tools directory, or use provided/system path."""
    if pandoc_path:
        p = Path(pandoc_path)
        if p.exists():
            return p
        else:
            print(f"Warning: Provided pandoc path does not exist: {pandoc_path}", file=sys.stderr)

    system_pandoc = shutil.which("pandoc")
    if system_pandoc:
        print(f"Using system Pandoc found at: {system_pandoc}", file=sys.stderr)
        return Path(system_pandoc)

    if PANDOC_EXE.exists():
        return PANDOC_EXE
        
    print("Pandoc not found locally. Preparing to install...", file=sys.stderr)
    PANDOC_TOOL_DIR.mkdir(parents=True, exist_ok=True)
    zip_path = PANDOC_TOOL_DIR / "pandoc.zip"
    
    download_with_retry(PANDOC_ZIP_URL, zip_path)
    
    print("Verifying download integrity...", file=sys.stderr)
    with open(zip_path, 'rb') as f:
        file_hash = hashlib.sha256(f.read()).hexdigest()
    if file_hash.lower() != PANDOC_EXPECTED_SHA256.lower():
        print(f"Security Error: Downloaded file hash ({file_hash}) does not match expected ({PANDOC_EXPECTED_SHA256}).", file=sys.stderr)
        zip_path.unlink()
        sys.exit(1)
    
    print("Extracting Pandoc...", file=sys.stderr)
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(PANDOC_TOOL_DIR)
    finally:
        if zip_path.exists():
            zip_path.unlink()
            
    if not PANDOC_EXE.exists():
        print(f"Extraction failed. Cannot find {PANDOC_EXE}.", file=sys.stderr)
        sys.exit(1)
        
    print("Pandoc successfully installed.", file=sys.stderr)
    return PANDOC_EXE

def convert(args):
    """Subcommand: convert"""
    if not args.text and not args.input_file and not args.stdin:
        print("Error: Must provide one of --text, --input-file, or --stdin.", file=sys.stderr)
        sys.exit(1)
        
    inputs_provided = sum([bool(args.text), bool(args.input_file), bool(args.stdin)])
    if inputs_provided > 1:
        print("Error: Provide only one of --text, --input-file, or --stdin.", file=sys.stderr)
        sys.exit(1)
        
    pandoc_exe = ensure_pandoc_installed(args.pandoc_path)
    
    # Prepare input
    temp_md_path = None
    if args.text or args.stdin:
        content = args.text if args.text else sys.stdin.read()
        fd, temp_md_path = tempfile.mkstemp(suffix=".md", text=True)
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            f.write(content)
        input_path = temp_md_path
    else:
        input_path = args.input_file
        
    # Execute Pandoc
    print(f"Running Pandoc to convert {input_path} to {args.output}...", file=sys.stderr)
    try:
        command = [str(pandoc_exe), str(input_path), "-o", str(args.output)]
        
        if args.reference_doc:
            command.extend(["--reference-doc", str(args.reference_doc)])
            
        resource_path = args.resource_path
        if not resource_path and args.input_file:
            resource_path = os.path.dirname(os.path.abspath(args.input_file))
            if not resource_path:
                resource_path = "."
                
        if resource_path:
            command.extend([f"--resource-path={resource_path}"])

        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding='utf-8'
        )
        if result.returncode != 0:
            print("Pandoc Conversion Failed!", file=sys.stderr)
            print("Stdout:", result.stdout, file=sys.stderr)
            print("Stderr:", result.stderr, file=sys.stderr)
            sys.exit(1)
    finally:
        # Cleanup temp file
        if temp_md_path and os.path.exists(temp_md_path):
            os.remove(temp_md_path)
            
    if args.first_line_indent:
        if args.reference_doc:
            print("Notice: --reference-doc is used. Skipping --first-line-indent (assume template handles indentation).", file=sys.stderr)
        elif not DOCX_AVAILABLE:
            print("Warning: python-docx is not installed. Skipping first-line indent.", file=sys.stderr)
        else:
            print("Applying first-line indent...", file=sys.stderr)
            doc = docx.Document(args.output)
            
            def apply_indent(container):
                for paragraph in container.paragraphs:
                    # Apply to standard text styles, including potential localized names
                    if paragraph.style.name in ['Normal', 'First Paragraph', 'Body Text', '正文', '首段']:
                        # Pt(21) is roughly equal to a 2-character indent for standard 10.5pt font
                        paragraph.paragraph_format.first_line_indent = Pt(21)
                for table in getattr(container, 'tables', []):
                    for row in table.rows:
                        for cell in row.cells:
                            apply_indent(cell)
                            
            apply_indent(doc)
            doc.save(args.output)
            
    print(f"Success! Document successfully written to: {args.output}")

def main():
    parser = argparse.ArgumentParser(description="Markdown to Word Converter (with Math support)")
    subparsers = parser.add_subparsers(dest="command", required=True)
    
    parser_convert = subparsers.add_parser("convert", help="Convert Markdown/Text to Word (.docx)")
    parser_convert.add_argument("--text", help="Raw text or markdown containing LaTeX formulas")
    parser_convert.add_argument("--input-file", help="Path to an existing .md file to convert")
    parser_convert.add_argument("--stdin", action="store_true", help="Read raw text or markdown from standard input")
    parser_convert.add_argument("--output", required=True, help="Destination path for the .docx file")
    parser_convert.add_argument("--first-line-indent", action="store_true", help="Add a first-line indent to all standard paragraphs")
    parser_convert.add_argument("--pandoc-path", help="Path to a custom pandoc executable")
    parser_convert.add_argument("--reference-doc", help="Path to a Word template (.docx) to style the output")
    parser_convert.add_argument("--resource-path", help="Path to search for images and other resources")
    
    args = parser.parse_args()
    
    if args.command == "convert":
        convert(args)

if __name__ == "__main__":
    main()
