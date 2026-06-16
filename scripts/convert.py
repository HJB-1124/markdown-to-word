import argparse
import json
import os
import subprocess
import sys
import tempfile
import time
import urllib.request
import zipfile
from pathlib import Path

# Paths
PANDOC_TOOL_DIR = Path.home() / ".gemini" / "tools" / "pandoc"
PANDOC_VERSION = "3.1.13"
PANDOC_ZIP_URL = f"https://github.com/jgm/pandoc/releases/download/{PANDOC_VERSION}/pandoc-{PANDOC_VERSION}-windows-x86_64.zip"
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
                    raise RateLimitError(f"Rate limit exceeded fetching {url}. Try again later.")
            elif e.code in (500, 502, 503, 504):
                delay = (2 ** attempt)
                print(f"Server error {e.code}. Retrying in {delay}s...", file=sys.stderr)
                time.sleep(delay)
            else:
                body = e.read().decode('utf-8', errors='ignore')
                print(f"Error {e.code} downloading Pandoc: {body}", file=sys.stderr)
                sys.exit(1)
        except urllib.error.URLError as e:
            delay = (2 ** attempt)
            print(f"Network error: {e.reason}. Retrying in {delay}s...", file=sys.stderr)
            time.sleep(delay)
            if attempt == max_retries - 1:
                print("Max retries reached. Network error.", file=sys.stderr)
                sys.exit(1)

def ensure_pandoc_installed():
    """Ensure pandoc is installed in the global tools directory."""
    if PANDOC_EXE.exists():
        return PANDOC_EXE
        
    print("Pandoc not found locally. Preparing to install...", file=sys.stderr)
    PANDOC_TOOL_DIR.mkdir(parents=True, exist_ok=True)
    zip_path = PANDOC_TOOL_DIR / "pandoc.zip"
    
    download_with_retry(PANDOC_ZIP_URL, zip_path)
    
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
    if not args.text and not args.input_file:
        print("Error: Must provide either --text or --input-file.", file=sys.stderr)
        sys.exit(1)
        
    if args.text and args.input_file:
        print("Error: Provide only one of --text or --input-file.", file=sys.stderr)
        sys.exit(1)
        
    pandoc_exe = ensure_pandoc_installed()
    
    # Prepare input
    temp_md_path = None
    if args.text:
        fd, temp_md_path = tempfile.mkstemp(suffix=".md", text=True)
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            f.write(args.text)
        input_path = temp_md_path
    else:
        input_path = args.input_file
        
    # Execute Pandoc
    print(f"Running Pandoc to convert {input_path} to {args.output}...", file=sys.stderr)
    try:
        result = subprocess.run(
            [str(pandoc_exe), str(input_path), "-o", str(args.output)],
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
            
    print(f"Success! Document successfully written to: {args.output}")

def main():
    parser = argparse.ArgumentParser(description="Markdown to Word Converter (with Math support)")
    subparsers = parser.add_subparsers(dest="command", required=True)
    
    parser_convert = subparsers.add_parser("convert", help="Convert Markdown/Text to Word (.docx)")
    parser_convert.add_argument("--text", help="Raw text or markdown containing LaTeX formulas")
    parser_convert.add_argument("--input-file", help="Path to an existing .md file to convert")
    parser_convert.add_argument("--output", required=True, help="Destination path for the .docx file")
    
    args = parser.parse_args()
    
    if args.command == "convert":
        convert(args)

if __name__ == "__main__":
    main()
