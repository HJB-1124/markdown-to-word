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
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import qn
    from lxml import etree
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

DEFAULT_EAST_ASIA_FONT = "宋体"
DEFAULT_LATIN_FONT = "Times New Roman"
DEFAULT_MATH_FONT = "Cambria Math"
DEFAULT_BODY_SIZE_HALF_POINTS = "21"  # 五号 = 10.5 pt; Word stores half-points.
DEFAULT_FIRST_LINE_CHARS = "200"  # Word stores firstLineChars in hundredths of a character.
DEFAULT_FIRST_LINE_TWIPS = "420"  # 21 pt fallback for readers without character-unit support.
DEFAULT_PAGE_WIDTH_TWIPS = "11906"  # A4 portrait: 210 x 297 mm.
DEFAULT_PAGE_HEIGHT_TWIPS = "16838"


def apply_default_formatting(output_path):
    """Apply the built-in typography when no reference document is supplied."""
    def get_or_add(parent, tag, first=False):
        child = parent.find(qn(tag))
        if child is None:
            child = OxmlElement(tag)
            if first:
                parent.insert(0, child)
            else:
                parent.append(child)
        return child

    def set_fonts(fonts, latin=DEFAULT_LATIN_FONT, east_asia=DEFAULT_EAST_ASIA_FONT):
        for attribute in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "csTheme", "cstheme"):
            fonts.attrib.pop(qn(f"w:{attribute}"), None)
        for attribute in ("ascii", "hAnsi", "cs"):
            fonts.set(qn(f"w:{attribute}"), latin)
        fonts.set(qn("w:eastAsia"), east_asia)

    def set_black(color):
        color.attrib.clear()
        color.set(qn("w:val"), "000000")

    def set_body_size(properties):
        size = properties.get_or_add_sz()
        size.set(qn("w:val"), DEFAULT_BODY_SIZE_HALF_POINTS)
        complex_size = properties.find(qn("w:szCs"))
        if complex_size is None:
            complex_size = OxmlElement("w:szCs")
            size.addnext(complex_size)
        complex_size.set(qn("w:val"), DEFAULT_BODY_SIZE_HALF_POINTS)

    def set_first_line(paragraph, indent_body):
        indent = paragraph.get_or_add_pPr().get_or_add_ind()
        for attribute in ("hanging", "hangingChars"):
            indent.attrib.pop(qn(f"w:{attribute}"), None)
        indent.set(qn("w:firstLineChars"), DEFAULT_FIRST_LINE_CHARS if indent_body else "0")
        indent.set(qn("w:firstLine"), DEFAULT_FIRST_LINE_TWIPS if indent_body else "0")

    with zipfile.ZipFile(output_path) as source:
        entries = [(info, source.read(info)) for info in source.infolist()]

    updated_entries = []
    for info, data in entries:
        if info.filename.startswith("word/") and info.filename.endswith(".xml"):
            root = parse_xml(data)
            if info.filename == "word/styles.xml":
                defaults = get_or_add(root, "w:docDefaults")
                properties = get_or_add(get_or_add(defaults, "w:rPrDefault"), "w:rPr")
                set_body_size(properties)
                body_styles = {
                    "Normal", "BodyText", "FirstParagraph", "BodyTextChar", "Compact", "VerbatimChar"
                }
                for style in root.findall(qn("w:style")):
                    properties = style.get_or_add_rPr()
                    if style.get(qn("w:styleId")) in body_styles:
                        set_body_size(properties)
                for properties in root.iter(qn("w:rPr")):
                    set_fonts(properties.get_or_add_rFonts())

            elif info.filename == "word/document.xml":
                for section in root.iter(qn("w:sectPr")):
                    page_size = section.get_or_add_pgSz()
                    page_size.set(qn("w:w"), DEFAULT_PAGE_WIDTH_TWIPS)
                    page_size.set(qn("w:h"), DEFAULT_PAGE_HEIGHT_TWIPS)
                    page_size.set(qn("w:orient"), "portrait")

                body_paragraph_styles = {"Normal", "BodyText", "FirstParagraph", "正文", "首段"}
                for paragraph in root.iter(qn("w:p")):
                    if any(ancestor.tag == qn("w:tc") for ancestor in paragraph.iterancestors()):
                        continue
                    properties = paragraph.find(qn("w:pPr"))
                    style = properties.find(qn("w:pStyle")) if properties is not None else None
                    style_id = style.get(qn("w:val")) if style is not None else "Normal"
                    if style_id not in body_paragraph_styles:
                        continue
                    if properties is not None and properties.find(qn("w:numPr")) is not None:
                        continue
                    has_text = any((text.text or "").strip() for text in paragraph.iter(qn("w:t")))
                    has_math = any(True for _ in paragraph.iter(qn("m:oMath")))
                    has_display_math = any(True for _ in paragraph.iter(qn("m:oMathPara")))
                    if has_display_math or (has_math and not has_text):
                        set_first_line(paragraph, False)
                    elif has_text:
                        set_first_line(paragraph, True)

            elif info.filename == "word/settings.xml":
                math_properties = root.find(qn("m:mathPr"))
                if math_properties is None:
                    math_properties = OxmlElement("m:mathPr")
                    root.insert_element_before(
                        math_properties, "w:uiCompat97To2003", "w:attachedSchema",
                        "w:themeFontLang", "w:clrSchemeMapping", "w:doNotIncludeSubdocsInStats",
                        "w:doNotAutoCompressPictures", "w:forceUpgrade", "w:captions",
                        "w:readModeInkLockDown", "w:smartTagType", "w:shapeDefaults",
                        "w:doNotEmbedSmartTags", "w:decimalSymbol", "w:listSeparator",
                    )
                get_or_add(math_properties, "m:mathFont", first=True).set(qn("m:val"), DEFAULT_MATH_FONT)

            elif info.filename.startswith("word/theme/"):
                # Match the theme to explicit styles so inherited text is consistent.
                for tag in ("a:majorFont", "a:minorFont"):
                    for family in root.iter(qn(tag)):
                        for tag_name, font in (
                            ("a:latin", DEFAULT_LATIN_FONT),
                            ("a:ea", DEFAULT_EAST_ASIA_FONT),
                            ("a:cs", DEFAULT_LATIN_FONT),
                        ):
                            face = family.find(qn(tag_name))
                            if face is not None:
                                face.set("typeface", font)
                        for face in family.findall(qn("a:font")):
                            if face.get("script") in {"Hans", "Hant"}:
                                face.set("typeface", DEFAULT_EAST_ASIA_FONT)

            # Native equation text uses the math font rather than body-text styles.
            for formula in root.iter(qn("m:oMath")):
                for run in formula.iter(qn("m:r")):
                    properties = run.find(qn("w:rPr"))
                    if properties is None:
                        properties = OxmlElement("w:rPr")
                        math_run_properties = run.find(qn("m:rPr"))
                        if math_run_properties is None:
                            run.insert(0, properties)
                        else:
                            math_run_properties.addnext(properties)
                for properties in formula.iter(qn("w:rPr")):
                    set_fonts(properties.get_or_add_rFonts(), DEFAULT_MATH_FONT, DEFAULT_MATH_FONT)

            # Cover nested tables and tables in headers, footers and footnotes too.
            for cell in root.iter(qn("w:tc")):
                cell.get_or_add_tcPr().get_or_add_vAlign().set(qn("w:val"), "center")
                for paragraph in cell.iter(qn("w:p")):
                    properties = paragraph.get_or_add_pPr()
                    properties.get_or_add_jc().set(qn("w:val"), "center")
                    indent = properties.get_or_add_ind()
                    indent.attrib.clear()
                    for attribute in ("left", "right", "firstLine", "firstLineChars"):
                        indent.set(qn(f"w:{attribute}"), "0")
                for math_paragraph in cell.iter(qn("m:oMathPara")):
                    properties = get_or_add(math_paragraph, "m:oMathParaPr", first=True)
                    get_or_add(properties, "m:jc").set(qn("m:val"), "center")

            # Include direct formatting, syntax highlighting, hyperlinks, footnotes,
            # headers, footers, numbering and native math control properties.
            for properties in root.iter(qn("w:rPr")):
                set_black(properties.get_or_add_color())
            for color in root.iter(qn("w:color")):
                set_black(color)
            data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        updated_entries.append((info, data))

    # Replace only after the complete archive has been written successfully.
    output_path = Path(output_path)
    fd, temporary_path = tempfile.mkstemp(suffix=".docx", dir=output_path.resolve().parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(temporary_path, "w") as destination:
            for info, data in updated_entries:
                destination.writestr(info, data)
        os.replace(temporary_path, output_path)
    finally:
        if os.path.exists(temporary_path):
            os.remove(temporary_path)


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

    if not args.reference_doc and not DOCX_AVAILABLE:
        print("Error: Default formatting requires python-docx. Run this script with uv run.", file=sys.stderr)
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
            
    if args.first_line_indent and args.reference_doc:
        print("Notice: --reference-doc is used. Skipping --first-line-indent (assume template handles indentation).", file=sys.stderr)

    if not args.reference_doc:
        print("Applying A4 portrait, default fonts/colors, two-character body indents and centered table contents...", file=sys.stderr)
        apply_default_formatting(args.output)
            
    print(f"Success! Document successfully written to: {args.output}")

def main():
    parser = argparse.ArgumentParser(description="Markdown to Word Converter (with Math support)")
    subparsers = parser.add_subparsers(dest="command", required=True)
    
    parser_convert = subparsers.add_parser("convert", help="Convert Markdown/Text to Word (.docx)")
    parser_convert.add_argument("--text", help="Raw text or markdown containing LaTeX formulas")
    parser_convert.add_argument("--input-file", help="Path to an existing .md file to convert")
    parser_convert.add_argument("--stdin", action="store_true", help="Read raw text or markdown from standard input")
    parser_convert.add_argument("--output", required=True, help="Destination path for the .docx file")
    parser_convert.add_argument("--first-line-indent", action="store_true", help="Compatibility flag; body paragraphs already default to a two-character first-line indent")
    parser_convert.add_argument("--pandoc-path", help="Path to a custom pandoc executable")
    parser_convert.add_argument("--reference-doc", help="Path to a Word template (.docx); overrides built-in formatting")
    parser_convert.add_argument("--resource-path", help="Path to search for images and other resources")
    
    args = parser.parse_args()
    
    if args.command == "convert":
        convert(args)

if __name__ == "__main__":
    main()
