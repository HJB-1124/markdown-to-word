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
import re
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

# Resolve the canonical skill directory so junctions share one dependency cache.
SKILL_ROOT = Path(__file__).resolve().parents[1]
PANDOC_TOOL_DIR = SKILL_ROOT / ".tools" / "pandoc"
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


DEFAULT_TABLE_HEADER_FILL = "F2F2F2"
DEFAULT_TABLE_BORDER_COLOR = "D9D9D9"
DEFAULT_TABLE_BORDER_SIZE = "4"  # Eighths of a point: 0.5 pt.


def apply_default_table_formatting(root):
    """Give actual table headers a gray fill and every cell a light-gray grid."""
    def ensure(parent, tag, before=()):
        child = parent.find(qn(tag))
        if child is None:
            child = OxmlElement(tag)
            following_tags = {qn(name) for name in before}
            following = next((node for node in parent if node.tag in following_tags), None)
            if following is None:
                parent.append(child)
            else:
                following.addprevious(child)
        return child

    def set_borders(parent, tag, edges, before):
        borders = ensure(parent, tag, before)
        for edge in edges:
            border = ensure(borders, f"w:{edge}")
            border.attrib.clear()
            for attribute, value in (
                ("val", "single"), ("color", DEFAULT_TABLE_BORDER_COLOR),
                ("sz", DEFAULT_TABLE_BORDER_SIZE), ("space", "0"),
            ):
                border.set(qn(f"w:{attribute}"), value)

    cell_border_followers = (
        "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText",
        "w:vAlign", "w:hideMark", "w:headers", "w:cellIns", "w:cellDel",
        "w:cellMerge", "w:tcPrChange",
    )
    for table in root.iter(qn("w:tbl")):
        properties = ensure(table, "w:tblPr", ("w:tblGrid", "w:tr"))
        set_borders(
            properties, "w:tblBorders",
            ("top", "left", "bottom", "right", "insideH", "insideV"),
            ("w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
             "w:tblCaption", "w:tblDescription", "w:tblPrChange"),
        )
        for row in table.findall(qn("w:tr")):
            header = row.find(f"{qn('w:trPr')}/{qn('w:tblHeader')}")
            is_header = header is not None and header.get(qn("w:val"), "true") not in {
                "0", "false", "off",
            }
            # Direct children only: nested tables apply their own header state.
            for cell in row.findall(qn("w:tc")):
                properties = cell.get_or_add_tcPr()
                set_borders(
                    properties, "w:tcBorders", ("top", "left", "bottom", "right"),
                    cell_border_followers,
                )
                shade = ensure(properties, "w:shd", cell_border_followers[1:])
                shade.attrib.clear()
                shade.set(qn("w:val"), "clear")
                shade.set(qn("w:color"), "auto")
                shade.set(qn("w:fill"), DEFAULT_TABLE_HEADER_FILL if is_header else "FFFFFF")


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

    # WPS may replace Pandoc's style IDs with numeric IDs when saving a DOCX.
    # Resolve paragraph styles by their names as well as their original IDs.
    styles_root = parse_xml(next(data for info, data in entries if info.filename == "word/styles.xml"))
    paragraph_styles = {
        style.get(qn("w:styleId")): style
        for style in styles_root.findall(qn("w:style"))
        if style.get(qn("w:type")) == "paragraph"
    }

    def style_key(value):
        return (value or "").replace(" ", "").replace("_", "").casefold()

    body_names = {"normal", "bodytext", "firstparagraph", "正文", "正文文本", "首段"}
    body_style_ids = set()
    normal_style_id = None
    for style_id, style in paragraph_styles.items():
        name = style.find(qn("w:name"))
        name_key = style_key(name.get(qn("w:val"))) if name is not None else ""
        if name_key in body_names or style_key(style_id) in body_names:
            body_style_ids.add(style_id)
        if name_key in {"normal", "正文"} or style_id == "Normal":
            normal_style_id = style_id
    if normal_style_id is None:
        normal_style_id = next(
            style_id for style_id, style in paragraph_styles.items()
            if style.get(qn("w:default")) in {"1", "true"}
        )
    body_style_ids.add(normal_style_id)

    def inherited_attributes(style_id, property_tag):
        chain, seen = [], set()
        while style_id in paragraph_styles and style_id not in seen:
            seen.add(style_id)
            style = paragraph_styles[style_id]
            chain.append(style)
            base = style.find(qn("w:basedOn"))
            style_id = base.get(qn("w:val")) if base is not None else None
        attributes = {}
        for style in reversed(chain):
            properties = style.find(qn("w:pPr"))
            value = properties.find(qn(property_tag)) if properties is not None else None
            if value is not None:
                attributes.update(value.attrib)
        return attributes

    original_spacing = {
        style_id: inherited_attributes(style_id, "w:spacing") for style_id in body_style_ids
    }
    original_indentation = {
        style_id: inherited_attributes(style_id, "w:ind") for style_id in paragraph_styles
    }

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
                    style_id = style.get(qn("w:styleId"))
                    properties = style.get_or_add_rPr()
                    if style_id in body_styles or style_id in body_style_ids:
                        set_body_size(properties)
                    if style.get(qn("w:type")) == "paragraph":
                        if style_id == normal_style_id:
                            # Store body indentation in Normal itself, so reapplying
                            # the style or adding a new paragraph retains two characters.
                            set_first_line(style, True)
                        else:
                            # Preserve the previous effective indentation of other
                            # styles instead of inheriting Normal's new body indent.
                            indent = style.get_or_add_pPr().get_or_add_ind()
                            previous = original_indentation.get(style_id, {})
                            for attribute in ("firstLine", "firstLineChars"):
                                key = qn(f"w:{attribute}")
                                if key not in indent.attrib:
                                    indent.set(key, previous.get(key, "0"))
                for properties in root.iter(qn("w:rPr")):
                    set_fonts(properties.get_or_add_rFonts())

            elif info.filename == "word/document.xml":
                for section in root.iter(qn("w:sectPr")):
                    page_size = section.get_or_add_pgSz()
                    page_size.set(qn("w:w"), DEFAULT_PAGE_WIDTH_TWIPS)
                    page_size.set(qn("w:h"), DEFAULT_PAGE_HEIGHT_TWIPS)
                    page_size.set(qn("w:orient"), "portrait")

                for paragraph in root.iter(qn("w:p")):
                    properties = paragraph.find(qn("w:pPr"))
                    style = properties.find(qn("w:pStyle")) if properties is not None else None
                    style_id = style.get(qn("w:val")) if style is not None else normal_style_id
                    if style_id not in body_style_ids:
                        continue
                    if properties is not None and properties.find(qn("w:numPr")) is not None:
                        continue
                    properties = paragraph.get_or_add_pPr()
                    # BodyText/FirstParagraph can have different paragraph spacing.
                    # Preserve it while replacing their style with the built-in Normal.
                    for attribute, value in original_spacing.get(style_id, {}).items():
                        if value != original_spacing[normal_style_id].get(attribute):
                            spacing = get_or_add(properties, "w:spacing")
                            if attribute not in spacing.attrib:
                                spacing.set(attribute, value)
                    properties.get_or_add_pStyle().set(qn("w:val"), normal_style_id)
                    if any(ancestor.tag == qn("w:tc") for ancestor in paragraph.iterancestors()):
                        continue  # The table pass below clears all cell indentation.
                    has_text = any((text.text or "").strip() for text in paragraph.iter(qn("w:t")))
                    has_math = any(True for _ in paragraph.iter(qn("m:oMath")))
                    has_display_math = any(True for _ in paragraph.iter(qn("m:oMathPara")))
                    if has_display_math or (has_math and not has_text):
                        set_first_line(paragraph, False)
                    elif has_text:
                        # Inherit from Normal rather than relying on direct formatting.
                        indent = properties.find(qn("w:ind"))
                        if indent is not None:
                            for attribute in ("firstLine", "firstLineChars", "hanging", "hangingChars"):
                                indent.attrib.pop(qn(f"w:{attribute}"), None)
                            if not indent.attrib:
                                properties.remove(indent)
                    else:
                        # Empty separators/image-only paragraphs are layout blocks,
                        # not prose; do not shift their borders or anchored content.
                        set_first_line(paragraph, False)

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

            apply_default_table_formatting(root)

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
    """Use a custom Pandoc, this skill's cache, or PATH; install into the skill cache."""
    if pandoc_path:
        p = Path(pandoc_path)
        if p.exists():
            return p
        else:
            print(f"Warning: Provided pandoc path does not exist: {pandoc_path}", file=sys.stderr)

    if PANDOC_EXE.is_file():
        return PANDOC_EXE

    system_pandoc = shutil.which("pandoc")
    if system_pandoc:
        print(f"Using system Pandoc found at: {system_pandoc}", file=sys.stderr)
        return Path(system_pandoc)
        
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

MARKDOWN_READER = "markdown+tex_math_single_backslash"
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
VISIBLE_METADATA = ("title", "subtitle", "author", "date", "abstract")
SOURCE_MARKER_RE = re.compile(
    r"\((?:19|20)\d{2}[A-Z][A-Za-z0-9]{0,15}\)"
    r"[A-Za-z0-9][A-Za-z0-9 _.\-+&():/]{0,160}?(?:…|\.{3})"
)


class ConversionError(Exception):
    """Conversion cannot produce a verified Word document."""


def run_pandoc(command, content):
    result = subprocess.run(
        command, input=content, capture_output=True, text=True, encoding="utf-8"
    )
    # Pandoc can otherwise return success while replacing invalid math with text.
    if result.stderr:
        print(result.stderr.rstrip(), file=sys.stderr)
    if result.returncode:
        raise ConversionError(f"Pandoc failed with exit code {result.returncode}.")
    return result.stdout


def remove_horizontal_rules(document):
    """Drop parsed thematic-break blocks without changing Markdown source text."""
    removed = 0

    def visit(value):
        nonlocal removed
        if isinstance(value, dict):
            for child in value.values():
                visit(child)
        elif isinstance(value, list):
            retained = []
            for child in value:
                if isinstance(child, dict) and child.get("t") == "HorizontalRule":
                    removed += 1
                else:
                    visit(child)
                    retained.append(child)
            value[:] = retained

    visit(document)
    return removed


def process_source_markers(document, strip=False):
    """Detect pasted filename citations in prose; keep code and math opaque."""
    count, samples = 0, []

    def process_run(nodes):
        nonlocal count
        content = "".join(
            node["c"] if node["t"] == "Str" else " " for node in nodes
        )
        matches = list(SOURCE_MARKER_RE.finditer(content))
        count += len(matches)
        samples.extend(match.group() for match in matches[:max(0, 3 - len(samples))])
        if not strip or not matches:
            return nodes
        result, position = [], 0
        for node in nodes:
            text = node["c"] if node["t"] == "Str" else " "
            end = position + len(text)
            ranges = [
                (max(match.start(), position) - position, min(match.end(), end) - position)
                for match in matches if match.start() < end and match.end() > position
            ]
            for start, stop in reversed(ranges):
                text = text[:start] + text[stop:]
            if text:
                result.append({"t": "Str", "c": text} if node["t"] == "Str" else node)
            position = end
        return result

    def visit(value):
        if isinstance(value, dict):
            if value.get("t") in {"Math", "Code", "CodeBlock", "RawInline", "RawBlock"}:
                return value
            for key, child in value.items():
                value[key] = visit(child)
        elif isinstance(value, list):
            result, run = [], []
            for child in value:
                child = visit(child)
                if isinstance(child, dict) and child.get("t") in {"Str", "Space", "SoftBreak"}:
                    run.append(child)
                    continue
                if run:
                    result.extend(process_run(run))
                    run = []
                if strip and isinstance(child, dict) and child.get("t") in {"Para", "Plain"} and not child.get("c"):
                    continue
                result.append(child)
            if run:
                result.extend(process_run(run))
            return result
        return value

    visit(document)
    return count, samples


def rendered_nodes(document):
    """Visit body, notes and visible metadata; image alt text is not native math."""
    def walk(value):
        if isinstance(value, dict):
            yield value
            if value.get("t") in {"Code", "CodeBlock", "Image"}:
                return
            for child in value.values():
                yield from walk(child)
        elif isinstance(value, list):
            for child in value:
                yield from walk(child)

    yield from walk(document.get("blocks", []))
    for name in VISIBLE_METADATA:
        yield from walk(document.get("meta", {}).get(name, {}))


def preserve_equation_tags(tex):
    """Render explicit tags as editable math text; Pandoc ignores tag macros."""
    pattern = re.compile(r"(?<!\\)\\tag(\*)?(?![A-Za-z])\s*")
    output, labels, position = [], [], 0
    while True:
        match = pattern.search(tex, position)
        if not match:
            output.append(tex[position:])
            break
        opening = match.end()
        if opening >= len(tex) or tex[opening] != "{":
            raise ConversionError("Equation tag requires a braced label.")
        depth, cursor = 1, opening + 1
        while cursor < len(tex) and depth:
            if tex[cursor] == "\\":
                cursor += 2
                continue
            if tex[cursor] == "{":
                depth += 1
            elif tex[cursor] == "}":
                depth -= 1
            cursor += 1
        if depth:
            raise ConversionError("Equation tag has an unmatched brace.")
        label = tex[opening + 1:cursor - 1]
        visible = label if match.group(1) else "(" + label + ")"
        output.extend([tex[position:match.start()], r"\qquad\text{" + visible + "}"])
        # Complex TeX labels are checked by Pandoc; simple labels also get a text check.
        if label and not re.search(r"[\\{}]", label):
            labels.append(visible)
        position = cursor
    return "".join(output), labels


def prepare_equations(document):
    equations = []
    for node in rendered_nodes(document):
        if node.get("t") in {"RawInline", "RawBlock"} and node["c"][0] in {"tex", "latex"}:
            preview = node["c"][1].replace("\n", " ")[:100]
            raise ConversionError(
                "Unconverted raw TeX would be lost in Word: " + preview
                + ". Put equations inside supported math delimiters and retry."
            )
        if node.get("t") == "Math":
            tex, labels = preserve_equation_tags(node["c"][1])
            node["c"][1] = tex
            equations.append({"kind": node["c"][0]["t"], "labels": labels})
    return equations


def validate_docx(output_path, expected_equations):
    """Check the ZIP and native equations in the body and note stories."""
    import xml.etree.ElementTree as ET
    from collections import Counter

    equations, display_count = [], 0
    with zipfile.ZipFile(output_path) as archive:
        if archive.testzip() is not None:
            raise ConversionError("The generated DOCX archive is damaged.")
        for name in ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml"):
            if name not in archive.namelist():
                if name == "word/document.xml":
                    raise ConversionError("The generated DOCX has no document body.")
                continue
            root = ET.fromstring(archive.read(name))
            equations.extend(root.iter(f"{{{MATH_NS}}}oMath"))
            display_count += sum(1 for _ in root.iter(f"{{{MATH_NS}}}oMathPara"))
    if len(equations) != len(expected_equations):
        raise ConversionError(
            f"Formula validation failed: source has {len(expected_equations)} equations, "
            f"but Word contains {len(equations)} native equation objects."
        )
    expected_display = sum(item["kind"] == "DisplayMath" for item in expected_equations)
    if display_count != expected_display:
        raise ConversionError(
            f"Formula validation failed: expected {expected_display} display equations, "
            f"found {display_count}."
        )
    texts = []
    for equation in equations:
        text = "".join(equation.itertext())
        visible_symbol = any(
            item.tag in {f"{{{MATH_NS}}}{name}" for name in ("chr", "begChr", "endChr")}
            and item.get(f"{{{MATH_NS}}}val")
            for item in equation.iter()
        )
        if not text.strip() and not visible_symbol:
            raise ConversionError("Formula validation failed: an equation object is empty.")
        texts.append(re.sub(r"\s+", "", text))
    required_labels = Counter(
        re.sub(r"\s+", "", label)
        for equation in expected_equations for label in equation["labels"]
    )
    for label, count in required_labels.items():
        if sum(text.count(label) for text in texts) < count:
            raise ConversionError(f"Formula validation failed: equation label {label} is missing.")
    return {"equations": len(equations), "display": display_count}


def convert(args):
    """Convert through an explicit Markdown AST and commit only verified output."""
    inputs = [args.text is not None, bool(args.input_file), bool(args.stdin)]
    if sum(inputs) != 1:
        raise ConversionError("Provide exactly one of --text, --input-file, or --stdin.")
    if not args.reference_doc and not DOCX_AVAILABLE:
        raise ConversionError("Default formatting requires python-docx. Run with uv run.")
    output_path = Path(args.output).expanduser().resolve()
    if output_path.suffix.lower() != ".docx":
        raise ConversionError("--output must name a .docx file.")
    if not output_path.parent.is_dir():
        raise ConversionError(f"Output directory does not exist: {output_path.parent}")

    if args.input_file:
        source_path = Path(args.input_file).expanduser().resolve()
        content = source_path.read_text(encoding="utf-8-sig")
        resource_path = args.resource_path or str(source_path.parent)
    else:
        content = args.text if args.text is not None else sys.stdin.read()
        resource_path = args.resource_path
    pandoc_exe = ensure_pandoc_installed(args.pandoc_path)
    parsed = run_pandoc(
        [str(pandoc_exe), "--from", MARKDOWN_READER, "--to", "json", "--fail-if-warnings"],
        content,
    )
    document = json.loads(parsed)
    if not getattr(args, "keep_horizontal_rules", False):
        removed_rules = remove_horizontal_rules(document)
        if removed_rules:
            print(f"Removed {removed_rules} Markdown horizontal rules.", file=sys.stderr)
    strip_markers = getattr(args, "strip_source_markers", False)
    marker_count, marker_samples = process_source_markers(document, strip=strip_markers)
    if marker_count:
        if strip_markers:
            print(f"Removed {marker_count} pasted source filename markers.", file=sys.stderr)
        else:
            print(
                f"Notice: detected {marker_count} pasted source filename markers "
                f"({'; '.join(marker_samples)}). Review them; use --strip-source-markers "
                "when a clean document is requested.",
                file=sys.stderr,
            )
    expected_equations = prepare_equations(document)

    fd, temporary_path = tempfile.mkstemp(
        suffix=".docx", prefix=".markdown-to-word-", dir=output_path.parent
    )
    os.close(fd)
    try:
        command = [
            str(pandoc_exe), "--from", "json", "--to", "docx", "--fail-if-warnings",
            "--output", temporary_path,
        ]
        if args.reference_doc:
            command.extend(["--reference-doc", str(args.reference_doc)])
        if resource_path:
            command.append(f"--resource-path={resource_path}")
        print(f"Converting Markdown to {output_path}...", file=sys.stderr)
        run_pandoc(command, json.dumps(document, ensure_ascii=False))
        if args.first_line_indent and args.reference_doc:
            print("Notice: --reference-doc overrides --first-line-indent.", file=sys.stderr)
        if not args.reference_doc:
            apply_default_formatting(temporary_path)
        summary = validate_docx(temporary_path, expected_equations)
        if not Path(temporary_path).stat().st_size:
            raise ConversionError("The generated DOCX is empty.")
        os.replace(temporary_path, output_path)
    finally:
        if os.path.exists(temporary_path):
            os.remove(temporary_path)
    print(
        f"Validated {summary['equations']} native Word equations "
        f"({summary['display']} display, {summary['equations'] - summary['display']} inline); "
        f"preserved {sum(len(item['labels']) for item in expected_equations)} equation labels."
    )
    print(f"Success! Document successfully written to: {output_path}")


def main():
    # Keep redirected Unicode input, diagnostics and paths consistently UTF-8.
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="Markdown to Word Converter (with verified native math)")
    subparsers = parser.add_subparsers(dest="command", required=True)
    parser_convert = subparsers.add_parser("convert", help="Convert Markdown/Text to Word (.docx)")
    parser_convert.add_argument("--text", help="Raw Markdown, including LaTeX formulas")
    parser_convert.add_argument("--input-file", help="Path to an existing UTF-8 Markdown/text file")
    parser_convert.add_argument("--stdin", action="store_true", help="Read Markdown from standard input")
    parser_convert.add_argument("--output", required=True, help="Destination .docx file")
    parser_convert.add_argument("--first-line-indent", action="store_true", help="Compatibility flag; default body indentation is already two characters")
    parser_convert.add_argument("--pandoc-path", help="Path to a custom Pandoc executable")
    parser_convert.add_argument("--reference-doc", help="Word template overriding built-in formatting")
    parser_convert.add_argument("--resource-path", help="Directory for local images/resources")
    parser_convert.add_argument(
        "--strip-source-markers", action="store_true",
        help="Remove pasted filename citations such as (2024TPAMI)Paper_Name… from prose",
    )
    parser_convert.add_argument(
        "--keep-horizontal-rules", action="store_true",
        help="Keep Markdown thematic breaks such as ---; they are omitted by default",
    )
    args = parser.parse_args()
    try:
        convert(args)
    except (ConversionError, OSError, ValueError, zipfile.BadZipFile) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
