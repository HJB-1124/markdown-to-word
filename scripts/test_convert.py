# /// script
# requires-python = ">=3.9"
# dependencies = ["python-docx"]
# ///
"""Behavioral regression tests; all outputs stay in temporary directories."""
import argparse
import importlib.util
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from lxml import etree
from docx import Document
from docx.shared import Pt

CONVERTER = Path(__file__).with_name("convert.py")
spec = importlib.util.spec_from_file_location("word_converter", CONVERTER)
converter = importlib.util.module_from_spec(spec)
spec.loader.exec_module(converter)
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}


class WordConversionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.pandoc = shutil.which("pandoc") or (
            str(converter.PANDOC_EXE) if converter.PANDOC_EXE.is_file() else None
        )
        if not cls.pandoc:
            raise unittest.SkipTest("Install Pandoc before running conversion regressions.")

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="word-conversion-test-")
        self.folder = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def convert(self, text, extra=(), success=True, stdin=False):
        source = self.folder / "input.txt"
        source.write_text(text, encoding="utf-8")
        output = self.folder / "output.docx"
        command = [sys.executable, str(CONVERTER), "convert"]
        command += ["--stdin"] if stdin else ["--input-file", str(source)]
        command += ["--output", str(output), "--pandoc-path", self.pandoc, *extra]
        result = subprocess.run(
            command, input=text if stdin else None,
            capture_output=True, text=True, encoding="utf-8",
        )
        if success:
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            with zipfile.ZipFile(output) as archive:
                self.assertIsNone(archive.testzip())
        else:
            self.assertNotEqual(result.returncode, 0)
        return output, result

    @staticmethod
    def xml(path, part="word/document.xml"):
        with zipfile.ZipFile(path) as archive:
            return etree.fromstring(archive.read(part))

    def test_mixed_delimiters_greek_cases_and_equation_numbers(self):
        text = r"""
# 数学笔记

权重 \(\theta\)，扰动 \(\delta\)。

\[
f:X\rightarrow Y
\]

\[
\hat y=f_\theta(x),\quad x\in X,\;y\in Y \tag{1}
\]

\[
\begin{cases}
x'=x+\delta\\
y'=f_\theta(x')
\end{cases}
\tag{2}
\]

距离 $L_\infty$。

$$\frac{x^2}{1+x}$$
"""
        output, result = self.convert(text)
        root = self.xml(output)
        self.assertEqual(len(root.xpath("//m:oMath", namespaces=NS)), 7)
        self.assertEqual(len(root.xpath("//m:oMathPara", namespaces=NS)), 4)
        self.assertTrue(root.xpath("//m:acc", namespaces=NS))
        self.assertTrue(root.xpath("//m:sSub", namespaces=NS))
        self.assertTrue(root.xpath("//m:m", namespaces=NS))
        self.assertTrue(root.xpath("//m:f", namespaces=NS))
        content = "".join(root.itertext())
        for token in ("θ", "δ", "→", "∈", "(1)", "(2)"):
            self.assertIn(token, content)
        self.assertIn("Validated 7 native Word equations", result.stdout)

    def test_marker_cleanup_is_explicit_and_preserves_code_and_citations(self):
        marker = "(2024TPAMI)Paper_Name…"
        code = chr(96) + marker + r" \(\theta\)" + chr(96)
        fence = chr(96) * 3
        text = (
            "正文[1]。" + marker + "\n\n"
            + "(2026IF)Multi-modal Attack Surv…\n\n"
            + r"(2025TNNLS)A_Survey_of_Attacks\_…" + "\n\n"
            + "年份 (2024)，公式 " + r"\(\delta\)" + "。\n\n"
            + code + "\n\n" + fence + "\n" + marker + "\n" + fence
        )
        original, result = self.convert(text)
        original_text = "".join(self.xml(original).itertext())
        self.assertEqual(original_text.count(marker), 3)
        self.assertIn("detected 3 pasted source filename markers", result.stderr)
        cleaned, result = self.convert(text, ["--strip-source-markers"])
        root = self.xml(cleaned)
        content = "".join(root.itertext())
        self.assertEqual(content.count(marker), 2)  # Literal code stays verbatim.
        self.assertNotIn("(2026IF)", content)
        self.assertNotIn("(2025TNNLS)", content)
        self.assertIn("[1]", content)
        self.assertIn("(2024)", content)
        self.assertEqual(len(root.xpath("//m:oMath", namespaces=NS)), 1)
        self.assertIn("Removed 3 pasted source filename markers", result.stderr)

    def test_invalid_formula_preserves_existing_output_and_exposes_warning(self):
        output = self.folder / "output.docx"
        sentinel = b"existing user document"
        output.write_bytes(sentinel)
        _, result = self.convert(r"\(\totallyinvalid{x}\)", success=False)
        self.assertEqual(output.read_bytes(), sentinel)
        self.assertIn("Could not convert TeX math", result.stderr)
        self.assertFalse(list(self.folder.glob(".markdown-to-word-*.docx")))

    def test_raw_tex_is_rejected_before_it_can_disappear(self):
        output, result = self.convert(r"\begin{equation}x=1\end{equation}", success=False)
        self.assertIn("Unconverted raw TeX", result.stderr)
        self.assertFalse(output.exists())

    def test_validation_rejects_removed_equation(self):
        output, _ = self.convert(r"\(\theta\)")
        root = self.xml(output)
        equation = root.xpath("//m:oMath", namespaces=NS)[0]
        equation.getparent().remove(equation)
        corrupted = self.folder / "missing-equation.docx"
        with zipfile.ZipFile(output) as source, zipfile.ZipFile(corrupted, "w") as dest:
            for info in source.infolist():
                data = source.read(info.filename)
                dest.writestr(info, etree.tostring(root) if info.filename == "word/document.xml" else data)
        with self.assertRaisesRegex(converter.ConversionError, "Formula validation failed"):
            converter.validate_docx(corrupted, [{"kind": "InlineMath", "labels": []}])

    def test_visible_metadata_footnotes_and_reference_doc(self):
        reference = self.folder / "reference.docx"
        template = Document()
        template.styles["Normal"].font.name = "Arial"
        template.styles["Normal"].font.size = Pt(14)
        template.save(reference)
        text = """---
title: "Title $t$"
abstract: "Abstract $a$"
custom: "Not rendered $z$"
---

- List $x$

| Value |
|---|
| $y$ |

Note[^n].

[^n]: Footnote $f$.
"""
        output, result = self.convert(text, ["--reference-doc", str(reference)])
        body, notes = self.xml(output), self.xml(output, "word/footnotes.xml")
        count = len(body.xpath("//m:oMath", namespaces=NS)) + len(notes.xpath("//m:oMath", namespaces=NS))
        self.assertEqual(count, 5)
        self.assertIn("Validated 5", result.stdout)
        self.assertEqual(Document(output).styles["Normal"].font.name, "Arial")
        self.assertEqual(len(body.xpath("//w:tbl", namespaces=NS)), 1)

    def test_default_layout_and_lists_survive_ast_conversion(self):
        output, _ = self.convert("# 标题\n\n正文。\n\n- 项目一\n- 项目二\n\n| A | B |\n|---|---|\n| 1 | 2 |\n")
        root, styles = self.xml(output), self.xml(output, "word/styles.xml")
        normal = styles.xpath("//w:style[@w:styleId='Normal']", namespaces=NS)[0]
        self.assertEqual(normal.xpath("./w:pPr/w:ind/@w:firstLineChars", namespaces=NS), ["200"])
        self.assertEqual(root.xpath("//w:pgSz/@w:w", namespaces=NS), ["11906"])
        self.assertEqual(root.xpath("//w:pgSz/@w:h", namespaces=NS), ["16838"])
        self.assertTrue(root.xpath("//w:pStyle[@w:val='Compact']", namespaces=NS))
        for cell in root.xpath("//w:tc", namespaces=NS):
            self.assertEqual(cell.xpath("./w:tcPr/w:vAlign/@w:val", namespaces=NS), ["center"])
            self.assertTrue(cell.xpath(".//w:jc[@w:val='center']", namespaces=NS))

    def test_stdin_and_explicit_starred_tags(self):
        output, _ = self.convert("中文测试 " + r"\[x=1\tag*{S1}\]", stdin=True)
        content = "".join(self.xml(output).itertext())
        self.assertIn("S1", content)
        self.assertIn("中文测试", content)
        self.assertNotIn("(S1)", content)

    def test_text_input_and_relative_image(self):
        import base64
        image = self.folder / "pixel.png"
        image.write_bytes(base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/l9sAAAAASUVORK5CYII="
        ))
        output, _ = self.convert(r"**Text** and \(\theta\)." + "\n\n![pixel](pixel.png)")
        with zipfile.ZipFile(output) as archive:
            self.assertTrue(any(name.startswith("word/media/") for name in archive.namelist()))
        direct = self.folder / "direct.docx"
        command = [
            sys.executable, str(CONVERTER), "convert", "--text", r"Text \(\delta\).",
            "--output", str(direct), "--pandoc-path", self.pandoc,
        ]
        result = subprocess.run(command, capture_output=True, text=True, encoding="utf-8")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(len(self.xml(direct).xpath("//m:oMath", namespaces=NS)), 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
