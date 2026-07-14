import importlib.util
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


PROGRAM = Path(__file__).resolve().parents[1] / "pdf_word_fidelity.py"
SPEC = importlib.util.spec_from_file_location("pdf_word_fidelity", PROGRAM)
converter = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = converter
SPEC.loader.exec_module(converter)


class AuditTests(unittest.TestCase):
    def test_token_recall_counts_repeated_words(self):
        result = converter.token_recall("alpha alpha beta", "alpha beta beta")
        self.assertEqual(result["source_token_count"], 3)
        self.assertEqual(result["matched_source_tokens"], 2)
        self.assertAlmostEqual(result["source_token_recall"], 2 / 3)

    def test_subset_font_names_normalize(self):
        self.assertEqual(converter.normalize_font_name("ABCDEF+Times-New_Roman"), "timesnewroman")

    def test_docx_audit_counts_equations_tables_and_media(self):
        document_xml = f'''<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="{converter.WORD_NS}" xmlns:m="{converter.MATH_NS}" xmlns:wp="{converter.DRAWING_NS}">
          <w:body><w:p><w:r><w:t>Lesson</w:t></w:r><w:tab/><m:oMathPara><m:oMath><m:t>x</m:t></m:oMath></m:oMathPara></w:p><w:tbl/><wp:inline/></w:body>
        </w:document>'''
        font_xml = f'''<?xml version="1.0" encoding="UTF-8"?>
        <w:fonts xmlns:w="{converter.WORD_NS}"><w:font w:name="Calibri"/></w:fonts>'''
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.docx"
            with zipfile.ZipFile(path, "w") as package:
                package.writestr("word/document.xml", document_xml)
                package.writestr("word/fontTable.xml", font_xml)
                package.writestr("word/media/image1.emf", b"not-an-image")
            audit = converter.audit_docx(path)
            source_text = converter.docx_text_snapshot(path)
        self.assertEqual(audit["omml_equations"], 1)
        self.assertEqual(audit["tables"], 1)
        self.assertEqual(audit["word_drawing_anchors"], 1)
        self.assertEqual(audit["embedded_vector_media_files"], 1)
        self.assertEqual(audit["declared_fonts"], ["calibri"])
        self.assertIn("Lesson", source_text)
        self.assertIn("x", source_text)

    def test_word_to_pdf_paths_are_stable(self):
        pdf_path, report_path = converter.safe_word_to_pdf_output_paths(Path("worksheet.docx"), Path("out"))
        self.assertEqual(pdf_path, Path("out/worksheet.converted.pdf"))
        self.assertEqual(report_path, Path("out/worksheet.word-to-pdf-report.json"))

    def test_render_page_uses_module_level_pymupdf_primitives(self):
        class FakePixmap:
            width = 4
            height = 2
            samples = b"12345678"

        class FakePage:
            def __init__(self):
                self.arguments = None

            def get_pixmap(self, **kwargs):
                self.arguments = kwargs
                return FakePixmap()

        class FakeDocument:
            def __init__(self):
                self.page = FakePage()

            def load_page(self, page_number):
                self.page_number = page_number
                return self.page

        class FakeFitz:
            csGRAY = "gray-space"

            @staticmethod
            def Matrix(x, y):
                return (x, y)

        class FakeImage:
            @staticmethod
            def frombytes(mode, size, samples):
                return (mode, size, samples)

        document = FakeDocument()
        image = converter.render_page(document, 0, 144, FakeImage, FakeFitz)
        self.assertEqual(image, ("L", (4, 2), b"12345678"))
        self.assertEqual(document.page.arguments["matrix"], (2.0, 2.0))
        self.assertEqual(document.page.arguments["colorspace"], "gray-space")


if __name__ == "__main__":
    unittest.main()
