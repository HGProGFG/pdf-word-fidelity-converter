#!/usr/bin/env python3
"""Fidelity-first PDF -> editable DOCX -> PDF converter for Windows.

The converter deliberately uses Microsoft Word's own PDF reflow engine instead
of trying to reconstruct a Word document from raw PDF drawing commands. It then
creates a round-trip PDF and produces an auditable report with visual, text,
font, table, drawing, image, and equation checks.

No PDF converter can guarantee an editable, byte-for-byte identical Word file:
a PDF stores positioned output, not the author's semantic Word structure. This
tool makes that limitation visible by failing the validation gate unless the
round-trip result is within the supplied tolerances.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import time
import traceback
import zipfile
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET


APP_NAME = "PDF Word Fidelity Converter"
REPORT_VERSION = 1
WORD_FORMAT_DOCX = 16
WORD_EXPORT_PDF = 17
WORD_ALERTS_NONE = 0
WORD_OPTIMIZE_PRINT = 0
WORD_BOOKMARK_HEADINGS = 1

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

MATH_SIGNAL = re.compile(
    r"[∑∏∫√∞≈≠≤≥±×÷∂∇∈∉∪∩⊂⊃⇒⇔αβγδεζηθικλμνξπρστυφχψωΑΒΓΔΕΖΗΘΙΚΛΜΝΞΠΡΣΤΥΦΧΨΩ]"
)
WORD_TOKEN = re.compile(r"[\w]+", re.UNICODE)
SUBSET_FONT = re.compile(r"^[A-Z]{6}\+")


class ConversionError(RuntimeError):
    """A conversion or validation prerequisite is unavailable."""


@dataclass(frozen=True)
class PageComparison:
    page: int
    source_size_points: tuple[float, float]
    roundtrip_size_points: tuple[float, float]
    size_delta_points: tuple[float, float]
    mean_abs_delta: float | None
    changed_pixel_ratio: float | None
    passed: bool
    diff_image: str | None
    note: str | None = None


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Convert a PDF with Microsoft Word, export it back to PDF, and "
            "produce an evidence-based fidelity report."
        )
    )
    parser.add_argument("--input", type=Path, help="Source PDF.")
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="Destination folder. Defaults to a sibling folder named <pdf>-word-output.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace existing output files in the destination folder.",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=144,
        help="Resolution used for round-trip visual comparison (default: 144).",
    )
    parser.add_argument(
        "--max-mean-delta",
        type=float,
        default=0.05,
        help="Maximum per-page mean grayscale difference, from 0 to 1 (default: 0.05).",
    )
    parser.add_argument(
        "--max-changed-pixel-ratio",
        type=float,
        default=0.15,
        help="Maximum share of materially changed pixels, from 0 to 1 (default: 0.15).",
    )
    parser.add_argument(
        "--text-overlap-threshold",
        type=float,
        default=0.97,
        help="Minimum source-token recall in the round-trip PDF, from 0 to 1 (default: 0.97).",
    )
    parser.add_argument(
        "--allow-difference",
        action="store_true",
        help="Return exit code 0 even when the visual or text fidelity gate fails.",
    )
    parser.add_argument(
        "--keep-all-diffs",
        action="store_true",
        help="Keep diff images for passing pages too; failed-page diffs are always kept.",
    )
    parser.add_argument(
        "--skip-roundtrip",
        action="store_true",
        help="Create DOCX only. No PDF export or fidelity gate is run.",
    )
    parser.add_argument(
        "--diagnose",
        action="store_true",
        help="Print environment diagnostics without converting a document.",
    )
    return parser.parse_args(argv)


def ensure_range(label: str, value: float, *, minimum: float = 0.0, maximum: float = 1.0) -> None:
    if not minimum <= value <= maximum:
        raise ConversionError(f"{label} must be between {minimum} and {maximum}; got {value}.")


def require_dependencies(*, visual: bool) -> tuple[Any, Any | None, Any | None]:
    """Import optional runtime dependencies only when they are needed."""
    try:
        import win32com.client  # type: ignore[import-not-found]
        import pythoncom  # type: ignore[import-not-found]
    except ImportError as exc:
        raise ConversionError(
            "pywin32 is required to automate desktop Microsoft Word. Install dependencies "
            "with: python -m pip install -r requirements.txt"
        ) from exc

    fitz = image = None
    if visual:
        try:
            import fitz  # type: ignore[import-not-found]
            from PIL import Image, ImageChops, ImageOps, ImageStat  # type: ignore[import-not-found]
        except ImportError as exc:
            raise ConversionError(
                "PyMuPDF and Pillow are required for the visual fidelity gate. Install dependencies "
                "with: python -m pip install -r requirements.txt"
            ) from exc
        image = (Image, ImageChops, ImageOps, ImageStat)
    return (win32com.client, pythoncom, (fitz, image))


def word_is_installed() -> bool:
    """Check standard Word locations plus App Paths without importing pywin32."""
    candidates = [
        Path(os.environ.get("ProgramFiles", r"C:\\Program Files")) / "Microsoft Office/root/Office16/WINWORD.EXE",
        Path(os.environ.get("ProgramFiles(x86)", r"C:\\Program Files (x86)")) / "Microsoft Office/root/Office16/WINWORD.EXE",
    ]
    if any(candidate.exists() for candidate in candidates):
        return True
    try:
        import winreg

        for key_path in (
            r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WINWORD.EXE",
            r"SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WINWORD.EXE",
        ):
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                    value, _ = winreg.QueryValueEx(key, None)
                    if Path(value).exists():
                        return True
            except OSError:
                continue
    except ImportError:
        pass
    return shutil.which("WINWORD.EXE") is not None


def environment_diagnostics() -> dict[str, Any]:
    dependencies: dict[str, bool] = {}
    for module in ("win32com.client", "pythoncom", "fitz", "PIL"):
        try:
            __import__(module)
            dependencies[module] = True
        except ImportError:
            dependencies[module] = False
    return {
        "application": APP_NAME,
        "python": sys.version.split()[0],
        "platform": sys.platform,
        "microsoft_word_detected": word_is_installed(),
        "dependencies": dependencies,
        "ready": word_is_installed() and all(dependencies.values()),
    }


def safe_output_paths(input_pdf: Path, output_dir: Path) -> tuple[Path, Path, Path, Path]:
    stem = input_pdf.stem
    return (
        output_dir / f"{stem}.editable.docx",
        output_dir / f"{stem}.roundtrip.pdf",
        output_dir / f"{stem}.fidelity-report.json",
        output_dir / f"{stem}.visual-diffs",
    )


def check_inputs(args: argparse.Namespace) -> None:
    if args.dpi < 72 or args.dpi > 600:
        raise ConversionError("--dpi must be between 72 and 600.")
    ensure_range("--max-mean-delta", args.max_mean_delta)
    ensure_range("--max-changed-pixel-ratio", args.max_changed_pixel_ratio)
    ensure_range("--text-overlap-threshold", args.text_overlap_threshold)
    if args.input is None:
        raise ConversionError("--input is required unless --diagnose is used.")
    if not args.input.is_file():
        raise ConversionError(f"Input PDF was not found: {args.input}")
    if args.input.suffix.lower() != ".pdf":
        raise ConversionError("--input must point to a .pdf file.")


def create_output_dir(output_dir: Path, paths: Iterable[Path], overwrite: bool) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    occupied = [path for path in paths if path.exists()]
    if occupied and not overwrite:
        names = ", ".join(str(path) for path in occupied)
        raise ConversionError(f"Refusing to overwrite existing outputs. Use --overwrite: {names}")


def convert_with_word(input_pdf: Path, output_docx: Path) -> None:
    """Open PDF through Word's native reflow engine and save a DOCX."""
    win32com, pythoncom, _ = require_dependencies(visual=False)
    pythoncom.CoInitialize()
    word = None
    document = None
    try:
        word = win32com.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = WORD_ALERTS_NONE
        word.ScreenUpdating = False
        document = word.Documents.Open(
            FileName=str(input_pdf.resolve()),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
            OpenAndRepair=True,
            NoEncodingDialog=True,
        )
        document.SaveAs2(
            FileName=str(output_docx.resolve()),
            FileFormat=WORD_FORMAT_DOCX,
            AddToRecentFiles=False,
            CompatibilityMode=15,
        )
    except Exception as exc:  # COM errors do not share a reliable exception type.
        raise ConversionError(
            "Microsoft Word could not convert the PDF. Confirm that a licensed desktop "
            "version of Word is installed and that the PDF is not password-protected. "
            f"Details: {exc}"
        ) from exc
    finally:
        if document is not None:
            try:
                document.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit(SaveChanges=False)
            except Exception:
                pass
        pythoncom.CoUninitialize()
    if not output_docx.exists() or output_docx.stat().st_size == 0:
        raise ConversionError("Word reported success but did not create a DOCX output.")


def export_word_to_pdf(input_docx: Path, output_pdf: Path) -> None:
    """Export the converted DOCX back to PDF through Word's print-quality exporter."""
    win32com, pythoncom, _ = require_dependencies(visual=False)
    pythoncom.CoInitialize()
    word = None
    document = None
    try:
        word = win32com.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = WORD_ALERTS_NONE
        word.ScreenUpdating = False
        document = word.Documents.Open(
            FileName=str(input_docx.resolve()),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
            OpenAndRepair=True,
            NoEncodingDialog=True,
        )
        document.ExportAsFixedFormat(
            OutputFileName=str(output_pdf.resolve()),
            ExportFormat=WORD_EXPORT_PDF,
            OpenAfterExport=False,
            OptimizeFor=WORD_OPTIMIZE_PRINT,
            Range=0,
            Item=0,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=WORD_BOOKMARK_HEADINGS,
            DocStructureTags=True,
            BitmapMissingFonts=False,
            UseISO19005_1=False,
        )
    except Exception as exc:
        raise ConversionError(f"Microsoft Word could not export the DOCX to PDF. Details: {exc}") from exc
    finally:
        if document is not None:
            try:
                document.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit(SaveChanges=False)
            except Exception:
                pass
        pythoncom.CoUninitialize()
    if not output_pdf.exists() or output_pdf.stat().st_size == 0:
        raise ConversionError("Word reported success but did not create the round-trip PDF.")


def normalize_font_name(name: str) -> str:
    name = SUBSET_FONT.sub("", name or "")
    return re.sub(r"[\s,_-]+", "", name).casefold()


def page_fonts(page: Any) -> set[str]:
    # get_fonts(full=False) tuple layout is xref, extension, type, basefont, name, encoding.
    return {normalize_font_name(str(font[3])) for font in page.get_fonts(full=False) if len(font) > 3 and font[3]}


def token_counts(text: str) -> dict[str, int]:
    counts: dict[str, int] = {}
    for token in WORD_TOKEN.findall(text.casefold()):
        counts[token] = counts.get(token, 0) + 1
    return counts


def token_recall(source_text: str, output_text: str) -> dict[str, Any]:
    source = token_counts(source_text)
    output = token_counts(output_text)
    source_total = sum(source.values())
    matched = sum(min(count, output.get(token, 0)) for token, count in source.items())
    return {
        "source_token_count": source_total,
        "roundtrip_token_count": sum(output.values()),
        "matched_source_tokens": matched,
        "source_token_recall": (matched / source_total) if source_total else None,
        "note": "No extractable source text; text recall is not applicable." if not source_total else None,
    }


def pdf_snapshot(pdf_path: Path, fitz: Any) -> dict[str, Any]:
    document = fitz.open(pdf_path)
    try:
        all_text: list[str] = []
        fonts: set[str] = set()
        page_sizes: list[tuple[float, float]] = []
        image_count = 0
        drawings = 0
        for page in document:
            all_text.append(page.get_text("text"))
            fonts.update(page_fonts(page))
            page_sizes.append((round(page.rect.width, 3), round(page.rect.height, 3)))
            image_count += len(page.get_images(full=True))
            drawings += len(page.get_drawings())
        text = "\n".join(all_text)
        return {
            "path": str(pdf_path),
            "page_count": document.page_count,
            "page_sizes_points": page_sizes,
            "extractable_text_characters": len(text),
            "extractable_text": text,
            "fonts": sorted(fonts),
            "image_references": image_count,
            "vector_drawing_paths": drawings,
            "mathematical_symbol_count": len(MATH_SIGNAL.findall(text)),
        }
    finally:
        document.close()


def _xml_count(root: ET.Element, tag: str) -> int:
    return sum(1 for _ in root.iter(tag))


def audit_docx(docx_path: Path) -> dict[str, Any]:
    """Inspect DOCX package contents without modifying the Word document."""
    try:
        with zipfile.ZipFile(docx_path) as package:
            names = package.namelist()
            document_xml = package.read("word/document.xml")
            root = ET.fromstring(document_xml)
            font_table = package.read("word/fontTable.xml") if "word/fontTable.xml" in names else b""
    except (KeyError, zipfile.BadZipFile, ET.ParseError) as exc:
        raise ConversionError(f"The converted file is not a readable DOCX package: {exc}") from exc

    fonts: list[str] = []
    if font_table:
        try:
            font_root = ET.fromstring(font_table)
            fonts = sorted(
                {
                    normalize_font_name(node.attrib.get(f"{{{WORD_NS}}}name", ""))
                    for node in font_root.iter(f"{{{WORD_NS}}}font")
                    if node.attrib.get(f"{{{WORD_NS}}}name")
                }
            )
        except ET.ParseError:
            fonts = []

    media_extensions = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf", ".svg"}
    media = [name for name in names if name.startswith("word/media/") and Path(name).suffix.lower() in media_extensions]
    return {
        "path": str(docx_path),
        "omml_equations": _xml_count(root, f"{{{MATH_NS}}}oMath"),
        "omml_equation_paragraphs": _xml_count(root, f"{{{MATH_NS}}}oMathPara"),
        "tables": _xml_count(root, f"{{{WORD_NS}}}tbl"),
        "word_drawing_anchors": _xml_count(root, f"{{{DRAWING_NS}}}anchor") + _xml_count(root, f"{{{DRAWING_NS}}}inline"),
        "embedded_media_files": len(media),
        "embedded_vector_media_files": sum(Path(name).suffix.lower() in {".emf", ".wmf", ".svg"} for name in media),
        "declared_fonts": fonts,
    }


def render_page(pdf: Any, page_number: int, dpi: int, image_cls: Any, fitz: Any) -> Any:
    """Render one page using module-level PyMuPDF primitives.

    A ``fitz.Document`` owns pages but does not expose ``Matrix`` or ``csGRAY``;
    those belong to the imported PyMuPDF module. Keeping them separate avoids a
    runtime failure during the post-conversion fidelity check.
    """
    page = pdf.load_page(page_number)
    scale = dpi / 72
    pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), colorspace=fitz.csGRAY, alpha=False)
    return image_cls.frombytes("L", (pixmap.width, pixmap.height), pixmap.samples)


def compare_visuals(
    source_pdf: Path,
    roundtrip_pdf: Path,
    diff_dir: Path,
    *,
    dpi: int,
    max_mean_delta: float,
    max_changed_pixel_ratio: float,
    keep_all_diffs: bool,
    fitz: Any,
    image_tools: tuple[Any, Any, Any, Any],
) -> tuple[list[PageComparison], bool]:
    """Render matching pages, calculate pixel deltas, and retain reviewable diff PNGs."""
    Image, ImageChops, ImageOps, ImageStat = image_tools
    source = fitz.open(source_pdf)
    output = fitz.open(roundtrip_pdf)
    comparisons: list[PageComparison] = []
    page_count_matches = source.page_count == output.page_count
    try:
        comparable_count = min(source.page_count, output.page_count)
        for index in range(comparable_count):
            source_page = source.load_page(index)
            output_page = output.load_page(index)
            source_size = (round(source_page.rect.width, 3), round(source_page.rect.height, 3))
            output_size = (round(output_page.rect.width, 3), round(output_page.rect.height, 3))
            size_delta = (round(output_size[0] - source_size[0], 3), round(output_size[1] - source_size[1], 3))
            source_image = render_page(source, index, dpi, Image, fitz)
            output_image = render_page(output, index, dpi, Image, fitz)
            same_canvas = source_image.size == output_image.size
            if not same_canvas:
                output_image = output_image.resize(source_image.size, Image.Resampling.LANCZOS)
            difference = ImageChops.difference(source_image, output_image)
            stats = ImageStat.Stat(difference)
            mean = float(stats.mean[0]) / 255.0
            pixels = list(difference.getdata())
            changed_ratio = sum(value > 20 for value in pixels) / len(pixels) if pixels else 0.0
            passed = (
                same_canvas
                and size_delta == (0.0, 0.0)
                and mean <= max_mean_delta
                and changed_ratio <= max_changed_pixel_ratio
            )
            diff_file: str | None = None
            if keep_all_diffs or not passed:
                diff_dir.mkdir(parents=True, exist_ok=True)
                file_path = diff_dir / f"page-{index + 1:03d}-difference.png"
                # Autocontrast makes small differences visible in manual review.
                ImageOps.autocontrast(difference).save(file_path)
                diff_file = str(file_path)
            note = None
            if not same_canvas:
                note = "Rendered page dimensions differ; image was resized only to calculate a diagnostic delta."
            comparisons.append(
                PageComparison(
                    page=index + 1,
                    source_size_points=source_size,
                    roundtrip_size_points=output_size,
                    size_delta_points=size_delta,
                    mean_abs_delta=round(mean, 6),
                    changed_pixel_ratio=round(changed_ratio, 6),
                    passed=passed,
                    diff_image=diff_file,
                    note=note,
                )
            )
        for index in range(comparable_count, max(source.page_count, output.page_count)):
            comparisons.append(
                PageComparison(
                    page=index + 1,
                    source_size_points=(0.0, 0.0),
                    roundtrip_size_points=(0.0, 0.0),
                    size_delta_points=(0.0, 0.0),
                    mean_abs_delta=None,
                    changed_pixel_ratio=None,
                    passed=False,
                    diff_image=None,
                    note="This page exists only in one PDF; page counts differ.",
                )
            )
    finally:
        source.close()
        output.close()
    return comparisons, page_count_matches and all(item.passed for item in comparisons)


def build_warnings(source: dict[str, Any], docx: dict[str, Any], text: dict[str, Any], text_threshold: float) -> list[str]:
    warnings: list[str] = []
    source_fonts = set(source["fonts"])
    docx_fonts = set(docx["declared_fonts"])
    missing_fonts = sorted(source_fonts - docx_fonts)
    if missing_fonts:
        warnings.append(
            "The DOCX font table does not declare all source PDF fonts: " + ", ".join(missing_fonts[:12])
        )
    if source["mathematical_symbol_count"] and not docx["omml_equations"]:
        warnings.append(
            "Mathematical symbols were detected in the source, but no editable OMML equation objects "
            "were found in the DOCX. Review equations in Word before accepting the conversion."
        )
    recall = text["source_token_recall"]
    if recall is not None and recall < text_threshold:
        warnings.append(
            f"Text-token recall is {recall:.2%}, below the configured {text_threshold:.2%} threshold."
        )
    if source["extractable_text_characters"] == 0:
        warnings.append(
            "The source appears image-only or protected. OCR/reconstruction may be needed for editable text."
        )
    return warnings


def write_report(path: Path, report: dict[str, Any]) -> None:
    path.write_text(json.dumps(report, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def convert(args: argparse.Namespace) -> tuple[dict[str, Any], bool]:
    check_inputs(args)
    output_dir = args.output_dir or args.input.parent / f"{args.input.stem}-word-output"
    output_docx, output_pdf, report_path, diff_dir = safe_output_paths(args.input, output_dir)
    targets = [output_docx, report_path] if args.skip_roundtrip else [output_docx, output_pdf, report_path]
    create_output_dir(output_dir, targets, args.overwrite)

    _, _, visual_modules = require_dependencies(visual=not args.skip_roundtrip)
    fitz = image_tools = None
    if visual_modules is not None:
        fitz, image_tools = visual_modules

    started = time.monotonic()
    source = pdf_snapshot(args.input, fitz) if fitz is not None else {"path": str(args.input)}
    convert_with_word(args.input, output_docx)
    docx = audit_docx(output_docx)
    report: dict[str, Any] = {
        "report_version": REPORT_VERSION,
        "application": APP_NAME,
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "input_pdf": str(args.input.resolve()),
        "editable_docx": str(output_docx.resolve()),
        "roundtrip_pdf": None,
        "thresholds": {
            "dpi": args.dpi,
            "max_mean_delta": args.max_mean_delta,
            "max_changed_pixel_ratio": args.max_changed_pixel_ratio,
            "text_overlap_threshold": args.text_overlap_threshold,
        },
        "source_pdf": source,
        "docx_editability_audit": docx,
        "roundtrip": None,
        "warnings": [],
        "passed": None,
    }

    if args.skip_roundtrip:
        report["warnings"].append("Round-trip validation was skipped at the command line.")
        report["passed"] = None
        report["elapsed_seconds"] = round(time.monotonic() - started, 2)
        write_report(report_path, report)
        return report, True

    assert fitz is not None and image_tools is not None
    export_word_to_pdf(output_docx, output_pdf)
    roundtrip = pdf_snapshot(output_pdf, fitz)
    text = token_recall(source["extractable_text"], roundtrip["extractable_text"])
    comparisons, visual_passed = compare_visuals(
        args.input,
        output_pdf,
        diff_dir,
        dpi=args.dpi,
        max_mean_delta=args.max_mean_delta,
        max_changed_pixel_ratio=args.max_changed_pixel_ratio,
        keep_all_diffs=args.keep_all_diffs,
        fitz=fitz,
        image_tools=image_tools,
    )
    text_passed = text["source_token_recall"] is None or text["source_token_recall"] >= args.text_overlap_threshold
    report["roundtrip_pdf"] = str(output_pdf.resolve())
    report["roundtrip"] = {
        "pdf_snapshot": roundtrip,
        "text_comparison": text,
        "visual_comparison": {
            "page_count_matches": source["page_count"] == roundtrip["page_count"],
            "passed": visual_passed,
            "pages": [asdict(comparison) for comparison in comparisons],
        },
    }
    report["warnings"] = build_warnings(source, docx, text, args.text_overlap_threshold)
    report["passed"] = visual_passed and text_passed
    report["elapsed_seconds"] = round(time.monotonic() - started, 2)
    write_report(report_path, report)
    return report, bool(report["passed"])


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if args.diagnose:
        diagnostics = environment_diagnostics()
        print(json.dumps(diagnostics, indent=2))
        return 0 if diagnostics["ready"] else 3
    try:
        report, passed = convert(args)
    except ConversionError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 3
    except Exception as exc:  # Preserve diagnostics from unexpected automation failures.
        print(f"UNEXPECTED ERROR: {exc}", file=sys.stderr)
        traceback.print_exc()
        return 4

    print(f"Editable DOCX: {report['editable_docx']}")
    if report["roundtrip_pdf"]:
        print(f"Round-trip PDF: {report['roundtrip_pdf']}")
    print(f"Fidelity report: {Path(report['editable_docx']).with_name(Path(report['editable_docx']).stem.replace('.editable', '') + '.fidelity-report.json')}")
    if report["passed"] is None:
        print("Round-trip validation was skipped.")
        return 0
    if passed:
        print("Fidelity gate: PASSED")
        return 0
    print("Fidelity gate: FAILED - inspect the JSON report and retained difference images.", file=sys.stderr)
    return 0 if args.allow_difference else 2


if __name__ == "__main__":
    raise SystemExit(main())
