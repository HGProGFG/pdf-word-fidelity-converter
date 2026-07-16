# PDF Word Fidelity Converter

This Windows program converts a PDF into an editable Word `.docx` or exports a Word document into a print-quality PDF, then makes the result auditable. It includes a simple teacher-facing desktop app as well as a command line for technical users.

It deliberately uses the desktop Microsoft Word PDF importer. That is the local conversion engine most likely to retain editable paragraphs, tables, positioned elements, and Word equation objects when the PDF contains enough semantic information. It does **not** pretend that every PDF can be reconstructed perfectly: a PDF preserves final page appearance, not the original Word structure. Scanned PDFs, flattened equations, outlined text, damaged fonts, and complex vector diagrams may need manual repair or the original authoring file.

The converter protects against silent loss by creating:

- `name.editable.docx` - the editable Word result.
- `name.roundtrip.pdf` - Word's print-quality PDF export (unless skipped).
- `name.fidelity-report.json` - table, drawing, media, declared-font, OMML-equation, text-recall, page-size, and per-page visual-diff evidence.
- `name.visual-diffs/` - PNG differences for any page that exceeds the visual tolerance.

## Requirements

- Windows 10/11.
- A licensed desktop installation of Microsoft Word (Word 2013 or newer; Microsoft 365 recommended). Word must have been opened and activated once for the signed-in Windows user.
- Python 3.10+.

Install the Python packages from this folder:

```powershell
python -m pip install -r requirements.txt
```

## Give it to a teacher

On the computer where you are preparing the app, open PowerShell in this folder and run:

```powershell
.\build_windows_app.ps1
```

This creates `dist\TeacherPdfConverter\TeacherPdfConverter.exe`. Give your friend the entire `TeacherPdfConverter` folder, not the `.exe` by itself. They can double-click `TeacherPdfConverter.exe`, choose a PDF, and press **Convert PDF to Editable Word**.

Microsoft Word must be installed and activated on their computer. The app does not upload their teaching materials or change the original PDF. It saves the editable Word document, round-trip PDF, review report, and any needed visual difference images in the folder they choose.

For best results, encourage them to use the original document when available. For example, start from the worksheet's Word file rather than a print-to-PDF copy; this retains actual editable equations and diagrams.

## Vietnamese app and easy installation

The Vietnamese interface is [teacher_pdf_converter_vi.py](teacher_pdf_converter_vi.py). Build its single portable executable with:

```powershell
.\build_vietnamese_exe.ps1
```

The result is `dist\ChuyenDoiPDFSangWord.exe`; it includes Python and its required conversion libraries. Give the teacher the entire project folder and have them double-click `CaiDatUngDungTiengViet.cmd` once. It adds a **Chuyển PDF sang Word** Start Menu shortcut without administrator permissions. Their computer still needs an activated desktop copy of Microsoft Word.

Vietnamese instructions are in [HUONG_DAN_TIENG_VIET.md](HUONG_DAN_TIENG_VIET.md).

## Use

First check the computer is ready:

```powershell
python .\pdf_word_fidelity.py --diagnose
```

Convert a PDF and send all artifacts to a new folder:

```powershell
python .\pdf_word_fidelity.py `
  --input "C:\Documents\source.pdf" `
  --output-dir "C:\Documents\source-word-output"
```

The default visual gate compares every original page to the PDF Word exports at 144 DPI. The process exits with code `2` if a page differs beyond the configured limits or source text is missing from the round-trip PDF. That is intentional: it prevents an unchecked result being treated as faithful.

To keep outputs even when the fidelity gate fails, add `--allow-difference`. The DOCX and report are still written either way.

```powershell
python .\pdf_word_fidelity.py --input "C:\Documents\source.pdf" --allow-difference
```

For a faster DOCX-only conversion (no round-trip proof), use `--skip-roundtrip`. This should only be used for drafts.

```powershell
python .\pdf_word_fidelity.py --input "C:\Documents\source.pdf" --skip-roundtrip
```

## Reviewing equations and diagrams

Open the `.docx` in Word and click each formula. An editable Word equation should expose Word's Equation ribbon; the report's `omml_equations` count records how many native Office Math objects were actually found. If mathematical symbols were present in the PDF but that count is zero, the report issues a warning instead of claiming those formulas are editable.

Check the retained visual-difference PNGs before accepting a failed gate, especially on pages containing equations, diagrams, tables, headers/footers, and non-Latin text. A small visual difference can be caused by font hinting, while a shifted equation, page-size mismatch, or lower text recall needs correction.

For a truly guaranteed editable source, obtain the original Word/LaTeX/InDesign document whenever possible. No PDF-to-DOCX converter can recover authoring semantics that were flattened or rasterized in the supplied PDF.

## Useful options

```text
--dpi 200                         More sensitive visual comparison
--max-mean-delta 0.03             Tighter mean image-difference limit
--max-changed-pixel-ratio 0.08    Tighter changed-pixel limit
--text-overlap-threshold 0.99     Require more source text in output PDF
--keep-all-diffs                  Save visual diff PNGs for passing pages too
--overwrite                       Replace existing output artifacts
```

The report contains the exact configured thresholds, so it can be retained with a document review record.

## Batch conversion and mathematical symbols

The Vietnamese desktop app supports selecting multiple PDFs and Word documents in one operation. It processes them one at a time so Microsoft Word can preserve layout reliably, and stores every result in its own numbered output folder.

During PDF-to-Word conversion, mathematical Unicode characters are assigned to **Cambria Math**. The converter can safely restore a missing numeric-set symbol such as `∈ □` only when the original PDF text unambiguously identifies it as `∈ ℤ`, `∈ ℕ`, `∈ ℚ`, or `∈ ℝ`. Any remaining placeholder is reported for manual review rather than guessed.

## Word to PDF

Choose a `.docx`, `.docm`, or legacy `.doc` file in either desktop app, then use the **Export Word to Print-Quality PDF** button. The converter uses Microsoft Word's native PDF engine, preserving the document's current fonts, spacing, page layout, tables, drawings, images, and equation rendering.

For `.docx` and `.docm` input it also produces a `name.word-to-pdf-report.json` file with source-font, table, drawing, media, Office Math, and text-token checks. A Word file has no fixed-page visual reference until Word lays it out, so this direction uses Word's print output and text/font evidence rather than a pixel comparison against a source PDF.
