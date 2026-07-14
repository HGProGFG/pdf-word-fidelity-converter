#!/usr/bin/env python3
"""Simple desktop interface for PDF Word Fidelity Converter.

Run this file with Python on Windows, or package it with build_windows_app.ps1.
"""

from __future__ import annotations

import os
import queue
import threading
import traceback
import tkinter as tk
from argparse import Namespace
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pdf_word_fidelity as converter


APP_TITLE = "Teacher PDF to Word Converter"


class TeacherConverterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.minsize(720, 460)
        self.resizable(True, False)
        self.pdf_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.status = tk.StringVar(value="Choose a PDF or Word document to convert.")
        self.open_folder_after = tk.BooleanVar(value=True)
        self.keep_all_diffs = tk.BooleanVar(value=False)
        self._events: queue.Queue[tuple[str, object]] = queue.Queue()
        self._build()
        self.after(150, self._check_events)

    def _build(self) -> None:
        container = ttk.Frame(self, padding=24)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)

        ttk.Label(container, text="PDF ↔ Word Converter", font=("Segoe UI", 18, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w"
        )
        ttk.Label(
            container,
            text=(
                "Convert PDFs into editable Word documents or export Word documents as print-quality PDFs. "
                "The app creates an audit for text, fonts, tables, drawings, and equations."
            ),
            wraplength=640,
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 22))

        ttk.Label(container, text="PDF or Word file").grid(row=2, column=0, sticky="w", padx=(0, 12))
        ttk.Entry(container, textvariable=self.pdf_path).grid(row=2, column=1, sticky="ew")
        ttk.Button(container, text="Choose file…", command=self._choose_pdf).grid(row=2, column=2, padx=(10, 0))

        ttk.Label(container, text="Save results in").grid(row=3, column=0, sticky="w", padx=(0, 12), pady=(14, 0))
        ttk.Entry(container, textvariable=self.output_dir).grid(row=3, column=1, sticky="ew", pady=(14, 0))
        ttk.Button(container, text="Choose folder…", command=self._choose_output).grid(
            row=3, column=2, padx=(10, 0), pady=(14, 0)
        )

        options = ttk.Frame(container)
        options.grid(row=4, column=0, columnspan=3, sticky="w", pady=(16, 0))
        ttk.Checkbutton(options, text="Open the results folder when finished", variable=self.open_folder_after).pack(
            side="left", padx=(0, 22)
        )
        ttk.Checkbutton(options, text="Save difference images when converting PDF to Word", variable=self.keep_all_diffs).pack(side="left")

        self.progress = ttk.Progressbar(container, mode="indeterminate")
        self.progress.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(24, 8))
        ttk.Label(container, textvariable=self.status, wraplength=650).grid(
            row=6, column=0, columnspan=3, sticky="w"
        )

        self.convert_button = ttk.Button(
            container,
            text="Choose a file to begin",
            command=self._start_conversion,
        )
        self.convert_button.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(24, 0), ipady=7)

        ttk.Separator(container).grid(row=8, column=0, columnspan=3, sticky="ew", pady=18)
        ttk.Label(
            container,
            text=(
                "Tip: When converting PDF to Word, click formulas to check editability. When exporting Word to PDF, "
                "review formulas, diagrams, tables, and special characters."
            ),
            wraplength=650,
            foreground="#555555",
        ).grid(row=9, column=0, columnspan=3, sticky="w")

    def _choose_pdf(self) -> None:
        selected = filedialog.askopenfilename(
            title="Choose a file to convert",
            filetypes=[("PDF files", "*.pdf"), ("Word documents", "*.docx *.docm *.doc"), ("All files", "*.*")],
        )
        if not selected:
            return
        document = Path(selected)
        self.pdf_path.set(str(document))
        if document.suffix.lower() == ".pdf":
            self.output_dir.set(str(document.parent / f"{document.stem}-word-output"))
            self.convert_button.configure(text="Convert PDF to Editable Word")
            self.status.set("Ready to convert PDF to Word. The original PDF will not be changed.")
        else:
            self.output_dir.set(str(document.parent / f"{document.stem}-pdf-output"))
            self.convert_button.configure(text="Export Word to Print-Quality PDF")
            self.status.set("Ready to export Word to PDF. The original Word document will not be changed.")

    def _choose_output(self) -> None:
        selected = filedialog.askdirectory(title="Choose where to save the converted files")
        if selected:
            self.output_dir.set(selected)

    def _start_conversion(self) -> None:
        input_document = Path(self.pdf_path.get()).expanduser()
        output_text = self.output_dir.get().strip()
        output_dir = Path(output_text).expanduser() if output_text else None
        if not input_document.is_file() or input_document.suffix.lower() not in {".pdf", ".docx", ".docm", ".doc"}:
            messagebox.showerror(APP_TITLE, "Please choose an existing PDF or Word document first.")
            return
        if output_dir is None:
            messagebox.showerror(APP_TITLE, "Please choose a folder for the converted files.")
            return
        diagnostics = converter.environment_diagnostics()
        if not diagnostics["ready"]:
            missing = []
            if not diagnostics["microsoft_word_detected"]:
                missing.append("a desktop installation of Microsoft Word")
            for module, installed in diagnostics["dependencies"].items():
                if not installed:
                    missing.append(module)
            messagebox.showerror(
                APP_TITLE,
                "This computer is not ready yet. Missing: "
                + ", ".join(missing)
                + "\n\nAsk the person who installed the app to run the setup instructions.",
            )
            return

        self.convert_button.configure(state="disabled")
        self.progress.start(12)
        if input_document.suffix.lower() == ".pdf":
            self.status.set("Converting PDF to Word. Large or scanned PDFs can take a few minutes…")
        else:
            self.status.set("Exporting the Word document to a print-quality PDF…")
        options = Namespace(
            input=input_document,
            output_dir=output_dir,
            overwrite=True,
            dpi=144,
            max_mean_delta=0.05,
            max_changed_pixel_ratio=0.15,
            text_overlap_threshold=0.97,
            allow_difference=True,
            keep_all_diffs=self.keep_all_diffs.get(),
            skip_roundtrip=False,
            diagnose=False,
        )
        threading.Thread(target=self._convert_worker, args=(options,), daemon=True).start()

    def _convert_worker(self, options: Namespace) -> None:
        try:
            report, passed = converter.convert(options)
            self._events.put(("success", (report, passed)))
        except Exception as exc:
            self._events.put(("error", (str(exc), traceback.format_exc())))

    def _check_events(self) -> None:
        try:
            while True:
                event, payload = self._events.get_nowait()
                self.progress.stop()
                self.convert_button.configure(state="normal")
                if event == "success":
                    report, passed = payload  # type: ignore[misc]
                    output_dir = Path(report["output_directory"])
                    is_word_to_pdf = report["conversion_type"] == "word_to_pdf"
                    if self.open_folder_after.get():
                        os.startfile(output_dir)  # type: ignore[attr-defined]
                    if passed:
                        self.status.set("Finished - the visual and text fidelity checks passed.")
                        success_message = (
                            "Finished. The PDF export passed the automatic checks.\n\n"
                            if is_word_to_pdf
                            else "Finished. The Word document and round-trip PDF passed the automatic checks.\n\n"
                        ) + f"Saved in:\n{output_dir}"
                        messagebox.showinfo(
                            APP_TITLE,
                            success_message,
                        )
                    else:
                        self.status.set("Finished - the files were created, but the result needs a quick review.")
                        messagebox.showwarning(
                            APP_TITLE,
                            "Finished, but please review the result before sharing it.\n\n"
                            "Open the fidelity report in the results folder. "
                            "Pay special attention to formulas, diagrams, and tables.\n\n"
                            f"Saved in:\n{output_dir}",
                        )
                else:
                    message, details = payload  # type: ignore[misc]
                    self.status.set("The conversion did not finish. The original file is unchanged.")
                    messagebox.showerror(
                        APP_TITLE,
                        f"The conversion could not finish:\n\n{message}\n\n"
                        "If this is a password-protected or scanned PDF, it may need a different source file or OCR."
                    )
                    print(details)
        except queue.Empty:
            pass
        self.after(150, self._check_events)


if __name__ == "__main__":
    TeacherConverterApp().mainloop()
