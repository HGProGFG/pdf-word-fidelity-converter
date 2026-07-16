#!/usr/bin/env python3
"""Giao diện tiếng Việt cho ứng dụng chuyển PDF và Word của thầy Bảo."""

from __future__ import annotations

import os
import queue
import threading
import traceback
import tkinter as tk
from argparse import Namespace
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

import pdf_word_fidelity as converter


APP_TITLE = "Chuyển PDF sang Word cho thầy Bảo"
SUPPORTED_SUFFIXES = {".pdf", ".docx", ".docm", ".doc"}


class VietnameseTeacherConverterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.minsize(760, 480)
        self.resizable(True, False)
        self.selection_text = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.status = tk.StringVar(value="Hãy chọn một hoặc nhiều tệp PDF / Word cần chuyển đổi.")
        self.open_folder_after = tk.BooleanVar(value=True)
        self.keep_all_diffs = tk.BooleanVar(value=False)
        self.selected_files: list[Path] = []
        self._events: queue.Queue[tuple[str, object]] = queue.Queue()
        self._build()
        self.after(150, self._check_events)

    def _build(self) -> None:
        container = ttk.Frame(self, padding=24)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)

        ttk.Label(container, text="Chuyển PDF ↔ Word cho thầy Bảo", font=("Segoe UI", 18, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w"
        )
        ttk.Label(
            container,
            text=(
                "Chọn nhiều tệp PDF hoặc Word cùng lúc. Ứng dụng lần lượt tạo Word có thể chỉnh sửa "
                "hoặc PDF chất lượng in, kèm báo cáo kiểm tra văn bản, phông chữ, bảng, hình và công thức."
            ),
            wraplength=690,
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 22))

        ttk.Label(container, text="Tệp PDF hoặc Word").grid(row=2, column=0, sticky="w", padx=(0, 12))
        ttk.Entry(container, textvariable=self.selection_text).grid(row=2, column=1, sticky="ew")
        ttk.Button(container, text="Chọn nhiều tệp…", command=self._choose_files).grid(row=2, column=2, padx=(10, 0))

        ttk.Label(container, text="Lưu kết quả tại").grid(row=3, column=0, sticky="w", padx=(0, 12), pady=(14, 0))
        ttk.Entry(container, textvariable=self.output_dir).grid(row=3, column=1, sticky="ew", pady=(14, 0))
        ttk.Button(container, text="Chọn thư mục…", command=self._choose_output).grid(
            row=3, column=2, padx=(10, 0), pady=(14, 0)
        )

        options = ttk.Frame(container)
        options.grid(row=4, column=0, columnspan=3, sticky="w", pady=(16, 0))
        ttk.Checkbutton(options, text="Mở thư mục kết quả khi hoàn tất", variable=self.open_folder_after).pack(
            side="left", padx=(0, 22)
        )
        ttk.Checkbutton(options, text="Lưu ảnh so sánh khi chuyển PDF sang Word", variable=self.keep_all_diffs).pack(side="left")

        self.progress = ttk.Progressbar(container, mode="determinate", value=0)
        self.progress.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(24, 8))
        ttk.Label(container, textvariable=self.status, wraplength=690).grid(
            row=6, column=0, columnspan=3, sticky="w"
        )

        self.convert_button = ttk.Button(
            container,
            text="Chọn tệp để bắt đầu",
            command=self._start_conversion,
        )
        self.convert_button.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(24, 0), ipady=7)

        ttk.Separator(container).grid(row=8, column=0, columnspan=3, sticky="ew", pady=18)
        ttk.Label(
            container,
            text=(
                "Mẹo: Các tệp được xử lý lần lượt để Microsoft Word giữ nguyên định dạng. "
                "Ký hiệu toán học được đặt phông Cambria Math; hãy kiểm tra công thức, sơ đồ, bảng biểu "
                "và ký tự tiếng Việt trong các kết quả có báo “cần kiểm tra”."
            ),
            wraplength=690,
            foreground="#555555",
        ).grid(row=9, column=0, columnspan=3, sticky="w")

    def _choose_files(self) -> None:
        selected = filedialog.askopenfilenames(
            title="Chọn các tệp cần chuyển",
            filetypes=[("Tệp được hỗ trợ", "*.pdf *.docx *.docm *.doc"), ("Tệp PDF", "*.pdf"), ("Tài liệu Word", "*.docx *.docm *.doc")],
        )
        if not selected:
            return
        self.selected_files = [Path(item) for item in selected]
        if len(self.selected_files) == 1:
            self.selection_text.set(str(self.selected_files[0]))
            document = self.selected_files[0]
            suffix = "Word-da-chinh-sua" if document.suffix.lower() == ".pdf" else "PDF-da-chuyen-doi"
            self.output_dir.set(str(document.parent / f"{document.stem}-{suffix}"))
        else:
            preview = ", ".join(path.name for path in self.selected_files[:3])
            if len(self.selected_files) > 3:
                preview += ", …"
            self.selection_text.set(f"Đã chọn {len(self.selected_files)} tệp: {preview}")
            self.output_dir.set(str(self.selected_files[0].parent / "ket-qua-chuyen-doi"))
        self._update_selection_status()

    def _update_selection_status(self) -> None:
        count = len(self.selected_files)
        pdfs = sum(path.suffix.lower() == ".pdf" for path in self.selected_files)
        words = count - pdfs
        if count == 1:
            if pdfs:
                self.convert_button.configure(text="Chuyển PDF thành Word có thể chỉnh sửa")
                self.status.set("Sẵn sàng chuyển PDF sang Word. Tệp PDF gốc sẽ không bị thay đổi.")
            else:
                self.convert_button.configure(text="Chuyển Word thành PDF chất lượng in")
                self.status.set("Sẵn sàng xuất Word sang PDF. Tệp Word gốc sẽ không bị thay đổi.")
        elif pdfs == count:
            self.convert_button.configure(text=f"Chuyển {count} PDF thành Word")
            self.status.set(f"Sẵn sàng chuyển lần lượt {count} tệp PDF sang Word.")
        elif words == count:
            self.convert_button.configure(text=f"Chuyển {count} Word thành PDF")
            self.status.set(f"Sẵn sàng xuất lần lượt {count} tài liệu Word sang PDF.")
        else:
            self.convert_button.configure(text=f"Chuyển {count} tệp đã chọn")
            self.status.set(f"Sẵn sàng chuyển lần lượt {pdfs} PDF và {words} tài liệu Word.")

    def _choose_output(self) -> None:
        selected = filedialog.askdirectory(title="Chọn nơi lưu các tệp đã chuyển đổi")
        if selected:
            self.output_dir.set(selected)

    def _input_documents(self) -> list[Path]:
        if self.selected_files:
            return list(self.selected_files)
        typed_path = Path(self.selection_text.get().strip()).expanduser()
        return [typed_path] if typed_path.is_file() else []

    def _start_conversion(self) -> None:
        input_documents = self._input_documents()
        output_text = self.output_dir.get().strip()
        output_root = Path(output_text).expanduser() if output_text else None
        if not input_documents or any(not path.is_file() or path.suffix.lower() not in SUPPORTED_SUFFIXES for path in input_documents):
            messagebox.showerror(APP_TITLE, "Vui lòng chọn một hoặc nhiều tệp PDF / Word có sẵn trước.")
            return
        if output_root is None:
            messagebox.showerror(APP_TITLE, "Vui lòng chọn thư mục để lưu kết quả.")
            return
        diagnostics = converter.environment_diagnostics()
        if not diagnostics["ready"]:
            missing = []
            if not diagnostics["microsoft_word_detected"]:
                missing.append("Microsoft Word bản cài đặt trên máy")
            missing.extend(module for module, installed in diagnostics["dependencies"].items() if not installed)
            messagebox.showerror(
                APP_TITLE,
                "Máy tính chưa sẵn sàng. Thiếu: " + ", ".join(missing) + "\n\nHãy chạy phần thiết lập trước.",
            )
            return

        self.convert_button.configure(state="disabled")
        self.progress.configure(value=0, maximum=len(input_documents))
        self.status.set(f"Đang chuẩn bị xử lý {len(input_documents)} tệp bằng Microsoft Word…")
        threading.Thread(target=self._convert_worker, args=(input_documents, output_root), daemon=True).start()

    def _task_output_dir(self, output_root: Path, document: Path, index: int, total: int) -> Path:
        if total == 1:
            return output_root
        suffix = "Word-da-chinh-sua" if document.suffix.lower() == ".pdf" else "PDF-da-chuyen-doi"
        return output_root / f"{index:02d}-{document.stem}-{suffix}"

    def _convert_worker(self, input_documents: list[Path], output_root: Path) -> None:
        results: list[dict[str, Any]] = []
        total = len(input_documents)
        try:
            for index, document in enumerate(input_documents, start=1):
                self._events.put(("progress", (index, total, document)))
                options = Namespace(
                    input=document,
                    output_dir=self._task_output_dir(output_root, document, index, total),
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
                try:
                    report, passed = converter.convert(options)
                    results.append({"input": document, "report": report, "passed": passed, "error": None})
                except Exception as exc:
                    results.append({"input": document, "report": None, "passed": False, "error": str(exc)})
            self._events.put(("complete", (results, output_root)))
        except Exception as exc:
            self._events.put(("fatal", (str(exc), traceback.format_exc())))

    def _check_events(self) -> None:
        try:
            while True:
                event, payload = self._events.get_nowait()
                if event == "progress":
                    index, total, document = payload  # type: ignore[misc]
                    self.progress.configure(value=index - 1, maximum=total)
                    self.status.set(f"Đang xử lý tệp {index}/{total}: {document.name}")
                elif event == "complete":
                    results, output_root = payload  # type: ignore[misc]
                    self.progress.configure(value=len(results), maximum=max(len(results), 1))
                    self.convert_button.configure(state="normal")
                    if self.open_folder_after.get():
                        os.startfile(output_root)  # type: ignore[attr-defined]
                    successful = sum(result["error"] is None and result["passed"] for result in results)
                    review = sum(result["error"] is None and not result["passed"] for result in results)
                    failed = [result for result in results if result["error"] is not None]
                    self.status.set(
                        f"Hoàn tất: {successful} đạt kiểm tra, {review} cần kiểm tra, {len(failed)} không thể chuyển."
                    )
                    details = f"Đã lưu kết quả tại:\n{output_root}\n\nĐạt kiểm tra: {successful}\nCần kiểm tra: {review}\nLỗi: {len(failed)}"
                    if failed:
                        names = ", ".join(result["input"].name for result in failed[:4])
                        messagebox.showwarning(APP_TITLE, details + f"\n\nKhông thể chuyển: {names}")
                    elif review:
                        messagebox.showwarning(APP_TITLE, details + "\n\nHãy mở báo cáo để kiểm tra công thức, sơ đồ và ký tự đặc biệt.")
                    else:
                        messagebox.showinfo(APP_TITLE, details)
                else:
                    message, details = payload  # type: ignore[misc]
                    self.convert_button.configure(state="normal")
                    self.status.set("Không thể hoàn tất chuyển đổi. Các tệp gốc không bị thay đổi.")
                    messagebox.showerror(APP_TITLE, f"Không thể hoàn tất chuyển đổi:\n\n{message}")
                    print(details)
        except queue.Empty:
            pass
        self.after(150, self._check_events)


if __name__ == "__main__":
    VietnameseTeacherConverterApp().mainloop()
