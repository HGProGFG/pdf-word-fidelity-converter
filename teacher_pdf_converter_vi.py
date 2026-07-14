#!/usr/bin/env python3
"""Giao diện tiếng Việt cho ứng dụng chuyển PDF sang Word có thể chỉnh sửa."""

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


APP_TITLE = "Chuyển PDF sang Word cho giáo viên"


class VietnameseTeacherConverterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.minsize(720, 460)
        self.resizable(True, False)
        self.pdf_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.status = tk.StringVar(value="Hãy chọn tệp PDF cần chuyển thành Word.")
        self.open_folder_after = tk.BooleanVar(value=True)
        self.keep_all_diffs = tk.BooleanVar(value=False)
        self._events: queue.Queue[tuple[str, object]] = queue.Queue()
        self._build()
        self.after(150, self._check_events)

    def _build(self) -> None:
        container = ttk.Frame(self, padding=24)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)

        ttk.Label(container, text="Chuyển PDF sang Word có thể chỉnh sửa", font=("Segoe UI", 18, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w"
        )
        ttk.Label(
            container,
            text=(
                "Ứng dụng tạo tệp Word có thể chỉnh sửa, xuất lại thành PDF và tự động kiểm tra "
                "mức độ giữ nguyên bố cục của tài liệu."
            ),
            wraplength=650,
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 22))

        ttk.Label(container, text="Tệp PDF").grid(row=2, column=0, sticky="w", padx=(0, 12))
        ttk.Entry(container, textvariable=self.pdf_path).grid(row=2, column=1, sticky="ew")
        ttk.Button(container, text="Chọn PDF…", command=self._choose_pdf).grid(row=2, column=2, padx=(10, 0))

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
        ttk.Checkbutton(options, text="Lưu ảnh so sánh của mọi trang", variable=self.keep_all_diffs).pack(side="left")

        self.progress = ttk.Progressbar(container, mode="indeterminate")
        self.progress.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(24, 8))
        ttk.Label(container, textvariable=self.status, wraplength=650).grid(
            row=6, column=0, columnspan=3, sticky="w"
        )

        self.convert_button = ttk.Button(
            container,
            text="Chuyển PDF thành Word có thể chỉnh sửa",
            command=self._start_conversion,
        )
        self.convert_button.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(24, 0), ipady=7)

        ttk.Separator(container).grid(row=8, column=0, columnspan=3, sticky="ew", pady=18)
        ttk.Label(
            container,
            text=(
                "Mẹo: Mở tệp Word rồi bấm vào công thức. Nếu công thức có thể chỉnh sửa, Word sẽ hiện "
                "công cụ Phương trình. Nếu ứng dụng báo “cần kiểm tra”, hãy xem báo cáo và các ảnh so sánh."
            ),
            wraplength=650,
            foreground="#555555",
        ).grid(row=9, column=0, columnspan=3, sticky="w")

    def _choose_pdf(self) -> None:
        selected = filedialog.askopenfilename(
            title="Chọn tệp PDF cần chuyển", filetypes=[("Tệp PDF", "*.pdf"), ("Tất cả tệp", "*.*")]
        )
        if not selected:
            return
        pdf = Path(selected)
        self.pdf_path.set(str(pdf))
        self.output_dir.set(str(pdf.parent / f"{pdf.stem}-Word-da-chinh-sua"))
        self.status.set("Sẵn sàng chuyển đổi. Tệp PDF gốc sẽ không bị thay đổi.")

    def _choose_output(self) -> None:
        selected = filedialog.askdirectory(title="Chọn nơi lưu các tệp đã chuyển đổi")
        if selected:
            self.output_dir.set(selected)

    def _start_conversion(self) -> None:
        input_pdf = Path(self.pdf_path.get()).expanduser()
        output_text = self.output_dir.get().strip()
        output_dir = Path(output_text).expanduser() if output_text else None
        if not input_pdf.is_file() or input_pdf.suffix.lower() != ".pdf":
            messagebox.showerror(APP_TITLE, "Vui lòng chọn một tệp PDF có sẵn trước.")
            return
        if output_dir is None:
            messagebox.showerror(APP_TITLE, "Vui lòng chọn thư mục để lưu kết quả.")
            return
        diagnostics = converter.environment_diagnostics()
        if not diagnostics["ready"]:
            missing = []
            if not diagnostics["microsoft_word_detected"]:
                missing.append("Microsoft Word bản cài đặt trên máy")
            for module, installed in diagnostics["dependencies"].items():
                if not installed:
                    missing.append(module)
            messagebox.showerror(
                APP_TITLE,
                "Máy tính chưa sẵn sàng. Thiếu: "
                + ", ".join(missing)
                + "\n\nHãy nhờ người cài đặt ứng dụng chạy phần thiết lập trước.",
            )
            return

        self.convert_button.configure(state="disabled")
        self.progress.start(12)
        self.status.set("Microsoft Word đang xử lý. PDF lớn hoặc PDF quét có thể mất vài phút…")
        options = Namespace(
            input=input_pdf,
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
                    output_dir = Path(report["editable_docx"]).parent
                    if self.open_folder_after.get():
                        os.startfile(output_dir)  # type: ignore[attr-defined]
                    if passed:
                        self.status.set("Hoàn tất - kiểm tra bố cục và văn bản đã đạt yêu cầu.")
                        messagebox.showinfo(
                            APP_TITLE,
                            "Đã hoàn tất. Tệp Word và PDF kiểm tra lại đã vượt qua kiểm tra tự động.\n\n"
                            f"Đã lưu tại:\n{output_dir}",
                        )
                    else:
                        self.status.set("Hoàn tất - đã tạo tệp nhưng cần kiểm tra nhanh trước khi chia sẻ.")
                        messagebox.showwarning(
                            APP_TITLE,
                            "Đã hoàn tất, nhưng vui lòng kiểm tra kết quả trước khi chia sẻ.\n\n"
                            "Hãy mở báo cáo kiểm tra và ảnh so sánh trong thư mục kết quả. "
                            "Đặc biệt kiểm tra công thức, sơ đồ và bảng biểu.\n\n"
                            f"Đã lưu tại:\n{output_dir}",
                        )
                else:
                    message, details = payload  # type: ignore[misc]
                    self.status.set("Không thể hoàn tất chuyển đổi. Tệp PDF gốc không bị thay đổi.")
                    messagebox.showerror(
                        APP_TITLE,
                        f"Không thể hoàn tất chuyển đổi:\n\n{message}\n\n"
                        "Nếu PDF được đặt mật khẩu hoặc là bản quét, có thể cần tệp nguồn khác hoặc OCR."
                    )
                    print(details)
        except queue.Empty:
            pass
        self.after(150, self._check_events)


if __name__ == "__main__":
    VietnameseTeacherConverterApp().mainloop()
