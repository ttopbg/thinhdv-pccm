"""
convert_teachers_local.py  –  Phiên bản chạy local (tkinter)
Yêu cầu: pip install openpyxl pandas anthropic
Chạy:    python convert_teachers_local.py
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import anthropic

from teacher_core import process_data

NIEN_KHOA_OPTIONS = ["2025-2026", "2026-2027", "2027-2028"]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Chuyển đổi dữ liệu giáo viên")
        self.resizable(False, False)
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        # ── File đầu vào ──────────────────────────────────────────────
        frm_in = tk.LabelFrame(self, text="File đầu vào", **pad)
        frm_in.pack(fill="x", **pad)

        self.var_input = tk.StringVar()
        tk.Entry(frm_in, textvariable=self.var_input, width=55,
                 state="readonly").pack(side="left", padx=(6, 4), pady=6)
        tk.Button(frm_in, text="Chọn...", command=self._pick_input).pack(side="left")

        # ── File đầu ra ───────────────────────────────────────────────
        frm_out = tk.LabelFrame(self, text="File đầu ra", **pad)
        frm_out.pack(fill="x", **pad)

        self.var_output = tk.StringVar()
        tk.Entry(frm_out, textvariable=self.var_output, width=55,
                 state="readonly").pack(side="left", padx=(6, 4), pady=6)
        tk.Button(frm_out, text="Lưu tại...", command=self._pick_output).pack(side="left")

        # ── Niên khóa ─────────────────────────────────────────────────
        frm_nk = tk.LabelFrame(self, text="Niên khóa", **pad)
        frm_nk.pack(fill="x", **pad)

        self.var_nk = tk.StringVar(value=NIEN_KHOA_OPTIONS[0])
        cb = ttk.Combobox(frm_nk, textvariable=self.var_nk,
                          values=NIEN_KHOA_OPTIONS, state="readonly", width=14)
        cb.pack(padx=6, pady=6, anchor="w")

        # ── Log ───────────────────────────────────────────────────────
        frm_log = tk.LabelFrame(self, text="Tiến trình", **pad)
        frm_log.pack(fill="both", expand=True, **pad)

        self.log_box = tk.Text(frm_log, height=10, width=70, state="disabled",
                               font=("Consolas", 9))
        scroll = tk.Scrollbar(frm_log, command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=scroll.set)
        self.log_box.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        scroll.pack(side="left", fill="y", pady=6)

        # ── Nút chạy ─────────────────────────────────────────────────
        self.btn_run = tk.Button(self, text="▶  Chạy chuyển đổi",
                                 command=self._run, bg="#1F4E79", fg="white",
                                 font=("Arial", 11, "bold"), pady=6)
        self.btn_run.pack(fill="x", padx=12, pady=(4, 12))

    # ── Helpers ───────────────────────────────────────────────────────

    def _pick_input(self):
        path = filedialog.askopenfilename(
            title="Chọn file Excel đầu vào",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")]
        )
        if path:
            self.var_input.set(path)

    def _pick_output(self):
        path = filedialog.asksaveasfilename(
            title="Lưu file Excel đầu ra",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.var_output.set(path)

    def _log(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def _run(self):
        input_path  = self.var_input.get().strip()
        output_path = self.var_output.get().strip()
        nien_khoa   = self.var_nk.get()

        if not input_path:
            messagebox.showerror("Lỗi", "Vui lòng chọn file đầu vào!")
            return
        if not output_path:
            messagebox.showerror("Lỗi", "Vui lòng chọn nơi lưu file đầu ra!")
            return

        self.btn_run.configure(state="disabled", text="⏳  Đang xử lý...")
        threading.Thread(target=self._worker,
                         args=(input_path, output_path, nien_khoa),
                         daemon=True).start()

    def _worker(self, input_path, output_path, nien_khoa):
        try:
            client = anthropic.Anthropic()
            result_bytes = process_data(
                input_path, nien_khoa, client,
                progress_cb=lambda m: self.after(0, self._log, m)
            )
            with open(output_path, "wb") as f:
                f.write(result_bytes)
            self.after(0, self._done_ok, output_path)
        except Exception as e:
            self.after(0, self._done_err, str(e))

    def _done_ok(self, output_path):
        self.btn_run.configure(state="normal", text="▶  Chạy chuyển đổi")
        messagebox.showinfo("Hoàn thành", f"✅ Chuyển đổi thành công!\nFile đã lưu:\n{output_path}")

    def _done_err(self, msg):
        self.btn_run.configure(state="normal", text="▶  Chạy chuyển đổi")
        self._log(f"❌ Lỗi: {msg}")
        messagebox.showerror("Lỗi", msg)


if __name__ == "__main__":
    App().mainloop()
