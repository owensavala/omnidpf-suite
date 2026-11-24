"""
Modern Desktop GUI for PDF to Image Converter
Dibuat dengan Tkinter & TTK.
"""

import os
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from converter import convert_pdf_to_images
except ImportError:
    from .converter import convert_pdf_to_images


class PDFToImageApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("PDF to Image Converter")
        self.geometry("640x660")
        self.minsize(550, 600)

        # Style configuration
        self.style = ttk.Style(self)
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.configure(bg="#F3F4F6")

        # Variables
        self.input_paths = []
        self.output_dir_var = tk.StringVar(value="")
        self.format_var = tk.StringVar(value="png")
        self.dpi_var = tk.StringVar(value="200")
        self.pages_var = tk.StringVar(value="all")
        self.status_var = tk.StringVar(value="Siap. Silakan pilih file PDF.")
        self.is_converting = False
        self.last_output_dir = None

        self._build_ui()

    def _build_ui(self):
        main_frame = tk.Frame(self, bg="#F3F4F6", padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header
        header_frame = tk.Frame(main_frame, bg="#F3F4F6")
        header_frame.pack(fill=tk.X, pady=(0, 15))

        title_lbl = tk.Label(
            header_frame,
            text="📄 PDF to Image Converter",
            font=("Segoe UI", 16, "bold"),
            bg="#F3F4F6",
            fg="#1F2937",
        )
        title_lbl.pack(anchor="w")

        subtitle_lbl = tk.Label(
            header_frame,
            text="Ubah dokumen PDF ke PNG, JPG, atau WEBP dengan resolusi tinggi.",
            font=("Segoe UI", 9),
            bg="#F3F4F6",
            fg="#4B5563",
        )
        subtitle_lbl.pack(anchor="w")

        # Section: File Input
        input_box = tk.LabelFrame(
            main_frame,
            text=" 1. Pilih Dokumen PDF ",
            font=("Segoe UI", 10, "bold"),
            bg="#FFFFFF",
            fg="#1F2937",
            padx=12,
            pady=10,
            relief=tk.GROOVE,
        )
        input_box.pack(fill=tk.X, pady=(0, 10))

        btn_row = tk.Frame(input_box, bg="#FFFFFF")
        btn_row.pack(fill=tk.X, pady=(0, 5))

        btn_add_file = tk.Button(
            btn_row,
            text="➕ Pilih File PDF",
            command=self._select_files,
            bg="#2563EB",
            fg="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            relief=tk.FLAT,
            padx=10,
            pady=4,
            cursor="hand2",
        )
        btn_add_file.pack(side=tk.LEFT, padx=(0, 8))

        btn_clear = tk.Button(
            btn_row,
            text="Hapus Pilihan",
            command=self._clear_files,
            bg="#E5E7EB",
            fg="#374151",
            font=("Segoe UI", 9),
            relief=tk.FLAT,
            padx=8,
            pady=4,
            cursor="hand2",
        )
        btn_clear.pack(side=tk.LEFT)

        self.file_listbox = tk.Listbox(
            input_box,
            height=4,
            font=("Segoe UI", 9),
            bg="#F9FAFB",
            selectbackground="#93C5FD",
            selectforeground="#1E3A8A",
            relief=tk.SOLID,
            bd=1,
        )
        self.file_listbox.pack(fill=tk.X, pady=(5, 0))

        # Section: Output Options
        opts_box = tk.LabelFrame(
            main_frame,
            text=" 2. Pengaturan Konversi ",
            font=("Segoe UI", 10, "bold"),
            bg="#FFFFFF",
            fg="#1F2937",
            padx=12,
            pady=10,
            relief=tk.GROOVE,
        )
        opts_box.pack(fill=tk.X, pady=(0, 10))

        # Output folder row
        out_row = tk.Frame(opts_box, bg="#FFFFFF")
        out_row.pack(fill=tk.X, pady=(0, 8))

        tk.Label(
            out_row,
            text="Folder Output:",
            bg="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            width=14,
            anchor="w",
        ).pack(side=tk.LEFT)

        out_entry = tk.Entry(
            out_row,
            textvariable=self.output_dir_var,
            font=("Segoe UI", 9),
            relief=tk.SOLID,
            bd=1,
        )
        out_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        btn_browse_out = tk.Button(
            out_row,
            text="Pilih...",
            command=self._select_output_dir,
            bg="#F3F4F6",
            fg="#1F2937",
            font=("Segoe UI", 8),
            relief=tk.SOLID,
            bd=1,
            cursor="hand2",
        )
        btn_browse_out.pack(side=tk.LEFT)

        # Format and DPI row
        fmt_dpi_row = tk.Frame(opts_box, bg="#FFFFFF")
        fmt_dpi_row.pack(fill=tk.X, pady=(0, 8))

        tk.Label(
            fmt_dpi_row,
            text="Format Gambar:",
            bg="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            width=14,
            anchor="w",
        ).pack(side=tk.LEFT)

        for fmt in [("PNG (Transparan/Tajam)", "png"), ("JPG (Ukuran Kecil)", "jpg"), ("WEBP (Modern)", "webp")]:
            rb = tk.Radiobutton(
                fmt_dpi_row,
                text=fmt[0],
                value=fmt[1],
                variable=self.format_var,
                bg="#FFFFFF",
                font=("Segoe UI", 9),
                activebackground="#FFFFFF",
            )
            rb.pack(side=tk.LEFT, padx=(0, 10))

        # DPI row
        dpi_row = tk.Frame(opts_box, bg="#FFFFFF")
        dpi_row.pack(fill=tk.X, pady=(0, 8))

        tk.Label(
            dpi_row,
            text="Kualitas / DPI:",
            bg="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            width=14,
            anchor="w",
        ).pack(side=tk.LEFT)

        dpi_combo = ttk.Combobox(
            dpi_row,
            textvariable=self.dpi_var,
            values=["150 (Web / Standar)", "200 (Bagus - Rekomendasi)", "300 (Sangat Tajam / Cetak)", "600 (Ultra High)"],
            state="readonly",
            width=26,
            font=("Segoe UI", 9),
        )
        dpi_combo.current(1)
        dpi_combo.pack(side=tk.LEFT)

        # Pages row
        page_row = tk.Frame(opts_box, bg="#FFFFFF")
        page_row.pack(fill=tk.X)

        tk.Label(
            page_row,
            text="Pilihan Halaman:",
            bg="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            width=14,
            anchor="w",
        ).pack(side=tk.LEFT)

        page_entry = tk.Entry(
            page_row,
            textvariable=self.pages_var,
            font=("Segoe UI", 9),
            relief=tk.SOLID,
            bd=1,
            width=20,
        )
        page_entry.pack(side=tk.LEFT, padx=(0, 10))

        tk.Label(
            page_row,
            text="Contoh: 'all' atau '1-3, 5, 7'",
            bg="#FFFFFF",
            fg="#6B7280",
            font=("Segoe UI", 8, "italic"),
        ).pack(side=tk.LEFT)

        # Progress & Status
        status_box = tk.Frame(main_frame, bg="#F3F4F6")
        status_box.pack(fill=tk.X, pady=(5, 10))

        self.progress_bar = ttk.Progressbar(status_box, mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))

        self.status_lbl = tk.Label(
            status_box,
            textvariable=self.status_var,
            font=("Segoe UI", 9),
            bg="#F3F4F6",
            fg="#374151",
            anchor="w",
        )
        self.status_lbl.pack(fill=tk.X)

        # Action Buttons
        action_row = tk.Frame(main_frame, bg="#F3F4F6")
        action_row.pack(fill=tk.X, pady=(5, 5))

        self.btn_convert = tk.Button(
            action_row,
            text="🚀 Mulai Konversi Sekarang",
            command=self._start_conversion_thread,
            bg="#059669",
            fg="#FFFFFF",
            font=("Segoe UI", 11, "bold"),
            relief=tk.FLAT,
            pady=8,
            cursor="hand2",
        )
        self.btn_convert.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        self.btn_open_folder = tk.Button(
            action_row,
            text="📂 Buka Folder Hasil",
            command=self._open_output_folder,
            bg="#4B5563",
            fg="#FFFFFF",
            font=("Segoe UI", 10),
            relief=tk.FLAT,
            pady=8,
            cursor="hand2",
            state=tk.DISABLED,
        )
        self.btn_open_folder.pack(side=tk.RIGHT)

    def _select_files(self):
        files = filedialog.askopenfilenames(
            title="Pilih File PDF",
            filetypes=[("PDF Documents", "*.pdf"), ("Semua File", "*.*")],
        )
        if files:
            for f in files:
                if f not in self.input_paths:
                    self.input_paths.append(f)
                    self.file_listbox.insert(tk.END, Path(f).name)
            self.status_var.set(f"{len(self.input_paths)} file PDF dipilih.")

    def _clear_files(self):
        self.input_paths.clear()
        self.file_listbox.delete(0, tk.END)
        self.status_var.set("Pilihan file dibersihkan.")

    def _select_output_dir(self):
        d = filedialog.askdirectory(title="Pilih Folder Output")
        if d:
            self.output_dir_var.set(d)

    def _start_conversion_thread(self):
        if self.is_converting:
            return

        if not self.input_paths:
            messagebox.showwarning("Perhatian", "Silakan pilih minimal satu file PDF terlebih dahulu!")
            return

        self.is_converting = True
        self.btn_convert.config(state=tk.DISABLED, bg="#9CA3AF")
        self.btn_open_folder.config(state=tk.DISABLED)
        self.progress_bar["value"] = 0

        # Parse DPI
        dpi_str = self.dpi_var.get().split()[0]
        try:
            dpi = int(dpi_str)
        except ValueError:
            dpi = 200

        fmt = self.format_var.get()
        page_range = self.pages_var.get().strip()
        custom_out = self.output_dir_var.get().strip() or None

        threading.Thread(
            target=self._run_conversion,
            args=(self.input_paths.copy(), custom_out, fmt, dpi, page_range),
            daemon=True,
        ).start()

    def _run_conversion(self, files, out_dir, fmt, dpi, page_range):
        total_files = len(files)
        total_created = 0
        last_dir = None

        try:
            for f_idx, pdf_file in enumerate(files, start=1):
                pdf_name = Path(pdf_file).name
                self.status_var.set(f"[{f_idx}/{total_files}] Membaca {pdf_name}...")

                def on_progress(cur, total, fname):
                    percent = int((cur / total) * 100)
                    self.progress_bar["value"] = percent
                    self.status_var.set(f"[{f_idx}/{total_files}] Konversi {cur}/{total}: {fname}")

                saved = convert_pdf_to_images(
                    pdf_path=pdf_file,
                    output_dir=out_dir,
                    image_format=fmt,
                    dpi=dpi,
                    page_range=page_range,
                    progress_callback=on_progress,
                )

                if saved:
                    total_created += len(saved)
                    last_dir = Path(saved[0]).parent

            self.last_output_dir = last_dir or out_dir
            self.status_var.set(f"Selesai! Berhasil membuat {total_created} gambar.")
            self.progress_bar["value"] = 100
            if self.last_output_dir and Path(self.last_output_dir).exists():
                self.btn_open_folder.config(state=tk.NORMAL)

            messagebox.showinfo("Berhasil", f"Konversi selesai!\nTotal {total_created} gambar berhasil dibuat.")
        except Exception as e:
            self.status_var.set(f"Terjadi error: {e}")
            messagebox.showerror("Error Konversi", f"Terjadi kesalahan:\n{e}")
        finally:
            self.is_converting = False
            self.btn_convert.config(state=tk.NORMAL, bg="#059669")

    def _open_output_folder(self):
        target = self.last_output_dir or self.output_dir_var.get().strip()
        if target and os.path.exists(target):
            if sys.platform == "win32":
                os.startfile(target)
            elif sys.platform == "darwin":
                os.system(f'open "{target}"')
            else:
                os.system(f'xdg-open "{target}"')


def main():
    app = PDFToImageApp()
    app.mainloop()


if __name__ == "__main__":
    main()
