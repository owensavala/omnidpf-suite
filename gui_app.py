"""
OmniConverter - Universal Desktop GUI
Sistem Konversi Lengkap Serba Bisa untuk Dokumen, PDF, Gambar, dan Data.
"""

import os
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Import modules
from modules.img_to_pdf import convert_images_to_pdf
from modules.pdf_to_img import convert_pdf_to_images
from modules.pdf_ops import merge_pdfs, split_pdf, extract_text_from_pdf, extract_embedded_images
from modules.img_ops import batch_convert_images
from modules.data_docs import convert_tabular_data, convert_text_to_docx, convert_markdown_to_pdf, convert_docx_to_text


def open_in_explorer(path_str: str):
    """Membuka folder atau memilih file di Windows File Explorer."""
    p = Path(path_str).resolve()
    target = p if p.is_dir() else p.parent
    if target.exists():
        if sys.platform == "win32":
            os.startfile(str(target))
        elif sys.platform == "darwin":
            os.system(f'open "{target}"')
        else:
            os.system(f'xdg-open "{target}"')


class OmniConverterGUI(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("OmniConverter - Sistem Konversi Lengkap Serba Bisa")
        self.geometry("740x720")
        self.minsize(680, 650)

        # Style configuration
        self.style = ttk.Style(self)
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.configure(bg="#F3F4F6")

        # Custom styles
        self.style.configure("TNotebook", background="#E5E7EB", padding=2)
        self.style.configure("TNotebook.Tab", font=("Segoe UI", 9, "bold"), padding=[12, 6])
        self.style.map("TNotebook.Tab", background=[("selected", "#FFFFFF")], foreground=[("selected", "#1E40AF")])

        self._build_header()
        self._build_tabs()

    def _build_header(self):
        header_frame = tk.Frame(self, bg="#1E3A8A", padx=20, pady=12)
        header_frame.pack(fill=tk.X)

        title = tk.Label(
            header_frame,
            text="⚡ OmniConverter",
            font=("Segoe UI", 16, "bold"),
            bg="#1E3A8A",
            fg="#FFFFFF",
        )
        title.pack(anchor="w")

        subtitle = tk.Label(
            header_frame,
            text="Alat Konversi Lengkap: Gambar ke PDF (Multi-file Berurutan), PDF ke Gambar, Pengolah PDF, Format Gambar & Data",
            font=("Segoe UI", 9),
            bg="#1E3A8A",
            fg="#BFDBFE",
        )
        subtitle.pack(anchor="w")

    def _build_tabs(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Tab 1: Gambar ke PDF (Prioritas Utama User)
        self.tab_img2pdf = tk.Frame(self.notebook, bg="#F9FAFB", padx=15, pady=12)
        self.notebook.add(self.tab_img2pdf, text=" 🖼️ JPG/Gambar ke PDF ")
        self._setup_tab_img_to_pdf()

        # Tab 2: PDF ke Gambar
        self.tab_pdf2img = tk.Frame(self.notebook, bg="#F9FAFB", padx=15, pady=12)
        self.notebook.add(self.tab_pdf2img, text=" 📄 PDF ke Gambar ")
        self._setup_tab_pdf_to_img()

        # Tab 3: Alat & Operasi PDF
        self.tab_pdfops = tk.Frame(self.notebook, bg="#F9FAFB", padx=15, pady=12)
        self.notebook.add(self.tab_pdfops, text=" 📑 Pengolah PDF ")
        self._setup_tab_pdf_ops()

        # Tab 4: Ubah Format Gambar
        self.tab_imgconv = tk.Frame(self.notebook, bg="#F9FAFB", padx=15, pady=12)
        self.notebook.add(self.tab_imgconv, text=" 🎨 Format Gambar ")
        self._setup_tab_img_conv()

        # Tab 5: Data & Dokumen
        self.tab_datadocs = tk.Frame(self.notebook, bg="#F9FAFB", padx=15, pady=12)
        self.notebook.add(self.tab_datadocs, text=" 📊 Data & Dokumen ")
        self._setup_tab_data_docs()

    # =========================================================================
    # TAB 1: GAMBAR KE PDF (MULTI-FILE BERURUTAN)
    # =========================================================================
    def _setup_tab_img_to_pdf(self):
        frame = self.tab_img2pdf

        desc_lbl = tk.Label(
            frame,
            text="Gabungkan satu atau banyak gambar (JPG, PNG, WEBP, dll.) menjadi satu dokumen PDF berurutan.\nGunakan tombol ⬆️ Naik dan ⬇️ Turun untuk menyesuaikan urutan halaman.",
            font=("Segoe UI", 9),
            bg="#F9FAFB",
            fg="#4B5563",
            justify="left",
        )
        desc_lbl.pack(anchor="w", pady=(0, 8))

        # Kontainer Daftar File & Tombol Urutan
        list_container = tk.Frame(frame, bg="#F9FAFB")
        list_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Listbox dengan Scrollbar
        list_sub = tk.Frame(list_container, bg="#F9FAFB")
        list_sub.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.img_listbox = tk.Listbox(
            list_sub,
            font=("Segoe UI", 9),
            selectmode=tk.SINGLE,
            relief=tk.SOLID,
            bd=1,
            bg="#FFFFFF",
        )
        self.img_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(list_sub, orient="vertical", command=self.img_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.img_listbox.config(yscrollcommand=scrollbar.set)

        # Tombol aksi samping (Urutan & Kelola)
        btn_side = tk.Frame(list_container, bg="#F9FAFB", padx=8)
        btn_side.pack(side=tk.RIGHT, fill=tk.Y)

        tk.Button(
            btn_side, text="➕ Tambah Gambar", command=self._i2p_add_files,
            bg="#2563EB", fg="#FFF", font=("Segoe UI", 9, "bold"), relief=tk.FLAT, pady=4, cursor="hand2"
        ).pack(fill=tk.X, pady=(0, 4))

        tk.Button(
            btn_side, text="📁 Tambah Folder", command=self._i2p_add_folder,
            bg="#3B82F6", fg="#FFF", font=("Segoe UI", 9), relief=tk.FLAT, pady=3, cursor="hand2"
        ).pack(fill=tk.X, pady=(0, 8))

        tk.Button(
            btn_side, text="⬆️ Geser Naik", command=self._i2p_move_up,
            bg="#E5E7EB", fg="#1F2937", font=("Segoe UI", 9), relief=tk.SOLID, bd=1, pady=3, cursor="hand2"
        ).pack(fill=tk.X, pady=(0, 4))

        tk.Button(
            btn_side, text="⬇️ Geser Turun", command=self._i2p_move_down,
            bg="#E5E7EB", fg="#1F2937", font=("Segoe UI", 9), relief=tk.SOLID, bd=1, pady=3, cursor="hand2"
        ).pack(fill=tk.X, pady=(0, 8))

        tk.Button(
            btn_side, text="🔤 Urutkan A-Z", command=self._i2p_sort_alpha,
            bg="#F3F4F6", fg="#374151", font=("Segoe UI", 8), relief=tk.SOLID, bd=1, pady=2, cursor="hand2"
        ).pack(fill=tk.X, pady=(0, 4))

        tk.Button(
            btn_side, text="🗑️ Hapus Terpilih", command=self._i2p_remove_selected,
            bg="#FEE2E2", fg="#B91C1C", font=("Segoe UI", 9), relief=tk.SOLID, bd=1, pady=3, cursor="hand2"
        ).pack(fill=tk.X, pady=(0, 4))

        tk.Button(
            btn_side, text="🧹 Hapus Semua", command=self._i2p_clear_all,
            bg="#F3F4F6", fg="#6B7280", font=("Segoe UI", 8), relief=tk.FLAT, pady=2, cursor="hand2"
        ).pack(fill=tk.X)

        self.i2p_file_paths = []

        # Kotak Pengaturan
        opt_box = tk.LabelFrame(frame, text=" Pengaturan Output PDF ", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", padx=10, pady=8)
        opt_box.pack(fill=tk.X, pady=(0, 8))

        # Row 1: File Output Name & Browse
        row1 = tk.Frame(opt_box, bg="#FFFFFF")
        row1.pack(fill=tk.X, pady=(0, 6))

        tk.Label(row1, text="Simpan Ke:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", width=12, anchor="w").pack(side=tk.LEFT)
        self.i2p_out_var = tk.StringVar(value="")
        tk.Entry(row1, textvariable=self.i2p_out_var, font=("Segoe UI", 9), relief=tk.SOLID, bd=1).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        tk.Button(row1, text="Browse...", command=self._i2p_browse_out, relief=tk.SOLID, bd=1, font=("Segoe UI", 8)).pack(side=tk.LEFT)

        # Row 2: Page size, Margin, Orientation
        row2 = tk.Frame(opt_box, bg="#FFFFFF")
        row2.pack(fill=tk.X)

        tk.Label(row2, text="Ukuran Halaman:", font=("Segoe UI", 9), bg="#FFFFFF").pack(side=tk.LEFT, padx=(0, 4))
        self.i2p_size_var = tk.StringVar(value="A4")
        ttk.Combobox(row2, textvariable=self.i2p_size_var, values=["A4", "Original", "Letter"], width=10, state="readonly").pack(side=tk.LEFT, padx=(0, 15))

        tk.Label(row2, text="Margin:", font=("Segoe UI", 9), bg="#FFFFFF").pack(side=tk.LEFT, padx=(0, 4))
        self.i2p_margin_var = tk.StringVar(value="Normal (20pt)")
        ttk.Combobox(row2, textvariable=self.i2p_margin_var, values=["Tanpa Margin (0pt)", "Kecil (10pt)", "Normal (20pt)", "Lebar (40pt)"], width=16, state="readonly").pack(side=tk.LEFT, padx=(0, 15))

        self.i2p_autorotate_var = tk.BooleanVar(value=True)
        tk.Checkbutton(row2, text="Otomatis Sesuaikan Orientasi", variable=self.i2p_autorotate_var, bg="#FFFFFF", font=("Segoe UI", 9)).pack(side=tk.LEFT)

        # Action bar & progress
        bar_frame = tk.Frame(frame, bg="#F9FAFB")
        bar_frame.pack(fill=tk.X, pady=(4, 0))

        self.i2p_progress = ttk.Progressbar(bar_frame, mode="determinate")
        self.i2p_progress.pack(fill=tk.X, pady=(0, 4))

        self.i2p_status_var = tk.StringVar(value="Siap. Silakan tambahkan file gambar.")
        tk.Label(bar_frame, textvariable=self.i2p_status_var, font=("Segoe UI", 9), bg="#F9FAFB", fg="#374151", anchor="w").pack(fill=tk.X)

        act_frame = tk.Frame(frame, bg="#F9FAFB")
        act_frame.pack(fill=tk.X, pady=(6, 0))

        self.i2p_btn_convert = tk.Button(
            act_frame, text="🚀 Gabung & Buat PDF Sekarang", command=self._i2p_execute,
            bg="#059669", fg="#FFFFFF", font=("Segoe UI", 11, "bold"), relief=tk.FLAT, pady=7, cursor="hand2"
        )
        self.i2p_btn_convert.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        self.i2p_btn_open = tk.Button(
            act_frame, text="📂 Buka Folder Hasil", command=self._i2p_open_folder,
            bg="#4B5563", fg="#FFFFFF", font=("Segoe UI", 10), relief=tk.FLAT, pady=7, state=tk.DISABLED, cursor="hand2"
        )
        self.i2p_btn_open.pack(side=tk.RIGHT)

    def _i2p_refresh_listbox(self):
        self.img_listbox.delete(0, tk.END)
        for idx, p in enumerate(self.i2p_file_paths, start=1):
            self.img_listbox.insert(tk.END, f"{idx}. {Path(p).name}")
        self.i2p_status_var.set(f"{len(self.i2p_file_paths)} gambar dalam antrean.")

    def _i2p_add_files(self):
        files = filedialog.askopenfilenames(
            title="Pilih Gambar",
            filetypes=[("Semua Format Gambar", "*.jpg *.jpeg *.png *.webp *.bmp *.tiff"), ("JPG", "*.jpg *.jpeg"), ("PNG", "*.png"), ("Semua File", "*.*")]
        )
        if files:
            for f in files:
                if f not in self.i2p_file_paths:
                    self.i2p_file_paths.append(f)
            self._i2p_refresh_listbox()
            if not self.i2p_out_var.get() and self.i2p_file_paths:
                first = Path(self.i2p_file_paths[0])
                self.i2p_out_var.set(str(first.parent / f"{first.stem}_gabungan.pdf"))

    def _i2p_add_folder(self):
        folder = filedialog.askdirectory(title="Pilih Folder Berisi Gambar")
        if folder:
            valid_exts = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff"}
            found = [str(f) for f in sorted(Path(folder).iterdir()) if f.suffix.lower() in valid_exts]
            for f in found:
                if f not in self.i2p_file_paths:
                    self.i2p_file_paths.append(f)
            self._i2p_refresh_listbox()
            if not self.i2p_out_var.get() and self.i2p_file_paths:
                self.i2p_out_var.set(str(Path(folder) / "hasil_gabungan.pdf"))

    def _i2p_move_up(self):
        sel = self.img_listbox.curselection()
        if not sel or sel[0] == 0:
            return
        idx = sel[0]
        self.i2p_file_paths[idx - 1], self.i2p_file_paths[idx] = self.i2p_file_paths[idx], self.i2p_file_paths[idx - 1]
        self._i2p_refresh_listbox()
        self.img_listbox.selection_set(idx - 1)

    def _i2p_move_down(self):
        sel = self.img_listbox.curselection()
        if not sel or sel[0] == len(self.i2p_file_paths) - 1:
            return
        idx = sel[0]
        self.i2p_file_paths[idx + 1], self.i2p_file_paths[idx] = self.i2p_file_paths[idx], self.i2p_file_paths[idx + 1]
        self._i2p_refresh_listbox()
        self.img_listbox.selection_set(idx + 1)

    def _i2p_sort_alpha(self):
        self.i2p_file_paths.sort(key=lambda x: Path(x).name.lower())
        self._i2p_refresh_listbox()

    def _i2p_remove_selected(self):
        sel = self.img_listbox.curselection()
        if not sel:
            return
        del self.i2p_file_paths[sel[0]]
        self._i2p_refresh_listbox()

    def _i2p_clear_all(self):
        self.i2p_file_paths.clear()
        self._i2p_refresh_listbox()

    def _i2p_browse_out(self):
        p = filedialog.asksaveasfilename(
            title="Simpan File PDF Sebagai",
            defaultextension=".pdf",
            filetypes=[("PDF Document", "*.pdf")]
        )
        if p:
            self.i2p_out_var.set(p)

    def _i2p_execute(self):
        if not self.i2p_file_paths:
            messagebox.showwarning("Perhatian", "Silakan tambahkan minimal satu gambar terlebih dahulu!")
            return

        out_pdf = self.i2p_out_var.get().strip()
        if not out_pdf:
            messagebox.showwarning("Perhatian", "Tentukan lokasi file PDF hasil konversi!")
            return

        margin_map = {
            "Tanpa Margin (0pt)": 0.0,
            "Kecil (10pt)": 10.0,
            "Normal (20pt)": 20.0,
            "Lebar (40pt)": 40.0,
        }
        margin = margin_map.get(self.i2p_margin_var.get(), 20.0)
        page_size = self.i2p_size_var.get()
        auto_rotate = self.i2p_autorotate_var.get()

        self.i2p_btn_convert.config(state=tk.DISABLED, bg="#9CA3AF")
        self.i2p_btn_open.config(state=tk.DISABLED)
        self.i2p_progress["value"] = 0

        def worker():
            try:
                def on_prog(curr, total, fname):
                    self.i2p_progress["value"] = int((curr / total) * 100)
                    self.i2p_status_var.set(f"Menambahkan [{curr}/{total}]: {fname}")

                convert_images_to_pdf(
                    image_paths=self.i2p_file_paths,
                    output_pdf_path=out_pdf,
                    page_size=page_size,
                    margin=margin,
                    auto_rotate=auto_rotate,
                    progress_callback=on_prog,
                )
                self.i2p_status_var.set(f"Selesai! PDF berhasil dibuat: {Path(out_pdf).name}")
                self.i2p_progress["value"] = 100
                self.i2p_btn_open.config(state=tk.NORMAL)
                messagebox.showinfo("Berhasil", f"Berhasil menggabungkan {len(self.i2p_file_paths)} gambar ke:\n{out_pdf}")
            except Exception as e:
                self.i2p_status_var.set(f"Error: {e}")
                messagebox.showerror("Error", f"Terjadi kesalahan saat membuat PDF:\n{e}")
            finally:
                self.i2p_btn_convert.config(state=tk.NORMAL, bg="#059669")

        threading.Thread(target=worker, daemon=True).start()

    def _i2p_open_folder(self):
        out_pdf = self.i2p_out_var.get().strip()
        if out_pdf:
            open_in_explorer(out_pdf)

    # =========================================================================
    # TAB 2: PDF KE GAMBAR
    # =========================================================================
    def _setup_tab_pdf_to_img(self):
        frame = self.tab_pdf2img

        tk.Label(
            frame,
            text="Ekstrak halaman PDF menjadi file gambar (PNG, JPG, atau WEBP) dengan pilihan DPI.",
            font=("Segoe UI", 9), bg="#F9FAFB", fg="#4B5563"
        ).pack(anchor="w", pady=(0, 10))

        box = tk.LabelFrame(frame, text=" Input File PDF & Output ", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", padx=12, pady=10)
        box.pack(fill=tk.X, pady=(0, 10))

        # Input PDF
        r1 = tk.Frame(box, bg="#FFFFFF")
        r1.pack(fill=tk.X, pady=(0, 8))
        tk.Label(r1, text="File PDF:", font=("Segoe UI", 9, "bold"), width=12, bg="#FFFFFF", anchor="w").pack(side=tk.LEFT)
        self.p2i_input_var = tk.StringVar()
        tk.Entry(r1, textvariable=self.p2i_input_var, font=("Segoe UI", 9), relief=tk.SOLID, bd=1).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        tk.Button(r1, text="Pilih PDF...", command=self._p2i_browse_pdf, relief=tk.SOLID, bd=1, font=("Segoe UI", 8)).pack(side=tk.LEFT)

        # Output folder
        r2 = tk.Frame(box, bg="#FFFFFF")
        r2.pack(fill=tk.X, pady=(0, 8))
        tk.Label(r2, text="Folder Hasil:", font=("Segoe UI", 9, "bold"), width=12, bg="#FFFFFF", anchor="w").pack(side=tk.LEFT)
        self.p2i_out_var = tk.StringVar()
        tk.Entry(r2, textvariable=self.p2i_out_var, font=("Segoe UI", 9), relief=tk.SOLID, bd=1).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        tk.Button(r2, text="Pilih Folder...", command=self._p2i_browse_out, relief=tk.SOLID, bd=1, font=("Segoe UI", 8)).pack(side=tk.LEFT)

        # Format & DPI
        r3 = tk.Frame(box, bg="#FFFFFF")
        r3.pack(fill=tk.X, pady=(0, 8))
        tk.Label(r3, text="Format:", font=("Segoe UI", 9), bg="#FFFFFF").pack(side=tk.LEFT, padx=(0, 4))
        self.p2i_fmt_var = tk.StringVar(value="png")
        for fmt in ["png", "jpg", "webp"]:
            tk.Radiobutton(r3, text=fmt.upper(), value=fmt, variable=self.p2i_fmt_var, bg="#FFFFFF").pack(side=tk.LEFT, padx=(0, 10))

        tk.Label(r3, text="Kualitas (DPI):", font=("Segoe UI", 9), bg="#FFFFFF").pack(side=tk.LEFT, padx=(15, 4))
        self.p2i_dpi_var = tk.StringVar(value="200")
        ttk.Combobox(r3, textvariable=self.p2i_dpi_var, values=["150", "200", "300", "600"], width=6, state="readonly").pack(side=tk.LEFT)

        # Page range
        r4 = tk.Frame(box, bg="#FFFFFF")
        r4.pack(fill=tk.X)
        tk.Label(r4, text="Halaman:", font=("Segoe UI", 9, "bold"), width=12, bg="#FFFFFF", anchor="w").pack(side=tk.LEFT)
        self.p2i_pages_var = tk.StringVar(value="all")
        tk.Entry(r4, textvariable=self.p2i_pages_var, font=("Segoe UI", 9), width=18, relief=tk.SOLID, bd=1).pack(side=tk.LEFT, padx=(0, 10))
        tk.Label(r4, text="(Contoh: 'all' atau '1-3, 5')", font=("Segoe UI", 8, "italic"), bg="#FFFFFF", fg="#6B7280").pack(side=tk.LEFT)

        # Progress & Action
        self.p2i_progress = ttk.Progressbar(frame, mode="determinate")
        self.p2i_progress.pack(fill=tk.X, pady=(15, 5))

        self.p2i_status_var = tk.StringVar(value="Siap.")
        tk.Label(frame, textvariable=self.p2i_status_var, font=("Segoe UI", 9), bg="#F9FAFB", fg="#374151", anchor="w").pack(fill=tk.X, pady=(0, 10))

        btn_row = tk.Frame(frame, bg="#F9FAFB")
        btn_row.pack(fill=tk.X)

        self.p2i_btn_run = tk.Button(
            btn_row, text="🚀 Konversi PDF ke Gambar", command=self._p2i_execute,
            bg="#059669", fg="#FFFFFF", font=("Segoe UI", 11, "bold"), relief=tk.FLAT, pady=7, cursor="hand2"
        )
        self.p2i_btn_run.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        self.p2i_btn_open = tk.Button(
            btn_row, text="📂 Buka Folder Hasil", command=self._p2i_open_folder,
            bg="#4B5563", fg="#FFFFFF", font=("Segoe UI", 10), relief=tk.FLAT, pady=7, state=tk.DISABLED, cursor="hand2"
        )
        self.p2i_btn_open.pack(side=tk.RIGHT)

    def _p2i_browse_pdf(self):
        f = filedialog.askopenfilename(title="Pilih File PDF", filetypes=[("PDF Document", "*.pdf")])
        if f:
            self.p2i_input_var.set(f)
            if not self.p2i_out_var.get():
                p = Path(f)
                self.p2i_out_var.set(str(p.parent / f"{p.stem}_images"))

    def _p2i_browse_out(self):
        d = filedialog.askdirectory(title="Pilih Folder Output")
        if d:
            self.p2i_out_var.set(d)

    def _p2i_execute(self):
        pdf_path = self.p2i_input_var.get().strip()
        if not pdf_path or not Path(pdf_path).is_file():
            messagebox.showwarning("Perhatian", "Pilih file PDF yang valid terlebih dahulu!")
            return

        out_dir = self.p2i_out_var.get().strip() or None
        fmt = self.p2i_fmt_var.get()
        dpi = int(self.p2i_dpi_var.get())
        pages = self.p2i_pages_var.get().strip()

        self.p2i_btn_run.config(state=tk.DISABLED, bg="#9CA3AF")
        self.p2i_btn_open.config(state=tk.DISABLED)
        self.p2i_progress["value"] = 0

        def worker():
            try:
                def on_prog(c, t, fname):
                    self.p2i_progress["value"] = int((c / t) * 100)
                    self.p2i_status_var.set(f"Menyimpan [{c}/{t}]: {fname}")

                res = convert_pdf_to_images(
                    pdf_path=pdf_path,
                    output_dir=out_dir,
                    image_format=fmt,
                    dpi=dpi,
                    page_range=pages,
                    progress_callback=on_prog,
                )
                self.p2i_progress["value"] = 100
                self.p2i_status_var.set(f"Selesai! {len(res)} gambar berhasil dibuat.")
                self.p2i_btn_open.config(state=tk.NORMAL)
                messagebox.showinfo("Berhasil", f"Selesai! Berhasil membuat {len(res)} file gambar.")
            except Exception as e:
                self.p2i_status_var.set(f"Error: {e}")
                messagebox.showerror("Error", str(e))
            finally:
                self.p2i_btn_run.config(state=tk.NORMAL, bg="#059669")

        threading.Thread(target=worker, daemon=True).start()

    def _p2i_open_folder(self):
        d = self.p2i_out_var.get().strip()
        if d:
            open_in_explorer(d)

    # =========================================================================
    # TAB 3: ALAT & PENGOLAH PDF (GABUNG, PISAH, EKSTRAK)
    # =========================================================================
    def _setup_tab_pdf_ops(self):
        frame = self.tab_pdfops

        tk.Label(
            frame,
            text="Fitur manajemen PDF: Gabungkan beberapa file PDF, Pisahkan halaman, dan Ekstrak Teks/Gambar asli.",
            font=("Segoe UI", 9), bg="#F9FAFB", fg="#4B5563"
        ).pack(anchor="w", pady=(0, 10))

        # Sub-fitur 1: Gabung Beberapa File PDF
        merge_box = tk.LabelFrame(frame, text=" 1. Gabung Beberapa PDF Menjadi Satu ", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", padx=10, pady=8)
        merge_box.pack(fill=tk.X, pady=(0, 10))

        self.pdf_merge_list = []
        self.merge_listbox = tk.Listbox(merge_box, height=3, font=("Segoe UI", 9), relief=tk.SOLID, bd=1)
        self.merge_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))

        m_btn_frame = tk.Frame(merge_box, bg="#FFFFFF")
        m_btn_frame.pack(side=tk.LEFT, fill=tk.Y)

        tk.Button(m_btn_frame, text="➕ Tambah PDF", command=self._merge_add_pdf, bg="#2563EB", fg="#FFF", font=("Segoe UI", 8, "bold"), relief=tk.FLAT).pack(fill=tk.X, pady=(0, 3))
        tk.Button(m_btn_frame, text="🧹 Bersihkan", command=self._merge_clear, bg="#E5E7EB", font=("Segoe UI", 8), relief=tk.FLAT).pack(fill=tk.X, pady=(0, 3))
        tk.Button(m_btn_frame, text="🚀 Gabung PDF", command=self._merge_execute, bg="#059669", fg="#FFF", font=("Segoe UI", 8, "bold"), relief=tk.FLAT).pack(fill=tk.X)

        # Sub-fitur 2: Pisah PDF (Split)
        split_box = tk.LabelFrame(frame, text=" 2. Pisahkan (Split) Halaman PDF ", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", padx=10, pady=8)
        split_box.pack(fill=tk.X, pady=(0, 10))

        s_r1 = tk.Frame(split_box, bg="#FFFFFF")
        s_r1.pack(fill=tk.X, pady=(0, 6))
        tk.Label(s_r1, text="File PDF:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", width=10, anchor="w").pack(side=tk.LEFT)
        self.split_input_var = tk.StringVar()
        tk.Entry(s_r1, textvariable=self.split_input_var, font=("Segoe UI", 9), relief=tk.SOLID, bd=1).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        tk.Button(s_r1, text="Pilih PDF...", command=lambda: self.split_input_var.set(filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])), relief=tk.SOLID, bd=1, font=("Segoe UI", 8)).pack(side=tk.LEFT)

        s_r2 = tk.Frame(split_box, bg="#FFFFFF")
        s_r2.pack(fill=tk.X)
        self.split_mode_var = tk.StringVar(value="each_page")
        tk.Radiobutton(s_r2, text="Pisah per 1 Halaman", variable=self.split_mode_var, value="each_page", bg="#FFFFFF").pack(side=tk.LEFT, padx=(0, 15))
        tk.Radiobutton(s_r2, text="Berdasarkan Rentang (misal: 1-3, 4-10):", variable=self.split_mode_var, value="ranges", bg="#FFFFFF").pack(side=tk.LEFT)
        self.split_ranges_var = tk.StringVar(value="1-3, 4-5")
        tk.Entry(s_r2, textvariable=self.split_ranges_var, font=("Segoe UI", 9), width=15, relief=tk.SOLID, bd=1).pack(side=tk.LEFT, padx=(6, 10))
        tk.Button(s_r2, text="✂️ Pisah Sekarang", command=self._split_execute, bg="#D97706", fg="#FFF", font=("Segoe UI", 8, "bold"), relief=tk.FLAT).pack(side=tk.RIGHT)

        # Sub-fitur 3: Ekstraksi Teks & Gambar Asli
        ext_box = tk.LabelFrame(frame, text=" 3. Ekstrak Teks & Gambar Asli dari PDF ", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", padx=10, pady=8)
        ext_box.pack(fill=tk.X, pady=(0, 10))

        e_r1 = tk.Frame(ext_box, bg="#FFFFFF")
        e_r1.pack(fill=tk.X, pady=(0, 6))
        tk.Label(e_r1, text="File PDF:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", width=10, anchor="w").pack(side=tk.LEFT)
        self.ext_input_var = tk.StringVar()
        tk.Entry(e_r1, textvariable=self.ext_input_var, font=("Segoe UI", 9), relief=tk.SOLID, bd=1).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        tk.Button(e_r1, text="Pilih PDF...", command=lambda: self.ext_input_var.set(filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])), relief=tk.SOLID, bd=1, font=("Segoe UI", 8)).pack(side=tk.LEFT)

        e_r2 = tk.Frame(ext_box, bg="#FFFFFF")
        e_r2.pack(fill=tk.X)
        tk.Button(e_r2, text="📝 Ekstrak ke File Teks (.txt)", command=self._extract_text_execute, bg="#4B5563", fg="#FFF", font=("Segoe UI", 9), relief=tk.FLAT, padx=10, pady=4).pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(e_r2, text="🖼️ Ekstrak Semua Gambar Tertanam", command=self._extract_images_execute, bg="#4B5563", fg="#FFF", font=("Segoe UI", 9), relief=tk.FLAT, padx=10, pady=4).pack(side=tk.LEFT)

    def _merge_add_pdf(self):
        fs = filedialog.askopenfilenames(title="Pilih File PDF untuk Digabung", filetypes=[("PDF", "*.pdf")])
        if fs:
            for f in fs:
                self.pdf_merge_list.append(f)
                self.merge_listbox.insert(tk.END, Path(f).name)

    def _merge_clear(self):
        self.pdf_merge_list.clear()
        self.merge_listbox.delete(0, tk.END)

    def _merge_execute(self):
        if len(self.pdf_merge_list) < 2:
            messagebox.showwarning("Perhatian", "Pilih minimal 2 file PDF untuk digabungkan!")
            return
        out_p = filedialog.asksaveasfilename(title="Simpan Hasil PDF Gabungan", defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if out_p:
            try:
                merge_pdfs(self.pdf_merge_list, out_p)
                messagebox.showinfo("Berhasil", f"PDF berhasil digabung ke:\n{out_p}")
                open_in_explorer(out_p)
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def _split_execute(self):
        src = self.split_input_var.get().strip()
        if not src or not Path(src).is_file():
            messagebox.showwarning("Perhatian", "Pilih file PDF yang valid terlebih dahulu!")
            return
        out_d = filedialog.askdirectory(title="Pilih Folder Hasil Pemisahan")
        if out_d:
            try:
                mode = self.split_mode_var.get()
                ranges = self.split_ranges_var.get().strip() if mode == "ranges" else None
                res = split_pdf(src, out_d, split_mode=mode, ranges_str=ranges)
                messagebox.showinfo("Berhasil", f"Berhasil memisahkan PDF menjadi {len(res)} file.")
                open_in_explorer(out_d)
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def _extract_text_execute(self):
        src = self.ext_input_var.get().strip()
        if not src or not Path(src).is_file():
            messagebox.showwarning("Perhatian", "Pilih file PDF terlebih dahulu!")
            return
        out_txt = filedialog.asksaveasfilename(title="Simpan Teks Sebagai", defaultextension=".txt", filetypes=[("Text File", "*.txt")])
        if out_txt:
            try:
                extract_text_from_pdf(src, out_txt)
                messagebox.showinfo("Berhasil", f"Teks berhasil diekstrak ke:\n{out_txt}")
                open_in_explorer(out_txt)
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def _extract_images_execute(self):
        src = self.ext_input_var.get().strip()
        if not src or not Path(src).is_file():
            messagebox.showwarning("Perhatian", "Pilih file PDF terlebih dahulu!")
            return
        out_d = filedialog.askdirectory(title="Pilih Folder Hasil Gambar")
        if out_d:
            try:
                res = extract_embedded_images(src, out_d)
                messagebox.showinfo("Berhasil", f"Ditemukan dan diekstrak {len(res)} gambar dari PDF.")
                open_in_explorer(out_d)
            except Exception as e:
                messagebox.showerror("Error", str(e))

    # =========================================================================
    # TAB 4: FORMAT GAMBAR (IMAGE TO IMAGE CONVERTER)
    # =========================================================================
    def _setup_tab_img_conv(self):
        frame = self.tab_imgconv

        tk.Label(
            frame,
            text="Ubah format gambar antar PNG, JPG, WEBP, ICO (icon), BMP, dan TIFF (Mendukung Batch Banyak Gambar).",
            font=("Segoe UI", 9), bg="#F9FAFB", fg="#4B5563"
        ).pack(anchor="w", pady=(0, 10))

        box = tk.LabelFrame(frame, text=" Pengaturan Konversi Gambar ", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", padx=12, pady=10)
        box.pack(fill=tk.X, pady=(0, 10))

        # File picker
        r1 = tk.Frame(box, bg="#FFFFFF")
        r1.pack(fill=tk.X, pady=(0, 8))
        self.imgconv_files = []
        self.imgconv_lbl = tk.Label(r1, text="0 gambar dipilih.", font=("Segoe UI", 9), bg="#FFFFFF")
        self.imgconv_lbl.pack(side=tk.LEFT)
        tk.Button(r1, text="➕ Pilih Gambar...", command=self._imgconv_pick_files, relief=tk.SOLID, bd=1, font=("Segoe UI", 8)).pack(side=tk.RIGHT)

        # Target format
        r2 = tk.Frame(box, bg="#FFFFFF")
        r2.pack(fill=tk.X, pady=(0, 8))
        tk.Label(r2, text="Format Target:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", width=14, anchor="w").pack(side=tk.LEFT)
        self.imgconv_target_fmt = tk.StringVar(value="webp")
        for f in ["png", "jpg", "webp", "ico", "bmp", "tiff"]:
            tk.Radiobutton(r2, text=f.upper(), value=f, variable=self.imgconv_target_fmt, bg="#FFFFFF").pack(side=tk.LEFT, padx=(0, 6))

        # Quality & Resize
        r3 = tk.Frame(box, bg="#FFFFFF")
        r3.pack(fill=tk.X, pady=(0, 8))
        tk.Label(r3, text="Kualitas (JPG/WEBP):", font=("Segoe UI", 9), bg="#FFFFFF").pack(side=tk.LEFT, padx=(0, 4))
        self.imgconv_quality = tk.IntVar(value=90)
        tk.Spinbox(r3, from_=10, to=100, increment=5, textvariable=self.imgconv_quality, width=5).pack(side=tk.LEFT, padx=(0, 15))

        tk.Label(r3, text="Resize (% Ukuran):", font=("Segoe UI", 9), bg="#FFFFFF").pack(side=tk.LEFT, padx=(0, 4))
        self.imgconv_resize = tk.IntVar(value=100)
        tk.Spinbox(r3, from_=10, to=200, increment=10, textvariable=self.imgconv_resize, width=5).pack(side=tk.LEFT)
        tk.Label(r3, text="(100 = ukuran asli)", font=("Segoe UI", 8, "italic"), bg="#FFFFFF", fg="#6B7280").pack(side=tk.LEFT, padx=(4, 0))

        # Output folder
        r4 = tk.Frame(box, bg="#FFFFFF")
        r4.pack(fill=tk.X)
        tk.Label(r4, text="Folder Output:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", width=14, anchor="w").pack(side=tk.LEFT)
        self.imgconv_out_var = tk.StringVar()
        tk.Entry(r4, textvariable=self.imgconv_out_var, font=("Segoe UI", 9), relief=tk.SOLID, bd=1).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        tk.Button(r4, text="Pilih Folder...", command=lambda: self.imgconv_out_var.set(filedialog.askdirectory()), relief=tk.SOLID, bd=1, font=("Segoe UI", 8)).pack(side=tk.LEFT)

        # Action
        tk.Button(
            frame, text="🚀 Konversi Gambar Sekarang", command=self._imgconv_execute,
            bg="#059669", fg="#FFFFFF", font=("Segoe UI", 11, "bold"), relief=tk.FLAT, pady=8, cursor="hand2"
        ).pack(fill=tk.X, pady=(15, 0))

    def _imgconv_pick_files(self):
        fs = filedialog.askopenfilenames(
            title="Pilih Gambar",
            filetypes=[("Semua Gambar", "*.jpg *.jpeg *.png *.webp *.bmp *.tiff"), ("Semua File", "*.*")]
        )
        if fs:
            self.imgconv_files = list(fs)
            self.imgconv_lbl.config(text=f"{len(self.imgconv_files)} gambar dipilih.")
            if not self.imgconv_out_var.get():
                self.imgconv_out_var.set(str(Path(self.imgconv_files[0]).parent / "hasil_konversi_gambar"))

    def _imgconv_execute(self):
        if not self.imgconv_files:
            messagebox.showwarning("Perhatian", "Pilih minimal satu gambar terlebih dahulu!")
            return
        out_d = self.imgconv_out_var.get().strip()
        if not out_d:
            messagebox.showwarning("Perhatian", "Pilih folder tujuan hasil konversi!")
            return

        fmt = self.imgconv_target_fmt.get()
        q = self.imgconv_quality.get()
        resize_p = self.imgconv_resize.get()

        try:
            res = batch_convert_images(
                image_paths=self.imgconv_files,
                output_format=fmt,
                output_dir=out_d,
                quality=q,
                resize_percent=resize_p if resize_p != 100 else None,
            )
            messagebox.showinfo("Berhasil", f"Berhasil mengonversi {len(res)} gambar ke format {fmt.upper()}!")
            open_in_explorer(out_d)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # =========================================================================
    # TAB 5: DATA & DOKUMEN (EXCEL, CSV, JSON, MD, TXT, DOCX)
    # =========================================================================
    def _setup_tab_data_docs(self):
        frame = self.tab_datadocs

        tk.Label(
            frame,
            text="Konversi data tabel (Excel <-> CSV <-> JSON) dan dokumen teks (Markdown/TXT -> Word DOCX / PDF).",
            font=("Segoe UI", 9), bg="#F9FAFB", fg="#4B5563"
        ).pack(anchor="w", pady=(0, 10))

        # Box 1: Data Tabel
        box1 = tk.LabelFrame(frame, text=" 1. Konversi Data Tabel (Excel, CSV, JSON) ", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", padx=10, pady=8)
        box1.pack(fill=tk.X, pady=(0, 10))

        r1 = tk.Frame(box1, bg="#FFFFFF")
        r1.pack(fill=tk.X, pady=(0, 6))
        tk.Label(r1, text="File Sumber:", font=("Segoe UI", 9, "bold"), width=12, bg="#FFFFFF", anchor="w").pack(side=tk.LEFT)
        self.data_src_var = tk.StringVar()
        tk.Entry(r1, textvariable=self.data_src_var, font=("Segoe UI", 9), relief=tk.SOLID, bd=1).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        tk.Button(r1, text="Pilih File...", command=self._data_pick_src, relief=tk.SOLID, bd=1, font=("Segoe UI", 8)).pack(side=tk.LEFT)

        r2 = tk.Frame(box1, bg="#FFFFFF")
        r2.pack(fill=tk.X)
        tk.Label(r2, text="Konversi Ke:", font=("Segoe UI", 9), bg="#FFFFFF").pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(r2, text="📊 Excel (.xlsx)", command=lambda: self._data_convert_to(".xlsx"), bg="#10B981", fg="#FFF", font=("Segoe UI", 8, "bold"), relief=tk.FLAT).pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(r2, text="📄 CSV (.csv)", command=lambda: self._data_convert_to(".csv"), bg="#3B82F6", fg="#FFF", font=("Segoe UI", 8, "bold"), relief=tk.FLAT).pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(r2, text="📦 JSON (.json)", command=lambda: self._data_convert_to(".json"), bg="#8B5CF6", fg="#FFF", font=("Segoe UI", 8, "bold"), relief=tk.FLAT).pack(side=tk.LEFT)

        # Box 2: Dokumen Teks & Word
        box2 = tk.LabelFrame(frame, text=" 2. Dokumen Teks (Markdown / TXT <-> Word / PDF) ", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", padx=10, pady=8)
        box2.pack(fill=tk.X, pady=(0, 10))

        r3 = tk.Frame(box2, bg="#FFFFFF")
        r3.pack(fill=tk.X, pady=(0, 6))
        tk.Label(r3, text="File Dokumen:", font=("Segoe UI", 9, "bold"), width=12, bg="#FFFFFF", anchor="w").pack(side=tk.LEFT)
        self.doc_src_var = tk.StringVar()
        tk.Entry(r3, textvariable=self.doc_src_var, font=("Segoe UI", 9), relief=tk.SOLID, bd=1).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        tk.Button(r3, text="Pilih Dokumen...", command=self._doc_pick_src, relief=tk.SOLID, bd=1, font=("Segoe UI", 8)).pack(side=tk.LEFT)

        r4 = tk.Frame(box2, bg="#FFFFFF")
        r4.pack(fill=tk.X)
        tk.Label(r4, text="Aksi Konversi:", font=("Segoe UI", 9), bg="#FFFFFF").pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(r4, text="📝 TXT/MD ke Word (.docx)", command=self._doc_to_docx, bg="#2563EB", fg="#FFF", font=("Segoe UI", 8, "bold"), relief=tk.FLAT).pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(r4, text="📄 TXT/MD ke PDF", command=self._doc_to_pdf, bg="#DC2626", fg="#FFF", font=("Segoe UI", 8, "bold"), relief=tk.FLAT).pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(r4, text="📑 Word (.docx) ke Teks (.txt)", command=self._docx_to_txt, bg="#4B5563", fg="#FFF", font=("Segoe UI", 8), relief=tk.FLAT).pack(side=tk.LEFT)

    def _data_pick_src(self):
        f = filedialog.askopenfilename(
            title="Pilih File Tabel Data",
            filetypes=[("Data Files", "*.xlsx *.xls *.csv *.json"), ("Semua File", "*.*")]
        )
        if f:
            self.data_src_var.set(f)

    def _data_convert_to(self, ext: str):
        src = self.data_src_var.get().strip()
        if not src or not Path(src).is_file():
            messagebox.showwarning("Perhatian", "Pilih file data (.xlsx, .csv, atau .json) terlebih dahulu!")
            return
        out_p = filedialog.asksaveasfilename(
            title=f"Simpan Sebagai {ext.upper()}",
            defaultextension=ext,
            filetypes=[(f"File {ext}", f"*{ext}")]
        )
        if out_p:
            try:
                res = convert_tabular_data(src, out_p)
                messagebox.showinfo("Berhasil", f"Data berhasil dikonversi ke:\n{res}")
                open_in_explorer(res)
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def _doc_pick_src(self):
        f = filedialog.askopenfilename(
            title="Pilih Dokumen",
            filetypes=[("Dokumen", "*.txt *.md *.docx"), ("Semua File", "*.*")]
        )
        if f:
            self.doc_src_var.set(f)

    def _doc_to_docx(self):
        src = self.doc_src_var.get().strip()
        if not src or not Path(src).is_file():
            messagebox.showwarning("Perhatian", "Pilih file .txt atau .md terlebih dahulu!")
            return
        out_p = filedialog.asksaveasfilename(title="Simpan Sebagai Word", defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if out_p:
            try:
                res = convert_text_to_docx(src, out_p)
                messagebox.showinfo("Berhasil", f"Dokumen Word berhasil dibuat:\n{res}")
                open_in_explorer(res)
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def _doc_to_pdf(self):
        src = self.doc_src_var.get().strip()
        if not src or not Path(src).is_file():
            messagebox.showwarning("Perhatian", "Pilih file .txt atau .md terlebih dahulu!")
            return
        out_p = filedialog.asksaveasfilename(title="Simpan Sebagai PDF", defaultextension=".pdf", filetypes=[("PDF Document", "*.pdf")])
        if out_p:
            try:
                res = convert_markdown_to_pdf(src, out_p)
                messagebox.showinfo("Berhasil", f"PDF berhasil dibuat:\n{res}")
                open_in_explorer(res)
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def _docx_to_txt(self):
        src = self.doc_src_var.get().strip()
        if not src or not Path(src).is_file():
            messagebox.showwarning("Perhatian", "Pilih file .docx terlebih dahulu!")
            return
        out_p = filedialog.asksaveasfilename(title="Simpan Sebagai TXT", defaultextension=".txt", filetypes=[("Text File", "*.txt")])
        if out_p:
            try:
                res = convert_docx_to_text(src, out_p)
                messagebox.showinfo("Berhasil", f"Teks berhasil diekstrak:\n{res}")
                open_in_explorer(res)
            except Exception as e:
                messagebox.showerror("Error", str(e))


def main():
    app = OmniConverterGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
