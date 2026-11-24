"""
Module: Image(s) to PDF Converter
Mengonversi satu atau banyak gambar (JPG, PNG, WEBP, BMP, TIFF) menjadi satu file PDF berurutan,
atau PDF terpisah untuk setiap gambar.
"""

import os
from pathlib import Path
from typing import Callable, List, Optional
try:
    import pymupdf as fitz
except ImportError:
    import fitz
from PIL import Image


# Dimensi halaman standar (Points, 1 inch = 72 pt)
PAGE_SIZES = {
    "A4": (595.32, 841.92),
    "Letter": (612.0, 792.0),
    "Legal": (612.0, 1008.0),
}


def convert_images_to_pdf(
    image_paths: List[str],
    output_pdf_path: str,
    page_size: str = "A4",          # "A4", "Letter", "Original"
    fit_mode: str = "fit",           # "fit" (preserves aspect ratio), "stretch"
    auto_rotate: bool = True,        # Otomatis orientasi Landscape jika gambar melebar
    margin: float = 20.0,            # Margin dalam points (0 untuk tanpa batas/borderless)
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> str:
    """
    Menggabungkan daftar gambar menjadi satu dokumen PDF berurutan.

    :param image_paths: List path gambar dengan urutan yang diinginkan.
    :param output_pdf_path: Path tujuan file PDF.
    :param page_size: "A4", "Letter", atau "Original" (mengikuti ukuran pixel gambar).
    :param fit_mode: "fit" (menjaga proporsi) atau "stretch".
    :param auto_rotate: Jika True dan ukuran A4/Letter, orientasi halaman disesuaikan dengan gambar.
    :param margin: Jarak tepi (margin) dalam point.
    :param progress_callback: Fungsi callback (index_sekarang, total, nama_file).
    :return: Path file PDF yang dihasilkan.
    """
    if not image_paths:
        raise ValueError("Daftar gambar kosong. Silakan pilih minimal 1 gambar.")

    out_path = Path(output_pdf_path).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    doc = fitz.open()

    total_images = len(image_paths)

    for idx, img_p in enumerate(image_paths, start=1):
        img_path = Path(img_p).resolve()
        if not img_path.is_file():
            continue

        # Dapatkan dimensi gambar asli
        with Image.open(str(img_path)) as pil_img:
            img_w, img_h = pil_img.size

        if page_size.upper() == "ORIGINAL":
            # 1 pixel = 1 point pada 72 DPI (atau ukuran proporsional)
            page_w, page_h = float(img_w), float(img_h)
            page = doc.new_page(width=page_w, height=page_h)
            rect = fitz.Rect(0, 0, page_w, page_h)
            page.insert_image(rect, filename=str(img_path))
        else:
            base_w, base_h = PAGE_SIZES.get(page_size.upper(), PAGE_SIZES["A4"])

            # Sesuaikan orientasi jika auto_rotate aktif
            if auto_rotate and img_w > img_h:
                # Gambar landscape
                page_w, page_h = max(base_w, base_h), min(base_w, base_h)
            else:
                # Gambar portrait
                page_w, page_h = min(base_w, base_h), max(base_w, base_h)

            page = doc.new_page(width=page_w, height=page_h)

            # Area konten setelah margin
            eff_w = max(10.0, page_w - (2 * margin))
            eff_h = max(10.0, page_h - (2 * margin))

            if fit_mode == "fit":
                # Hitung skala aspek rasio terbaik
                scale = min(eff_w / img_w, eff_h / img_h)
                target_w = img_w * scale
                target_h = img_h * scale

                # Posisikan gambar di tengah (center)
                x0 = margin + (eff_w - target_w) / 2.0
                y0 = margin + (eff_h - target_h) / 2.0
                rect = fitz.Rect(x0, y0, x0 + target_w, y0 + target_h)
            else:
                # Stretch memenuhi batas margin
                rect = fitz.Rect(margin, margin, page_w - margin, page_h - margin)

            page.insert_image(rect, filename=str(img_path))

        if progress_callback:
            progress_callback(idx, total_images, img_path.name)

    # Simpan dengan kompresi optimal
    doc.save(str(out_path), deflate=True, garbage=3)
    doc.close()
    return str(out_path)
