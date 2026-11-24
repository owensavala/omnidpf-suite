"""
Core PDF to Image Converter Engine
Menggunakan PyMuPDF (fitz) - cepat, tanpa perlu install poppler tambahan di Windows.
"""

import os
import sys
from pathlib import Path
from typing import Callable, List, Optional
try:
    import pymupdf as fitz
except ImportError:
    import fitz


def parse_page_range(range_str: str, total_pages: int) -> List[int]:
    """
    Mengurai string rentang halaman seperti "1-3, 5, 7-9" (1-based index)
    menjadi list 0-based index.
    """
    if not range_str or range_str.strip().lower() == "all":
        return list(range(total_pages))

    pages = set()
    parts = range_str.split(",")
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            try:
                start_s, end_s = part.split("-", 1)
                start = max(1, int(start_s.strip()))
                end = min(total_pages, int(end_s.strip()))
                if start <= end:
                    for p in range(start, end + 1):
                        pages.add(p - 1)
            except ValueError:
                continue
        else:
            try:
                p = int(part)
                if 1 <= p <= total_pages:
                    pages.add(p - 1)
            except ValueError:
                continue

    return sorted(list(pages))


def convert_pdf_to_images(
    pdf_path: str,
    output_dir: Optional[str] = None,
    image_format: str = "png",
    dpi: int = 200,
    page_range: Optional[str] = None,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> List[str]:
    """
    Mengonversi halaman PDF menjadi file gambar.

    :param pdf_path: Path file PDF sumber.
    :param output_dir: Direktori tujuan untuk menyimpan gambar (default: subfolder nama PDF).
    :param image_format: Format gambar ('png', 'jpg'/'jpeg', 'webp').
    :param dpi: Resolusi gambar per inch (default 200, standar 150-300).
    :param page_range: Rentang halaman (misal: "1-5, 8", atau None untuk semua).
    :param progress_callback: Fungsi callback (current_page, total_pages, file_output_name).
    :return: List path file gambar yang dihasilkan.
    """
    pdf_path_obj = Path(pdf_path).resolve()
    if not pdf_path_obj.exists() or not pdf_path_obj.is_file():
        raise FileNotFoundError(f"File PDF tidak ditemukan: {pdf_path}")

    # Format normalisasi
    fmt = image_format.lower().replace(".", "").strip()
    if fmt == "jpeg":
        fmt = "jpg"
    if fmt not in ["png", "jpg", "webp"]:
        fmt = "png"

    # Direktori output
    if not output_dir:
        output_dir_obj = pdf_path_obj.parent / f"{pdf_path_obj.stem}_images"
    else:
        output_dir_obj = Path(output_dir).resolve()

    output_dir_obj.mkdir(parents=True, exist_ok=True)

    # Buka dokumen PDF
    doc = fitz.open(str(pdf_path_obj))
    total_doc_pages = len(doc)
    if total_doc_pages == 0:
        doc.close()
        return []

    target_page_indices = parse_page_range(page_range, total_doc_pages)
    if not target_page_indices:
        target_page_indices = list(range(total_doc_pages))

    saved_files = []
    # Matrix zoom berdasarkan DPI (PyMuPDF default 72 DPI)
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)

    pad_digits = len(str(total_doc_pages))

    total_to_process = len(target_page_indices)
    for i, page_idx in enumerate(target_page_indices, start=1):
        page = doc[page_idx]
        pix = page.get_pixmap(matrix=mat, alpha=(fmt == "png"))

        page_num_str = str(page_idx + 1).zfill(max(3, pad_digits))
        filename = f"{pdf_path_obj.stem}_hal_{page_num_str}.{fmt}"
        dest_file = output_dir_obj / filename

        pix.save(str(dest_file))
        saved_files.append(str(dest_file))

        if progress_callback:
            progress_callback(i, total_to_process, filename)

    doc.close()
    return saved_files
