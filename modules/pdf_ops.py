"""
Module: PDF Operations
- Menggabungkan beberapa file PDF menjadi satu file.
- Memisahkan PDF (Split halaman).
- Mengekstrak teks dari PDF (TXT / Markdown).
- Mengekstrak semua gambar yang tertanam di dalam PDF.
"""

from pathlib import Path
from typing import Callable, List, Optional
try:
    import pymupdf as fitz
except ImportError:
    import fitz


def merge_pdfs(
    pdf_paths: List[str],
    output_pdf_path: str,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> str:
    """Menggabungkan banyak file PDF menjadi satu dokumen PDF baru."""
    if not pdf_paths:
        raise ValueError("Daftar PDF kosong.")

    out_path = Path(output_pdf_path).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    merged_doc = fitz.open()
    total = len(pdf_paths)

    for idx, p in enumerate(pdf_paths, start=1):
        p_obj = Path(p).resolve()
        if not p_obj.is_file():
            continue
        sub_doc = fitz.open(str(p_obj))
        merged_doc.insert_pdf(sub_doc)
        sub_doc.close()

        if progress_callback:
            progress_callback(idx, total, p_obj.name)

    merged_doc.save(str(out_path), deflate=True, garbage=3)
    merged_doc.close()
    return str(out_path)


def split_pdf(
    pdf_path: str,
    output_dir: str,
    split_mode: str = "each_page",  # "each_page" atau "ranges"
    ranges_str: Optional[str] = None, # misal "1-3, 4-10"
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> List[str]:
    """Memisahkan dokumen PDF per halaman atau per rentang halaman."""
    src_path = Path(pdf_path).resolve()
    if not src_path.is_file():
        raise FileNotFoundError(f"File PDF tidak ditemukan: {pdf_path}")

    out_dir_obj = Path(output_dir).resolve()
    out_dir_obj.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(str(src_path))
    total_pages = len(doc)
    created_files = []

    if split_mode == "each_page":
        pad = len(str(total_pages))
        for i in range(total_pages):
            out_doc = fitz.open()
            out_doc.insert_pdf(doc, from_page=i, to_page=i)
            page_num_str = str(i + 1).zfill(max(3, pad))
            fname = f"{src_path.stem}_hal_{page_num_str}.pdf"
            target_file = out_dir_obj / fname
            out_doc.save(str(target_file))
            out_doc.close()
            created_files.append(str(target_file))

            if progress_callback:
                progress_callback(i + 1, total_pages, fname)
    else:
        # Split per kelompok rentang
        groups = [r.strip() for r in (ranges_str or "").split(",") if r.strip()]
        for idx, grp in enumerate(groups, start=1):
            if "-" in grp:
                s_str, e_str = grp.split("-", 1)
                start_p = max(0, int(s_str.strip()) - 1)
                end_p = min(total_pages - 1, int(e_str.strip()) - 1)
            else:
                start_p = max(0, int(grp.strip()) - 1)
                end_p = start_p

            if start_p <= end_p:
                out_doc = fitz.open()
                out_doc.insert_pdf(doc, from_page=start_p, to_page=end_p)
                fname = f"{src_path.stem}_bagian_{idx}_(hal_{start_p+1}-{end_p+1}).pdf"
                target_file = out_dir_obj / fname
                out_doc.save(str(target_file))
                out_doc.close()
                created_files.append(str(target_file))

                if progress_callback:
                    progress_callback(idx, len(groups), fname)

    doc.close()
    return created_files


def extract_text_from_pdf(pdf_path: str, output_txt_path: Optional[str] = None) -> str:
    """Mengekstrak seluruh teks dari PDF menjadi file .txt atau string."""
    src_path = Path(pdf_path).resolve()
    doc = fitz.open(str(src_path))
    text_content = []

    for idx, page in enumerate(doc, start=1):
        text_content.append(f"--- HALAMAN {idx} ---\n" + page.get_text())

    doc.close()
    full_text = "\n\n".join(text_content)

    if output_txt_path:
        out_p = Path(output_txt_path).resolve()
        out_p.parent.mkdir(parents=True, exist_ok=True)
        out_p.write_text(full_text, encoding="utf-8")
        return str(out_p)

    return full_text


def extract_embedded_images(
    pdf_path: str,
    output_dir: str,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> List[str]:
    """Mengekstrak gambar asli yang disematkan (embedded) di dalam PDF."""
    src_path = Path(pdf_path).resolve()
    out_dir_obj = Path(output_dir).resolve()
    out_dir_obj.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(str(src_path))
    saved_images = []
    img_counter = 0

    for page_idx, page in enumerate(doc, start=1):
        image_list = page.get_images(full=True)
        for img_info in image_list:
            xref = img_info[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]

            img_counter += 1
            img_fname = f"{src_path.stem}_ekstrak_{img_counter:03d}_p{page_idx}.{image_ext}"
            img_path = out_dir_obj / img_fname
            img_path.write_bytes(image_bytes)
            saved_images.append(str(img_path))

            if progress_callback:
                progress_callback(page_idx, len(doc), img_fname)

    doc.close()
    return saved_images
