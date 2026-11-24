"""
iLovePDF Complete Offline Suite Engine
Menyediakan 28 fungsi manipulasi PDF dan konversi dokumen:
1. Merge PDF (Gabung)
2. Split PDF (Pisah)
3. Remove Pages (Hapus Halaman)
4. Rotate PDF (Putar Halaman)
5. Organize PDF (Urutkan Halaman Bebas)
6. Compress PDF (Kompresi Ukuran)
7. Repair PDF (Perbaiki Dokumen)
8. Protect PDF (Beri Password)
9. Unlock PDF (Buka Password)
10. Watermark PDF (Cap Air Teks)
11. Page Numbers (Nomor Halaman)
12. Sign PDF (Tanda Tangan Digital / e-Signature)
13. Redact PDF (Sensor Teks Rahasia Permanen)
14. Crop PDF (Potong Margin Halaman)
15. Edit PDF (Tambah Teks & Catatan)
16. Compare PDF (Bandingkan 2 PDF)
17. Images to PDF (JPG/PNG/WEBP ke PDF Berurutan)
18. PDF to Images (PDF ke JPG/PNG/WEBP dalam ZIP)
19. PDF to Word (DOCX)
20. Word to PDF (DOCX to PDF via Native COM / Fallback)
21. PDF to PowerPoint (PPTX)
22. PowerPoint to PDF (PPTX to PDF via Native COM / Fallback)
23. PDF to Excel (XLSX)
24. Excel to PDF (XLSX to PDF via Native COM / Fallback)
25. HTML to PDF (HTML/Web ke PDF)
26. PDF to PDF/A (Format Standar Arsip ISO)
27. PDF to Text (TXT)
28. Extract Embedded Images (Ekstrak Gambar Asli ke ZIP)
"""

import os
import io
import re
import base64
import difflib
import zipfile
from pathlib import Path
from typing import List, Optional, Tuple, Dict, Any

try:
    import pymupdf as fitz
except ImportError:
    import fitz

from PIL import Image
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
import pandas as pd

# Windows COM automation for exact Office rendering
HAS_WIN32COM = False
try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False


PAGE_SIZES = {
    "A4": (595.32, 841.92),
    "Letter": (612.0, 792.0),
    "Legal": (612.0, 1008.0),
}


def parse_page_indices(range_str: str, total_pages: int) -> List[int]:
    """Helper parse rentang halaman seperti '1-3, 5' ke 0-based index."""
    if not range_str or range_str.strip().lower() in ["all", "semua", ""]:
        return list(range(total_pages))
    pages = set()
    for part in range_str.split(","):
        part = part.strip()
        if "-" in part:
            try:
                s, e = part.split("-", 1)
                for p in range(max(1, int(s)), min(total_pages, int(e)) + 1):
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
    return sorted(list(pages)) if pages else list(range(total_pages))


# ----------------------------------------------------------------------
# 1. ORGANIZE PDF
# ----------------------------------------------------------------------

def engine_merge_pdfs(pdf_paths: List[str], output_path: str) -> str:
    out_p = Path(output_path).resolve()
    out_p.parent.mkdir(parents=True, exist_ok=True)

    merged = fitz.open()
    for p in pdf_paths:
        if Path(p).is_file():
            sub = fitz.open(p)
            merged.insert_pdf(sub)
            sub.close()

    merged.save(str(out_p), deflate=True, garbage=3)
    merged.close()
    return str(out_p)


def engine_split_pdf(pdf_path: str, output_path: str, mode: str = "ranges", ranges_str: str = "1") -> str:
    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    src_stem = Path(pdf_path).stem

    if mode == "each_page":
        zip_path = Path(output_path).with_suffix(".zip")
        with zipfile.ZipFile(str(zip_path), "w", zipfile.ZIP_DEFLATED) as zf:
            pad = len(str(total_pages))
            for i in range(total_pages):
                single_doc = fitz.open()
                single_doc.insert_pdf(doc, from_page=i, to_page=i)
                buf = single_doc.tobytes()
                single_doc.close()
                zf.writestr(f"{src_stem}_hal_{str(i+1).zfill(pad)}.pdf", buf)
        doc.close()
        return str(zip_path)
    else:
        target_indices = parse_page_indices(ranges_str, total_pages)
        out_doc = fitz.open()
        for p_idx in target_indices:
            out_doc.insert_pdf(doc, from_page=p_idx, to_page=p_idx)

        out_p = Path(output_path).resolve()
        out_doc.save(str(out_p), deflate=True)
        out_doc.close()
        doc.close()
        return str(out_p)


def engine_remove_pages(pdf_path: str, pages_to_remove_str: str, output_path: str) -> str:
    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    remove_indices = parse_page_indices(pages_to_remove_str, total_pages)

    for idx in sorted(remove_indices, reverse=True):
        if 0 <= idx < len(doc):
            doc.delete_page(idx)

    out_p = Path(output_path).resolve()
    doc.save(str(out_p), deflate=True, garbage=3)
    doc.close()
    return str(out_p)


def engine_rotate_pdf(pdf_path: str, rotation_deg: int, output_path: str, pages_str: str = "all") -> str:
    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    target_indices = parse_page_indices(pages_str, total_pages)

    for idx in target_indices:
        page = doc[idx]
        page.set_rotation((page.rotation + rotation_deg) % 360)

    out_p = Path(output_path).resolve()
    doc.save(str(out_p), deflate=True)
    doc.close()
    return str(out_p)


def engine_organize_pdf(pdf_path: str, new_order_str: str, output_path: str) -> str:
    doc = fitz.open(pdf_path)
    total_pages = len(doc)

    pages = []
    for part in new_order_str.split(","):
        part = part.strip()
        if part.isdigit():
            p_num = int(part)
            if 1 <= p_num <= total_pages:
                pages.append(p_num - 1)

    if not pages:
        pages = list(range(total_pages))

    doc.select(pages)
    out_p = Path(output_path).resolve()
    doc.save(str(out_p), deflate=True, garbage=3)
    doc.close()
    return str(out_p)


# ----------------------------------------------------------------------
# 2. OPTIMIZE & SECURITY
# ----------------------------------------------------------------------

def engine_compress_pdf(pdf_path: str, level: str, output_path: str) -> Tuple[str, int, int]:
    doc = fitz.open(pdf_path)
    orig_size = os.path.getsize(pdf_path)
    out_p = Path(output_path).resolve()

    if level == "extreme":
        new_doc = fitz.open()
        mat = fitz.Matrix(100 / 72.0, 100 / 72.0)
        for page in doc:
            pix = page.get_pixmap(matrix=mat)
            img_bytes = pix.pil_tobytes(format="JPEG", quality=60)
            rect = page.rect
            new_page = new_doc.new_page(width=rect.width, height=rect.height)
            new_page.insert_image(rect, stream=img_bytes)
        new_doc.save(str(out_p), deflate=True, garbage=4, clean=True)
        new_doc.close()
    elif level == "recommended":
        new_doc = fitz.open()
        mat = fitz.Matrix(140 / 72.0, 140 / 72.0)
        for page in doc:
            pix = page.get_pixmap(matrix=mat)
            img_bytes = pix.pil_tobytes(format="JPEG", quality=75)
            rect = page.rect
            new_page = new_doc.new_page(width=rect.width, height=rect.height)
            new_page.insert_image(rect, stream=img_bytes)
        new_doc.save(str(out_p), deflate=True, garbage=4, clean=True)
        new_doc.close()
    else:
        doc.save(str(out_p), deflate=True, garbage=4, deflate_images=True, deflate_fonts=True, clean=True)

    doc.close()
    new_size = os.path.getsize(str(out_p))
    return str(out_p), orig_size, new_size


def engine_repair_pdf(pdf_path: str, output_path: str) -> str:
    doc = fitz.open(pdf_path)
    out_p = Path(output_path).resolve()
    doc.save(str(out_p), garbage=4, clean=True, deflate=True)
    doc.close()
    return str(out_p)


def engine_protect_pdf(pdf_path: str, password: str, output_path: str) -> str:
    doc = fitz.open(pdf_path)
    out_p = Path(output_path).resolve()
    doc.save(
        str(out_p),
        encryption=fitz.PDF_ENCRYPT_AES_256,
        user_pw=password,
        owner_pw=password,
    )
    doc.close()
    return str(out_p)


def engine_unlock_pdf(pdf_path: str, password: str, output_path: str) -> str:
    doc = fitz.open(pdf_path)
    if doc.is_encrypted:
        success = doc.authenticate(password)
        if not success:
            doc.close()
            raise ValueError("Password yang dimasukkan salah.")

    out_p = Path(output_path).resolve()
    doc.save(str(out_p), encryption=fitz.PDF_ENCRYPT_NONE)
    doc.close()
    return str(out_p)


def engine_add_watermark(
    pdf_path: str,
    text: str,
    output_path: str,
    color_rgb: Tuple[float, float, float] = (0.7, 0.7, 0.7),
    opacity: float = 0.35,
    font_size: int = 40,
    angle: int = 45,
) -> str:
    doc = fitz.open(pdf_path)
    out_p = Path(output_path).resolve()

    for page in doc:
        rect = page.rect
        center_x = rect.width / 2.0
        center_y = rect.height / 2.0

        p = fitz.Point(center_x - (len(text) * font_size * 0.22), center_y)
        morph_tuple = (p, fitz.Matrix(angle)) if angle != 0 else None

        page.insert_text(
            p,
            text,
            fontsize=font_size,
            fontname="helv",
            color=color_rgb,
            morph=morph_tuple,
            fill_opacity=opacity,
        )

    doc.save(str(out_p), deflate=True)
    doc.close()
    return str(out_p)


def engine_add_page_numbers(
    pdf_path: str,
    output_path: str,
    position: str = "bottom-center",
    format_str: str = "Halaman {n} dari {total}",
    font_size: int = 10,
) -> str:
    doc = fitz.open(pdf_path)
    total = len(doc)
    out_p = Path(output_path).resolve()

    for i, page in enumerate(doc, start=1):
        num_text = format_str.replace("{n}", str(i)).replace("{total}", str(total))
        rect = page.rect

        if position == "bottom-right":
            pt = fitz.Point(rect.width - 150, rect.height - 25)
        elif position == "top-center":
            pt = fitz.Point((rect.width / 2.0) - 50, 30)
        else:
            pt = fitz.Point((rect.width / 2.0) - 50, rect.height - 25)

        page.insert_text(pt, num_text, fontsize=font_size, fontname="helv", color=(0.3, 0.3, 0.3))

    doc.save(str(out_p), deflate=True)
    doc.close()
    return str(out_p)


# ----------------------------------------------------------------------
# 3. EDIT, SIGN, CROP, REDACT & COMPARE
# ----------------------------------------------------------------------

def engine_sign_pdf(
    pdf_path: str,
    output_path: str,
    signature_data: str,
    page_num: int = 1,
    position: str = "bottom-right",
    signer_name: Optional[str] = None,
) -> str:
    doc = fitz.open(pdf_path)
    out_p = Path(output_path).resolve()

    target_idx = max(0, min(len(doc) - 1, page_num - 1))
    page = doc[target_idx]
    rect = page.rect

    if signature_data.startswith("data:image"):
        img_b64 = signature_data.split(",", 1)[1]
        sig_bytes = base64.b64decode(img_b64)
    elif Path(signature_data).is_file():
        sig_bytes = Path(signature_data).read_bytes()
    else:
        from PIL import ImageDraw
        sig_img = Image.new("RGBA", (300, 100), (255, 255, 255, 0))
        d = ImageDraw.Draw(sig_img)
        d.text((20, 35), signature_data, fill=(10, 40, 120))
        buf = io.BytesIO()
        sig_img.save(buf, format="PNG")
        sig_bytes = buf.getvalue()

    sig_w, sig_h = 160.0, 70.0
    margin = 40.0

    if position == "bottom-right":
        x0 = rect.width - sig_w - margin
        y0 = rect.height - sig_h - margin
    elif position == "bottom-left":
        x0 = margin
        y0 = rect.height - sig_h - margin
    elif position == "center":
        x0 = (rect.width - sig_w) / 2.0
        y0 = (rect.height - sig_h) / 2.0
    else:
        x0 = rect.width - sig_w - margin
        y0 = rect.height - sig_h - margin

    sig_rect = fitz.Rect(x0, y0, x0 + sig_w, y0 + sig_h)
    page.insert_image(sig_rect, stream=sig_bytes)

    if signer_name:
        page.insert_text(
            fitz.Point(x0, y0 + sig_h + 12),
            f"Ditandatangani oleh: {signer_name}",
            fontsize=8,
            fontname="helv",
            color=(0.2, 0.2, 0.2),
        )

    doc.save(str(out_p), deflate=True)
    doc.close()
    return str(out_p)


def engine_crop_pdf(
    pdf_path: str,
    output_path: str,
    crop_margin_pt: float = 36.0,
    pages_str: str = "all",
) -> str:
    doc = fitz.open(pdf_path)
    out_p = Path(output_path).resolve()
    target_indices = parse_page_indices(pages_str, len(doc))

    for idx in target_indices:
        page = doc[idx]
        rect = page.rect
        new_box = fitz.Rect(
            rect.x0 + crop_margin_pt,
            rect.y0 + crop_margin_pt,
            rect.x1 - crop_margin_pt,
            rect.y1 - crop_margin_pt,
        )
        if new_box.width > 50 and new_box.height > 50:
            page.set_cropbox(new_box)

    doc.save(str(out_p), deflate=True)
    doc.close()
    return str(out_p)


def engine_redact_pdf(
    pdf_path: str,
    output_path: str,
    search_keywords: List[str],
    redact_label: str = "[DISENSOR]",
) -> Tuple[str, int]:
    doc = fitz.open(pdf_path)
    out_p = Path(output_path).resolve()
    total_redacted = 0

    for page in doc:
        for kw in search_keywords:
            kw = kw.strip()
            if not kw:
                continue
            matches = page.search_for(kw)
            for rect in matches:
                page.add_redact_annot(
                    rect,
                    text=redact_label,
                    fill=(0, 0, 0),
                    text_color=(1, 1, 1),
                    fontsize=8,
                )
                total_redacted += 1

        page.apply_redactions()

    doc.save(str(out_p), deflate=True, garbage=4, clean=True)
    doc.close()
    return str(out_p), total_redacted


def engine_edit_pdf(
    pdf_path: str,
    output_path: str,
    page_num: int,
    text_content: str,
    x: float = 50.0,
    y: float = 100.0,
    font_size: int = 12,
    color_hex: str = "#000000",
) -> str:
    doc = fitz.open(pdf_path)
    out_p = Path(output_path).resolve()

    p_idx = max(0, min(len(doc) - 1, page_num - 1))
    page = doc[p_idx]

    color_hex = color_hex.lstrip("#")
    r = int(color_hex[0:2], 16) / 255.0
    g = int(color_hex[2:4], 16) / 255.0
    b = int(color_hex[4:6], 16) / 255.0

    rect = fitz.Rect(x, y, page.rect.width - 50, y + 200)
    page.insert_textbox(rect, text_content, fontsize=font_size, fontname="helv", color=(r, g, b))

    doc.save(str(out_p), deflate=True)
    doc.close()
    return str(out_p)


def engine_compare_pdf(pdf_path1: str, pdf_path2: str, output_path: str) -> str:
    doc1 = fitz.open(pdf_path1)
    doc2 = fitz.open(pdf_path2)
    out_p = Path(output_path).resolve()

    report_lines = [
        "============================================================",
        " LAPORAN PERBANDINGAN DOKUMEN PDF (PDF COMPARE REPORT)",
        "============================================================",
        f"Dokumen 1 (Asal): {Path(pdf_path1).name} | Total Halaman: {len(doc1)}",
        f"Dokumen 2 (Revisi): {Path(pdf_path2).name} | Total Halaman: {len(doc2)}",
        "------------------------------------------------------------\n",
    ]

    max_pages = max(len(doc1), len(doc2))
    has_diff = False

    for i in range(max_pages):
        t1 = doc1[i].get_text() if i < len(doc1) else "[Halaman ini tidak ada di Dokumen 1]"
        t2 = doc2[i].get_text() if i < len(doc2) else "[Halaman ini tidak ada di Dokumen 2]"

        if t1.strip() != t2.strip():
            has_diff = True
            report_lines.append(f"--- PERBEDAAN PADA HALAMAN {i+1} ---")
            diff = difflib.unified_diff(
                t1.splitlines(),
                t2.splitlines(),
                fromfile=f"Dokumen_1_Hal_{i+1}",
                tofile=f"Dokumen_2_Hal_{i+1}",
                lineterm="",
            )
            report_lines.extend(diff)
            report_lines.append("\n")

    if not has_diff:
        report_lines.append("HASIL: Kedua dokumen PDF memiliki isi teks yang 100% identik.")

    doc1.close()
    doc2.close()

    full_report = "\n".join(report_lines)
    out_p.write_text(full_report, encoding="utf-8")
    return str(out_p)


# ----------------------------------------------------------------------
# 4. CONVERT TO PDF
# ----------------------------------------------------------------------

def engine_images_to_pdf(
    image_paths: List[str],
    output_path: str,
    page_size: str = "A4",
    fit_mode: str = "fit",
    margin: float = 20.0,
    auto_rotate: bool = True,
) -> str:
    if not image_paths:
        raise ValueError("Tidak ada gambar yang dipilih.")

    out_p = Path(output_path).resolve()
    doc = fitz.open()

    for img_p in image_paths:
        p = Path(img_p)
        if not p.is_file():
            continue
        with Image.open(str(p)) as pil_img:
            img_w, img_h = pil_img.size

        if page_size.upper() == "ORIGINAL":
            page = doc.new_page(width=float(img_w), height=float(img_h))
            page.insert_image(fitz.Rect(0, 0, img_w, img_h), filename=str(p))
        else:
            base_w, base_h = PAGE_SIZES.get(page_size.upper(), PAGE_SIZES["A4"])
            if auto_rotate and img_w > img_h:
                pw, ph = max(base_w, base_h), min(base_w, base_h)
            else:
                pw, ph = min(base_w, base_h), max(base_w, base_h)

            page = doc.new_page(width=pw, height=ph)
            eff_w = max(10.0, pw - (2 * margin))
            eff_h = max(10.0, ph - (2 * margin))

            if fit_mode == "fit":
                scale = min(eff_w / img_w, eff_h / img_h)
                tw, th = img_w * scale, img_h * scale
                x0 = margin + (eff_w - tw) / 2.0
                y0 = margin + (eff_h - th) / 2.0
                rect = fitz.Rect(x0, y0, x0 + tw, y0 + th)
            else:
                rect = fitz.Rect(margin, margin, pw - margin, ph - margin)

            page.insert_image(rect, filename=str(p))

    doc.save(str(out_p), deflate=True, garbage=3)
    doc.close()
    return str(out_p)


def engine_word_to_pdf(docx_path: str, output_path: str) -> str:
    out_p = Path(output_path).resolve()
    src_abs = os.path.abspath(docx_path)
    out_abs = os.path.abspath(str(out_p))

    if HAS_WIN32COM:
        try:
            pythoncom.CoInitialize()
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            wb = word.Documents.Open(src_abs)
            wb.ExportAsFixedFormat(out_abs, 17) # 17 = wdExportFormatPDF
            wb.Close(False)
            word.Quit()
            pythoncom.CoUninitialize()
            return str(out_p)
        except Exception:
            pass

    doc = Document(docx_path)
    pdf_doc = fitz.open()
    pw, ph = PAGE_SIZES["A4"]
    page = pdf_doc.new_page(width=pw, height=ph)
    curr_y = 50.0

    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            curr_y += 12.0
            continue

        is_heading = p.style.name.startswith("Heading")
        fs = 16 if "1" in p.style.name else (14 if "2" in p.style.name else 11)
        fontname = "helv-bold" if is_heading else "helv"
        color = (0.1, 0.2, 0.4) if is_heading else (0.15, 0.15, 0.15)

        if curr_y > ph - 60:
            page = pdf_doc.new_page(width=pw, height=ph)
            curr_y = 50.0

        rect = fitz.Rect(50, curr_y, pw - 50, ph - 40)
        page.insert_textbox(rect, txt, fontsize=fs, fontname=fontname, color=color)
        curr_y += (fs * 1.5) + 6

    pdf_doc.save(str(out_p), deflate=True)
    pdf_doc.close()
    return str(out_p)


def engine_powerpoint_to_pdf(pptx_path: str, output_path: str) -> str:
    out_p = Path(output_path).resolve()
    src_abs = os.path.abspath(pptx_path)
    out_abs = os.path.abspath(str(out_p))

    if HAS_WIN32COM:
        try:
            pythoncom.CoInitialize()
            powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
            deck = powerpoint.Presentations.Open(src_abs, WithWindow=False)
            deck.SaveAs(out_abs, 32) # 32 = ppSaveAsPDF
            deck.Close()
            powerpoint.Quit()
            pythoncom.CoUninitialize()
            return str(out_p)
        except Exception:
            pass

    prs = Presentation(pptx_path)
    pdf_doc = fitz.open()
    pw = float(prs.slide_width) / 12700.0 * 72.0 / 72.0
    ph = float(prs.slide_height) / 12700.0 * 72.0 / 72.0
    if pw <= 0 or ph <= 0:
        pw, ph = 792.0, 595.0

    for slide in prs.slides:
        page = pdf_doc.new_page(width=pw, height=ph)
        curr_y = 50.0
        for shape in slide.shapes:
            if shape.has_text_frame:
                txt = shape.text_frame.text.strip()
                if txt:
                    rect = fitz.Rect(40, curr_y, pw - 40, curr_y + 100)
                    page.insert_textbox(rect, txt, fontsize=12, fontname="helv")
                    curr_y += 35.0

    pdf_doc.save(str(out_p), deflate=True)
    pdf_doc.close()
    return str(out_p)


def engine_excel_to_pdf(xlsx_path: str, output_path: str) -> str:
    out_p = Path(output_path).resolve()
    src_abs = os.path.abspath(xlsx_path)
    out_abs = os.path.abspath(str(out_p))

    if HAS_WIN32COM:
        try:
            pythoncom.CoInitialize()
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Open(src_abs)
            wb.ExportAsFixedFormat(0, out_abs) # 0 = xlTypePDF
            wb.Close(False)
            excel.Quit()
            pythoncom.CoUninitialize()
            return str(out_p)
        except Exception:
            pass

    df = pd.read_excel(xlsx_path)
    pdf_doc = fitz.open()
    pw, ph = PAGE_SIZES["A4"][1], PAGE_SIZES["A4"][0]
    page = pdf_doc.new_page(width=pw, height=ph)

    title = Path(xlsx_path).stem
    page.insert_text(fitz.Point(40, 40), f"Tabel Data: {title}", fontsize=14, fontname="helv-bold", color=(0.1, 0.2, 0.5))
    col_str = " | ".join(str(c)[:15] for c in df.columns)
    page.insert_text(fitz.Point(40, 65), col_str, fontsize=9, fontname="helv-bold", color=(0.1, 0.1, 0.1))
    page.draw_line(fitz.Point(40, 72), fitz.Point(pw - 40, 72), color=(0.6, 0.6, 0.6), width=1)

    curr_y = 88.0
    for _, row in df.iterrows():
        if curr_y > ph - 40:
            page = pdf_doc.new_page(width=pw, height=ph)
            curr_y = 50.0

        row_str = " | ".join(str(val)[:15] for val in row.values)
        page.insert_text(fitz.Point(40, curr_y), row_str, fontsize=8, fontname="helv", color=(0.2, 0.2, 0.2))
        curr_y += 14.0

    pdf_doc.save(str(out_p), deflate=True)
    pdf_doc.close()
    return str(out_p)


def engine_html_to_pdf(html_path_or_content: str, output_path: str) -> str:
    if Path(html_path_or_content).is_file():
        content = Path(html_path_or_content).read_text(encoding="utf-8", errors="ignore")
    else:
        content = html_path_or_content

    out_p = Path(output_path).resolve()
    pdf_doc = fitz.open()
    pw, ph = PAGE_SIZES["A4"]
    page = pdf_doc.new_page(width=pw, height=ph)
    rect = fitz.Rect(40, 40, pw - 40, ph - 40)
    page.insert_htmlbox(rect, content)
    pdf_doc.save(str(out_p), deflate=True)
    pdf_doc.close()
    return str(out_p)


def engine_pdf_to_pdfa(pdf_path: str, output_path: str) -> str:
    doc = fitz.open(pdf_path)
    out_p = Path(output_path).resolve()

    meta = doc.metadata or {}
    meta["title"] = meta.get("title") or Path(pdf_path).stem
    meta["creator"] = "OmniPDF PDF/A Archival Converter"
    meta["producer"] = "PyMuPDF PDF/A Engine"
    meta["format"] = "PDF/A-1b"
    doc.set_metadata(meta)

    doc.save(
        str(out_p),
        deflate=True,
        garbage=4,
        clean=True,
        deflate_fonts=True,
        deflate_images=True,
    )
    doc.close()
    return str(out_p)


# ----------------------------------------------------------------------
# 5. CONVERT FROM PDF
# ----------------------------------------------------------------------

def engine_pdf_to_images(pdf_path: str, output_zip_path: str, image_format: str = "png", dpi: int = 200) -> str:
    doc = fitz.open(pdf_path)
    zip_p = Path(output_zip_path).resolve()
    fmt = image_format.lower()
    mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    stem = Path(pdf_path).stem
    pad = len(str(len(doc)))

    with zipfile.ZipFile(str(zip_p), "w", zipfile.ZIP_DEFLATED) as zf:
        for i, page in enumerate(doc, start=1):
            pix = page.get_pixmap(matrix=mat, alpha=(fmt == "png"))
            buf = pix.pil_tobytes(format="PNG" if fmt == "png" else "JPEG")
            num_str = str(i).zfill(max(3, pad))
            zf.writestr(f"{stem}_hal_{num_str}.{fmt}", buf)

    doc.close()
    return str(zip_p)


def engine_extract_embedded_images(pdf_path: str, output_zip_path: str) -> Tuple[str, int]:
    doc = fitz.open(pdf_path)
    zip_p = Path(output_zip_path).resolve()
    stem = Path(pdf_path).stem
    img_count = 0

    with zipfile.ZipFile(str(zip_p), "w", zipfile.ZIP_DEFLATED) as zf:
        for page_idx, page in enumerate(doc, start=1):
            for img_info in page.get_images(full=True):
                xref = img_info[0]
                base_img = doc.extract_image(xref)
                if base_img:
                    img_count += 1
                    ext = base_img.get("ext", "png")
                    img_bytes = base_img.get("image")
                    zf.writestr(f"{stem}_foto_{img_count:03d}_p{page_idx}.{ext}", img_bytes)

    doc.close()
    return str(zip_p), img_count


def engine_pdf_to_word(pdf_path: str, output_path: str) -> str:
    doc = fitz.open(pdf_path)
    word_doc = Document()

    for i, page in enumerate(doc, start=1):
        if i > 1:
            word_doc.add_page_break()

        blocks = page.get_text("blocks")
        for b in blocks:
            text = b[4].strip()
            if text:
                word_doc.add_paragraph(text)

    out_p = Path(output_path).resolve()
    word_doc.save(str(out_p))
    doc.close()
    return str(out_p)


def engine_pdf_to_powerpoint(pdf_path: str, output_path: str) -> str:
    doc = fitz.open(pdf_path)
    prs = Presentation()
    p0 = doc[0]
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(10 * (p0.rect.height / p0.rect.width))

    blank_layout = prs.slide_layouts[6]
    mat = fitz.Matrix(200 / 72.0, 200 / 72.0)

    for page in doc:
        slide = prs.slides.add_slide(blank_layout)
        pix = page.get_pixmap(matrix=mat)
        img_bytes = pix.pil_tobytes(format="PNG")

        img_stream = io.BytesIO(img_bytes)
        slide.shapes.add_picture(img_stream, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)

    out_p = Path(output_path).resolve()
    prs.save(str(out_p))
    doc.close()
    return str(out_p)


def engine_pdf_to_excel(pdf_path: str, output_path: str) -> str:
    doc = fitz.open(pdf_path)
    writer = pd.ExcelWriter(output_path, engine="openpyxl")
    found_any_table = False

    sheet_count = 0
    for i, page in enumerate(doc, start=1):
        tabs = page.find_tables()
        if tabs and len(tabs.tables) > 0:
            for t_idx, tab in enumerate(tabs.tables, start=1):
                df = tab.to_pandas()
                sheet_count += 1
                sheet_name = f"Hal_{i}_Tabel_{t_idx}"[:30]
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                found_any_table = True

    if not found_any_table:
        lines = []
        for page in doc:
            for line in page.get_text().splitlines():
                if line.strip():
                    lines.append([line.strip()])
        df = pd.DataFrame(lines, columns=["Isi Teks"])
        df.to_excel(writer, sheet_name="Teks PDF", index=False)

    writer.close()
    doc.close()
    return output_path


def engine_pdf_to_text(pdf_path: str, output_path: str) -> str:
    doc = fitz.open(pdf_path)
    texts = []
    for i, page in enumerate(doc, start=1):
        texts.append(f"=== HALAMAN {i} ===\n" + page.get_text())
    doc.close()

    out_p = Path(output_path).resolve()
    out_p.write_text("\n\n".join(texts), encoding="utf-8")
    return str(out_p)
