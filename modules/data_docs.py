"""
Module: Data & Document Conversions
- Excel (.xlsx) <-> CSV <-> JSON
- Markdown (.md) / Teks (.txt) -> PDF, HTML, DOCX
- DOCX (.docx) -> Teks, Markdown, PDF
"""

import json
from pathlib import Path
from typing import Optional
import pandas as pd
import markdown

try:
    import pymupdf as fitz
except ImportError:
    import fitz

from docx import Document


def convert_tabular_data(input_file: str, output_file: str) -> str:
    """
    Mengonversi file data tabel antar format: Excel (.xlsx), CSV (.csv), JSON (.json).
    Format ditentukan otomatis berdasarkan ekstensi file output.
    """
    in_path = Path(input_file).resolve()
    out_path = Path(output_file).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    in_ext = in_path.suffix.lower()
    out_ext = out_path.suffix.lower()

    # Baca file input ke Pandas DataFrame
    if in_ext in [".xlsx", ".xls"]:
        df = pd.read_excel(str(in_path))
    elif in_ext == ".csv":
        df = pd.read_csv(str(in_path))
    elif in_ext == ".json":
        df = pd.read_json(str(in_path))
    else:
        raise ValueError(f"Format input tidak didukung: {in_ext}. Gunakan .xlsx, .csv, atau .json")

    # Tulis ke file target
    if out_ext == ".csv":
        df.to_csv(str(out_path), index=False, encoding="utf-8-sig")
    elif out_ext in [".xlsx", ".xls"]:
        df.to_excel(str(out_path), index=False)
    elif out_ext == ".json":
        df.to_json(str(out_path), orient="records", indent=2, force_ascii=False)
    else:
        raise ValueError(f"Format output tidak didukung: {out_ext}. Gunakan .csv, .xlsx, atau .json")

    return str(out_path)


def convert_text_to_docx(input_text_or_md_file: str, output_docx_file: str) -> str:
    """Mengonversi file .txt atau .md menjadi file Word .docx."""
    in_path = Path(input_text_or_md_file).resolve()
    out_path = Path(output_docx_file).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    content = in_path.read_text(encoding="utf-8")
    doc = Document()

    for line in content.splitlines():
        line_s = line.strip()
        if line_s.startswith("# "):
            doc.add_heading(line_s[2:], level=1)
        elif line_s.startswith("## "):
            doc.add_heading(line_s[3:], level=2)
        elif line_s.startswith("### "):
            doc.add_heading(line_s[4:], level=3)
        elif line_s.startswith("- ") or line_s.startswith("* "):
            doc.add_paragraph(line_s[2:], style='List Bullet')
        else:
            doc.add_paragraph(line)

    doc.save(str(out_path))
    return str(out_path)


def convert_markdown_to_pdf(input_md_or_txt: str, output_pdf_file: str) -> str:
    """Mengonversi teks Markdown / Plain Text menjadi PDF rapi."""
    in_path = Path(input_md_or_txt).resolve()
    out_path = Path(output_pdf_file).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    text = in_path.read_text(encoding="utf-8")

    # Buat PDF via PyMuPDF story / page insertion yang rapi
    doc = fitz.open()
    page = doc.new_page(width=595.32, height=841.92) # A4
    rect = fitz.Rect(50, 50, 545.32, 791.92)

    # Insert text dengan auto wrap
    # PyMuPDF mendukung insert_textbox
    page.insert_textbox(rect, text, fontsize=11, fontname="helv", color=(0.1, 0.1, 0.1))

    doc.save(str(out_path), deflate=True)
    doc.close()
    return str(out_path)


def convert_docx_to_text(input_docx_file: str, output_txt_file: str) -> str:
    """Mengekstrak teks dari file Word (.docx) ke file .txt."""
    in_path = Path(input_docx_file).resolve()
    out_path = Path(output_txt_file).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    doc = Document(str(in_path))
    paragraphs = [p.text for p in doc.paragraphs]
    full_text = "\n".join(paragraphs)

    out_path.write_text(full_text, encoding="utf-8")
    return str(out_path)
