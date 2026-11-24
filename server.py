"""
OmniPDF Offline Server (iLovePDF Complete Clone)
Web server lokal Flask yang menyediakan 28 fitur iLovePDF secara 100% offline.
"""

import os
import sys
import uuid
import time
import shutil
import webbrowser
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file, url_for

from modules.ilovepdf_engine import (
    engine_merge_pdfs,
    engine_split_pdf,
    engine_remove_pages,
    engine_rotate_pdf,
    engine_organize_pdf,
    engine_compress_pdf,
    engine_repair_pdf,
    engine_protect_pdf,
    engine_unlock_pdf,
    engine_add_watermark,
    engine_add_page_numbers,
    engine_sign_pdf,
    engine_crop_pdf,
    engine_redact_pdf,
    engine_edit_pdf,
    engine_compare_pdf,
    engine_images_to_pdf,
    engine_pdf_to_images,
    engine_extract_embedded_images,
    engine_pdf_to_word,
    engine_word_to_pdf,
    engine_pdf_to_powerpoint,
    engine_powerpoint_to_pdf,
    engine_pdf_to_excel,
    engine_excel_to_pdf,
    engine_html_to_pdf,
    engine_pdf_to_pdfa,
    engine_pdf_to_text,
)

BASE_DIR = Path(__file__).parent.resolve()
UPLOAD_DIR = BASE_DIR / "temp_uploads"
OUTPUT_DIR = BASE_DIR / "temp_outputs"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = Flask(__name__, template_folder="templates", static_folder="static")
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024


def cleanup_old_files():
    now = time.time()
    for folder in [UPLOAD_DIR, OUTPUT_DIR]:
        for item in folder.iterdir():
            try:
                if item.is_file() and now - item.stat().st_mtime > 7200:
                    item.unlink()
            except Exception:
                pass


@app.route("/")
def index():
    cleanup_old_files()
    return render_template("index.html")


@app.route("/api/process/<tool_name>", methods=["POST"])
def process_tool(tool_name):
    cleanup_old_files()
    req_id = str(uuid.uuid4())[:8]

    try:
        uploaded_files = request.files.getlist("files")
        saved_input_paths = []

        if uploaded_files and uploaded_files[0].filename != "":
            for f in uploaded_files:
                safe_name = f"{req_id}_{f.filename}"
                target_path = UPLOAD_DIR / safe_name
                f.save(str(target_path))
                saved_input_paths.append(str(target_path))

        first_stem = Path(uploaded_files[0].filename).stem if uploaded_files and uploaded_files[0].filename else "dokumen"

        # ----------------------------------------------------
        # ROUTING ALAT LENGKAP (28 TOOLS)
        # ----------------------------------------------------

        if tool_name == "merge":
            out_name = f"{first_stem}_gabungan.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            file_order_str = request.form.get("file_order")
            if file_order_str:
                indices = [int(x) for x in file_order_str.split(",") if x.isdigit()]
                ordered_paths = [saved_input_paths[i] for i in indices if i < len(saved_input_paths)]
                if ordered_paths:
                    saved_input_paths = ordered_paths
            res_path = engine_merge_pdfs(saved_input_paths, out_file)

        elif tool_name == "split":
            mode = request.form.get("mode", "ranges")
            ranges = request.form.get("ranges", "1")
            ext = ".zip" if mode == "each_page" else ".pdf"
            out_name = f"{first_stem}_split{ext}"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_split_pdf(saved_input_paths[0], out_file, mode=mode, ranges_str=ranges)
            out_name = Path(res_path).name.replace(f"{req_id}_", "")

        elif tool_name == "remove_pages":
            pages_to_remove = request.form.get("pages", "")
            out_name = f"{first_stem}_halaman_dihapus.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_remove_pages(saved_input_paths[0], pages_to_remove, out_file)

        elif tool_name == "rotate":
            deg = int(request.form.get("degrees", 90))
            pages = request.form.get("pages", "all")
            out_name = f"{first_stem}_diputar_{deg}deg.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_rotate_pdf(saved_input_paths[0], deg, out_file, pages_str=pages)

        elif tool_name == "organize":
            new_order = request.form.get("order", "")
            out_name = f"{first_stem}_diurutkan.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_organize_pdf(saved_input_paths[0], new_order, out_file)

        elif tool_name == "compress":
            level = request.form.get("level", "recommended")
            out_name = f"{first_stem}_dikompres.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path, orig_sz, new_sz = engine_compress_pdf(saved_input_paths[0], level, out_file)
            saved_pct = max(0, int((1 - (new_sz / orig_sz)) * 100))
            return jsonify({
                "status": "ok",
                "filename": out_name,
                "download_url": f"/download/{Path(res_path).name}?name={out_name}",
                "detail": f"Ukuran berkurang dari {(orig_sz/1024):.1f} KB menjadi {(new_sz/1024):.1f} KB (Hemat {saved_pct}%)."
            })

        elif tool_name == "repair":
            out_name = f"{first_stem}_diperbaiki.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_repair_pdf(saved_input_paths[0], out_file)

        elif tool_name == "protect":
            password = request.form.get("password", "")
            if not password:
                return jsonify({"status": "error", "message": "Password tidak boleh kosong."}), 400
            out_name = f"{first_stem}_terlindungi.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_protect_pdf(saved_input_paths[0], password, out_file)

        elif tool_name == "unlock":
            password = request.form.get("password", "")
            out_name = f"{first_stem}_terbuka.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_unlock_pdf(saved_input_paths[0], password, out_file)

        elif tool_name == "watermark":
            text = request.form.get("text", "CONFIDENTIAL")
            angle = int(request.form.get("angle", 45))
            opacity = float(request.form.get("opacity", 0.35))
            font_size = int(request.form.get("font_size", 40))
            out_name = f"{first_stem}_watermark.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_add_watermark(saved_input_paths[0], text, out_file, opacity=opacity, font_size=font_size, angle=angle)

        elif tool_name == "page_numbers":
            pos = request.form.get("position", "bottom-center")
            fmt = request.form.get("format", "Halaman {n} dari {total}")
            out_name = f"{first_stem}_bernomor.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_add_page_numbers(saved_input_paths[0], out_file, position=pos, format_str=fmt)

        elif tool_name == "sign":
            sig_data = request.form.get("signature_data", "")
            page_num = int(request.form.get("page_num", 1))
            pos = request.form.get("position", "bottom-right")
            signer = request.form.get("signer_name", "")
            out_name = f"{first_stem}_bertanda_tangan.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_sign_pdf(saved_input_paths[0], out_file, sig_data, page_num=page_num, position=pos, signer_name=signer)

        elif tool_name == "crop":
            margin_pt = float(request.form.get("margin", 36.0))
            pages = request.form.get("pages", "all")
            out_name = f"{first_stem}_dipotong.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_crop_pdf(saved_input_paths[0], out_file, crop_margin_pt=margin_pt, pages_str=pages)

        elif tool_name == "redact":
            kws = [k.strip() for k in request.form.get("keywords", "").split(",") if k.strip()]
            label = request.form.get("label", "[DISENSOR]")
            out_name = f"{first_stem}_disensor.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path, count = engine_redact_pdf(saved_input_paths[0], out_file, search_keywords=kws, redact_label=label)
            return jsonify({
                "status": "ok",
                "filename": out_name,
                "download_url": f"/download/{Path(res_path).name}?name={out_name}",
                "detail": f"Berhasil menyensor {count} kemunculan data sensitif secara permanen."
            })

        elif tool_name == "edit":
            page_num = int(request.form.get("page_num", 1))
            txt = request.form.get("text", "")
            color = request.form.get("color", "#000000")
            fs = int(request.form.get("font_size", 12))
            out_name = f"{first_stem}_diedit.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_edit_pdf(saved_input_paths[0], out_file, page_num=page_num, text_content=txt, font_size=fs, color_hex=color)

        elif tool_name == "compare":
            if len(saved_input_paths) < 2:
                return jsonify({"status": "error", "message": "Harap unggah minimal 2 file PDF untuk dibandingkan."}), 400
            out_name = f"Laporan_Perbandingan_{first_stem}.txt"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_compare_pdf(saved_input_paths[0], saved_input_paths[1], out_file)

        elif tool_name == "img2pdf":
            file_order_str = request.form.get("file_order")
            if file_order_str:
                indices = [int(x) for x in file_order_str.split(",") if x.isdigit()]
                ordered_paths = [saved_input_paths[i] for i in indices if i < len(saved_input_paths)]
                if ordered_paths:
                    saved_input_paths = ordered_paths

            page_size = request.form.get("page_size", "A4")
            fit_mode = request.form.get("fit_mode", "fit")
            margin = float(request.form.get("margin", 20.0))
            auto_rotate = request.form.get("auto_rotate", "true").lower() == "true"

            out_name = f"{first_stem}_gabungan_gambar.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_images_to_pdf(saved_input_paths, out_file, page_size=page_size, fit_mode=fit_mode, margin=margin, auto_rotate=auto_rotate)

        elif tool_name == "pdf2img":
            fmt = request.form.get("format", "png")
            dpi = int(request.form.get("dpi", 200))
            out_name = f"{first_stem}_gambar.zip"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_pdf_to_images(saved_input_paths[0], out_file, image_format=fmt, dpi=dpi)

        elif tool_name == "extract_images":
            out_name = f"{first_stem}_gambar_tertanam.zip"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path, count = engine_extract_embedded_images(saved_input_paths[0], out_file)
            return jsonify({
                "status": "ok",
                "filename": out_name,
                "download_url": f"/download/{Path(res_path).name}?name={out_name}",
                "detail": f"Berhasil mengekstrak {count} gambar asli ke dalam file ZIP."
            })

        elif tool_name == "pdf2word":
            out_name = f"{first_stem}.docx"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_pdf_to_word(saved_input_paths[0], out_file)

        elif tool_name == "word2pdf":
            out_name = f"{first_stem}.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_word_to_pdf(saved_input_paths[0], out_file)

        elif tool_name == "pdf2pptx":
            out_name = f"{first_stem}.pptx"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_pdf_to_powerpoint(saved_input_paths[0], out_file)

        elif tool_name == "powerpoint2pdf":
            out_name = f"{first_stem}.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_powerpoint_to_pdf(saved_input_paths[0], out_file)

        elif tool_name == "pdf2excel":
            out_name = f"{first_stem}.xlsx"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_pdf_to_excel(saved_input_paths[0], out_file)

        elif tool_name == "excel2pdf":
            out_name = f"{first_stem}.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_excel_to_pdf(saved_input_paths[0], out_file)

        elif tool_name == "html2pdf":
            raw_html = request.form.get("html_content")
            src = saved_input_paths[0] if saved_input_paths else (raw_html or "<h1>Dokumen Kosong</h1>")
            out_name = f"{first_stem}.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_html_to_pdf(src, out_file)

        elif tool_name == "pdf2pdfa":
            out_name = f"{first_stem}_pdfa.pdf"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_pdf_to_pdfa(saved_input_paths[0], out_file)

        elif tool_name == "pdf2text":
            out_name = f"{first_stem}.txt"
            out_file = str(OUTPUT_DIR / f"{req_id}_{out_name}")
            res_path = engine_pdf_to_text(saved_input_paths[0], out_file)

        else:
            return jsonify({"status": "error", "message": f"Alat '{tool_name}' tidak dikenali."}), 400

        final_filename = Path(res_path).name
        return jsonify({
            "status": "ok",
            "filename": out_name,
            "download_url": f"/download/{final_filename}?name={out_name}",
            "detail": "Proses konversi berhasil diselesaikan!"
        })

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/download/<path:filename>")
def download_file(filename):
    target = OUTPUT_DIR / filename
    custom_name = request.args.get("name", filename)
    if not target.is_file():
        return "File tidak ditemukan atau telah kedaluwarsa.", 404
    return send_file(str(target), as_attachment=True, download_name=custom_name)


def main():
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", 5000))
    url = f"http://localhost:{port}"

    print("=" * 60)
    print("🚀 OmniPDF Complete (iLovePDF Offline Suite) Sedang Berjalan!")
    print(f"👉 Akses Lokal:   {url}")
    print(f"👉 Akses Jaringan: http://0.0.0.0:{port}")
    print("   100% Bebas, Aman, Cepat & Tanpa Batas Ukuran File.")
    print("=" * 60)

    # Otomatis buka browser hanya jika dijalankan interaktif di Windows lokal
    if os.environ.get("NO_BROWSER") != "1" and sys.platform == "win32":
        import threading
        def open_browser():
            time.sleep(1.2)
            webbrowser.open(url)
        threading.Thread(target=open_browser, daemon=True).start()

    app.run(host=host, port=port, debug=False)


if __name__ == "__main__":
    main()
