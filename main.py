"""
OmniConverter CLI & GUI Entrypoint
"""

import sys
import argparse
from pathlib import Path

from modules.img_to_pdf import convert_images_to_pdf
from modules.pdf_to_img import convert_pdf_to_images
from modules.pdf_ops import merge_pdfs, split_pdf, extract_text_from_pdf, extract_embedded_images
from modules.img_ops import batch_convert_images, convert_image
from modules.data_docs import convert_tabular_data, convert_text_to_docx, convert_markdown_to_pdf


def run_img_to_pdf_cli(args):
    files = []
    for item in args.images:
        p = Path(item).resolve()
        if p.is_dir():
            valid_exts = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff"}
            found = [str(f) for f in sorted(p.iterdir()) if f.suffix.lower() in valid_exts]
            files.extend(found)
        elif p.is_file():
            files.append(str(p))

    if not files:
        print("[Error] Tidak ada file gambar yang valid ditemukan.")
        sys.exit(1)

    out_pdf = args.output or str(Path(files[0]).parent / "hasil_gabungan.pdf")
    print(f"Menggabungkan {len(files)} gambar menjadi 1 PDF berurutan...")

    def on_prog(c, t, fname):
        print(f"   [{c}/{t}] Memasukkan: {fname}")

    res = convert_images_to_pdf(
        image_paths=files,
        output_pdf_path=out_pdf,
        page_size=args.size,
        fit_mode=args.fit,
        auto_rotate=not args.no_auto_rotate,
        margin=args.margin,
        progress_callback=on_prog,
    )
    print(f"[Berhasil] PDF tersimpan di: {res}")


def run_pdf_to_img_cli(args):
    p = Path(args.pdf).resolve()
    if not p.is_file():
        print(f"[Error] File PDF tidak ditemukan: {p}")
        sys.exit(1)

    print(f"Mengonversi {p.name} ke gambar {args.format.upper()} (DPI: {args.dpi})...")

    def on_prog(c, t, fname):
        print(f"   [{c}/{t}] Disimpan: {fname}")

    res = convert_pdf_to_images(
        pdf_path=str(p),
        output_dir=args.output,
        image_format=args.format,
        dpi=args.dpi,
        page_range=args.pages,
        progress_callback=on_prog,
    )
    print(f"[Berhasil] Total {len(res)} gambar dibuat.")


def main():
    # Jika tidak ada argumen sama sekali, buka antarmuka grafis (GUI)
    if len(sys.argv) == 1 or (len(sys.argv) == 2 and sys.argv[1] == "--gui"):
        from gui_app import main as launch_gui
        launch_gui()
        return

    parser = argparse.ArgumentParser(
        description="⚡ OmniConverter - Sistem Konversi Lengkap Serba Bisa"
    )
    subparsers = parser.add_subparsers(dest="command", help="Perintah konversi yang tersedia")

    # Command: img2pdf
    p_i2p = subparsers.add_parser("img2pdf", help="Gabung satu/banyak gambar menjadi satu PDF berurutan")
    p_i2p.add_argument("images", nargs="+", help="File gambar atau folder gambar")
    p_i2p.add_argument("-o", "--output", help="Path file output PDF (default: hasil_gabungan.pdf)")
    p_i2p.add_argument("--size", default="A4", choices=["A4", "Original", "Letter"], help="Ukuran halaman")
    p_i2p.add_argument("--fit", default="fit", choices=["fit", "stretch"], help="Mode fit gambar")
    p_i2p.add_argument("--margin", type=float, default=20.0, help="Margin tepi dalam point (default 20)")
    p_i2p.add_argument("--no-auto-rotate", action="store_true", help="Jangan otomatis rotate orientasi")

    # Command: pdf2img
    p_p2i = subparsers.add_parser("pdf2img", help="Konversi PDF menjadi gambar (PNG/JPG/WEBP)")
    p_p2i.add_argument("pdf", help="Path file PDF")
    p_p2i.add_argument("-o", "--output", help="Folder output")
    p_p2i.add_argument("-f", "--format", default="png", choices=["png", "jpg", "webp"], help="Format gambar")
    p_p2i.add_argument("--dpi", type=int, default=200, help="Resolusi DPI (default 200)")
    p_p2i.add_argument("-p", "--pages", default=None, help="Rentang halaman (contoh: '1-3, 5')")

    # Command: merge
    p_merge = subparsers.add_parser("merge", help="Gabungkan beberapa file PDF menjadi satu")
    p_merge.add_argument("pdfs", nargs="+", help="Daftar file PDF yang ingin digabungkan")
    p_merge.add_argument("-o", "--output", required=True, help="Path file PDF hasil penggabungan")

    # Command: split
    p_split = subparsers.add_parser("split", help="Pisahkan halaman PDF")
    p_split.add_argument("pdf", help="Path file PDF")
    p_split.add_argument("-o", "--output", required=True, help="Folder hasil pemisahan")
    p_split.add_argument("-r", "--ranges", help="Rentang halaman pemisahan (contoh: '1-3, 4-10')")

    # Command: imgconv
    p_imgconv = subparsers.add_parser("imgconv", help="Ubah format gambar (PNG/JPG/WEBP/ICO/BMP/TIFF)")
    p_imgconv.add_argument("images", nargs="+", help="File gambar yang ingin diubah")
    p_imgconv.add_argument("-f", "--format", required=True, choices=["png", "jpg", "webp", "ico", "bmp", "tiff"])
    p_imgconv.add_argument("-o", "--output", required=True, help="Folder tujuan")
    p_imgconv.add_argument("-q", "--quality", type=int, default=90, help="Kualitas kompresi 1-100")
    p_imgconv.add_argument("--resize", type=float, default=None, help="Resize persentase (misal 50 untuk 50%)")

    # Command: data
    p_data = subparsers.add_parser("data", help="Konversi data tabel (Excel, CSV, JSON)")
    p_data.add_argument("input", help="File input (.xlsx, .csv, atau .json)")
    p_data.add_argument("-o", "--output", required=True, help="File output (.xlsx, .csv, atau .json)")

    args = parser.parse_args()

    if args.command == "img2pdf":
        run_img_to_pdf_cli(args)
    elif args.command == "pdf2img":
        run_pdf_to_img_cli(args)
    elif args.command == "merge":
        res = merge_pdfs(args.pdfs, args.output)
        print(f"[Berhasil] PDF gabungan tersimpan di: {res}")
    elif args.command == "split":
        mode = "ranges" if args.ranges else "each_page"
        res = split_pdf(args.pdf, args.output, split_mode=mode, ranges_str=args.ranges)
        print(f"[Berhasil] Berhasil memisahkan PDF menjadi {len(res)} file di: {args.output}")
    elif args.command == "imgconv":
        res = batch_convert_images(args.images, args.format, args.output, quality=args.quality, resize_percent=args.resize)
        print(f"[Berhasil] Berhasil mengonversi {len(res)} gambar ke {args.output}")
    elif args.command == "data":
        res = convert_tabular_data(args.input, args.output)
        print(f"[Berhasil] Data berhasil dikonversi ke: {res}")
    else:
        # Fallback to single argument PDF
        inp = Path(sys.argv[1]).resolve()
        if inp.is_file() and inp.suffix.lower() == ".pdf":
            from modules.pdf_to_img import convert_pdf_to_images
            res = convert_pdf_to_images(str(inp))
            print(f"[Berhasil] {len(res)} gambar dibuat dari PDF.")
        else:
            parser.print_help()


if __name__ == "__main__":
    main()
