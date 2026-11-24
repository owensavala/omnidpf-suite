"""
Module: Universal Image Format Converter
Mengonversi format gambar antar PNG, JPG, WEBP, ICO, BMP, TIFF dengan kontrol kualitas dan resize.
"""

from pathlib import Path
from typing import Callable, List, Optional, Tuple
from PIL import Image


def convert_image(
    input_image_path: str,
    output_format: str,
    output_path: Optional[str] = None,
    quality: int = 90,
    resize_dimensions: Optional[Tuple[int, int]] = None,
    resize_percent: Optional[float] = None,
) -> str:
    """
    Mengonversi satu gambar ke format lain.

    :param input_image_path: Path gambar asal.
    :param output_format: 'png', 'jpg', 'webp', 'ico', 'bmp', 'tiff'.
    :param output_path: Path tujuan (opsional).
    :param quality: Kualitas kompresi 1-100 (untuk JPG / WEBP).
    :param resize_dimensions: Tuple (width, height) atau None.
    :param resize_percent: Persentase ukuran (misal 50 untuk setengah ukuran).
    :return: Path file gambar hasil konversi.
    """
    src_path = Path(input_image_path).resolve()
    if not src_path.is_file():
        raise FileNotFoundError(f"File gambar tidak ditemukan: {input_image_path}")

    target_fmt = output_format.lower().replace(".", "").strip()
    if target_fmt == "jpeg":
        target_fmt = "jpg"

    if not output_path:
        out_file = src_path.parent / f"{src_path.stem}_converted.{target_fmt}"
    else:
        out_file = Path(output_path).resolve()

    out_file.parent.mkdir(parents=True, exist_ok=True)

    with Image.open(str(src_path)) as img:
        # Resize jika diminta
        if resize_percent and resize_percent > 0 and resize_percent != 100:
            new_w = max(1, int(img.width * (resize_percent / 100.0)))
            new_h = max(1, int(img.height * (resize_percent / 100.0)))
            img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
        elif resize_dimensions:
            rw, rh = resize_dimensions
            if rw > 0 and rh > 0:
                img = img.resize((rw, rh), Image.Resampling.LANCZOS)

        # Penyesuaian mode warna
        if target_fmt in ["jpg", "jpeg"]:
            if img.mode in ("RGBA", "LA", "P"):
                # Buat background putih untuk transparansi
                bg = Image.new("RGB", img.size, (255, 255, 255))
                if img.mode == "P":
                    img = img.convert("RGBA")
                bg.paste(img, mask=img.split()[-1] if img.mode == "RGBA" else None)
                img = bg
            elif img.mode != "RGB":
                img = img.convert("RGB")
            img.save(str(out_file), "JPEG", quality=quality, optimize=True)

        elif target_fmt == "png":
            img.save(str(out_file), "PNG", optimize=True)

        elif target_fmt == "webp":
            img.save(str(out_file), "WEBP", quality=quality)

        elif target_fmt == "ico":
            # Icon Windows standar
            icon_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
            img.save(str(out_file), format="ICO", sizes=icon_sizes)

        elif target_fmt == "bmp":
            if img.mode not in ("RGB", "L"):
                img = img.convert("RGB")
            img.save(str(out_file), "BMP")

        elif target_fmt == "tiff":
            img.save(str(out_file), "TIFF")

        else:
            img.save(str(out_file))

    return str(out_file)


def batch_convert_images(
    image_paths: List[str],
    output_format: str,
    output_dir: str,
    quality: int = 90,
    resize_percent: Optional[float] = None,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> List[str]:
    """Konversi banyak gambar sekaligus ke format target."""
    out_dir_obj = Path(output_dir).resolve()
    out_dir_obj.mkdir(parents=True, exist_ok=True)

    results = []
    total = len(image_paths)

    fmt = output_format.lower().replace(".", "").strip()
    if fmt == "jpeg":
        fmt = "jpg"

    for idx, p in enumerate(image_paths, start=1):
        p_obj = Path(p).resolve()
        if not p_obj.is_file():
            continue
        dest_path = out_dir_obj / f"{p_obj.stem}.{fmt}"
        res = convert_image(
            input_image_path=str(p_obj),
            output_format=fmt,
            output_path=str(dest_path),
            quality=quality,
            resize_percent=resize_percent,
        )
        results.append(res)
        if progress_callback:
            progress_callback(idx, total, dest_path.name)

    return results
