#!/usr/bin/env python3
"""Render a PPTX to per-slide PNGs and a labeled contact sheet."""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from xml.sax.saxutils import escape

from PIL import Image, ImageDraw, ImageFont


def find_soffice() -> str:
    candidates = [
        os.environ.get("SOFFICE_PATH"),
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return str(candidate)
    raise FileNotFoundError(
        "LibreOffice not found. Install LibreOffice or set SOFFICE_PATH."
    )


def find_pdftoppm() -> str:
    candidate = os.environ.get("PDFTOPPM_PATH") or shutil.which("pdftoppm")
    if candidate and Path(candidate).exists():
        return str(candidate)
    raise FileNotFoundError(
        "pdftoppm not found. Install Poppler or set PDFTOPPM_PATH."
    )


def natural_key(path: Path) -> list[object]:
    return [
        int(piece) if piece.isdigit() else piece.lower()
        for piece in re.split(r"(\d+)", path.name)
    ]


def font_directories() -> list[Path]:
    candidates = [
        Path("/System/Library/Fonts"),
        Path("/System/Library/Fonts/Supplemental"),
        Path("/Library/Fonts"),
        Path.home() / "Library/Fonts",
        Path("/usr/share/fonts"),
        Path("/usr/local/share/fonts"),
        Path.home() / ".fonts",
        Path.home() / ".local/share/fonts",
        Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts",
    ]
    deduped: list[Path] = []
    seen: set[str] = set()
    for path in candidates:
        key = str(path)
        if key not in seen and path.exists():
            deduped.append(path)
            seen.add(key)
    return deduped


def write_fontconfig(tmp: Path) -> Path:
    config_path = tmp / "fonts.conf"
    cache_path = tmp / "fontcache"
    cache_path.mkdir()
    dirs = "\n".join(
        f"  <dir>{escape(str(path))}</dir>" for path in font_directories()
    )
    config_path.write_text(
        "<?xml version=\"1.0\"?>\n"
        "<!DOCTYPE fontconfig SYSTEM \"urn:fontconfig:fonts.dtd\">\n"
        "<fontconfig>\n"
        f"{dirs}\n"
        f"  <cachedir>{escape(str(cache_path))}</cachedir>\n"
        "</fontconfig>\n",
        encoding="utf-8",
    )
    return config_path


def load_label_font(size: int) -> ImageFont.ImageFont:
    candidates = [
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        r"C:\Windows\Fonts\arial.ttf",
    ]
    for candidate in candidates:
        if Path(candidate).exists():
            return ImageFont.truetype(candidate, size)
    return ImageFont.load_default()


def make_contact_sheet(images: list[Path], destination: Path) -> None:
    if not images:
        raise ValueError("No rendered images for contact sheet")
    thumb_w = 480
    thumb_h = 270
    gap = 24
    outer = 28
    label_h = 34
    cols = 3 if len(images) > 4 else 2
    rows = (len(images) + cols - 1) // cols
    width = outer * 2 + cols * thumb_w + (cols - 1) * gap
    height = outer * 2 + rows * (thumb_h + label_h) + (rows - 1) * gap
    canvas = Image.new("RGB", (width, height), "#171A1F")
    draw = ImageDraw.Draw(canvas)
    font = load_label_font(18)

    for index, image_path in enumerate(images):
        row = index // cols
        col = index % cols
        x = outer + col * (thumb_w + gap)
        y = outer + row * (thumb_h + label_h + gap)
        with Image.open(image_path) as image:
            image = image.convert("RGB")
            image.thumbnail((thumb_w, thumb_h), Image.Resampling.LANCZOS)
            tile = Image.new("RGB", (thumb_w, thumb_h), "#FFFFFF")
            tile.paste(
                image,
                ((thumb_w - image.width) // 2, (thumb_h - image.height) // 2),
            )
        canvas.paste(tile, (x, y))
        draw.text((x, y + thumb_h + 9), f"{index + 1:02d}", fill="#DDE4EA", font=font)

    destination.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(destination, "PNG")


def render(pptx: Path, output_dir: Path, dpi: int, overwrite: bool, contact: bool) -> dict:
    if not pptx.exists():
        raise FileNotFoundError(f"PPTX not found: {pptx}")
    if output_dir.exists() and not output_dir.is_dir():
        raise ValueError(f"Output path is not a directory: {output_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)
    existing = list(output_dir.glob("slide-*.png")) + list(output_dir.glob("contact-sheet.png"))
    if existing and not overwrite:
        raise FileExistsError(
            f"Render output already contains PNGs: {output_dir}. "
            "Choose a new directory or pass --overwrite."
        )

    soffice = find_soffice()
    pdftoppm = find_pdftoppm()
    with tempfile.TemporaryDirectory(prefix="ppt-deck-render-") as tmp_name:
        tmp = Path(tmp_name)
        profile = tmp / "lo-profile"
        font_config = write_fontconfig(tmp)
        environment = os.environ.copy()
        environment["FONTCONFIG_FILE"] = str(font_config)
        environment["FONTCONFIG_PATH"] = str(tmp)
        command = [
            soffice,
            "--headless",
            f"-env:UserInstallation={profile.as_uri()}",
            "--convert-to",
            "pdf",
            "--outdir",
            str(tmp),
            str(pptx.resolve()),
        ]
        proc = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=240,
            env=environment,
        )
        if proc.returncode != 0:
            detail = (proc.stderr or proc.stdout).strip()
            raise RuntimeError(f"LibreOffice conversion failed: {detail}")
        pdf_path = tmp / f"{pptx.stem}.pdf"
        if not pdf_path.exists():
            candidates = list(tmp.glob("*.pdf"))
            if not candidates:
                raise RuntimeError("LibreOffice did not create a PDF")
            pdf_path = candidates[0]

        prefix = tmp / "slide"
        proc = subprocess.run(
            [
                pdftoppm,
                "-png",
                "-r",
                str(dpi),
                str(pdf_path),
                str(prefix),
            ],
            capture_output=True,
            text=True,
            timeout=240,
        )
        if proc.returncode != 0:
            detail = (proc.stderr or proc.stdout).strip()
            raise RuntimeError(f"pdftoppm failed: {detail}")

        rendered_temp = sorted(tmp.glob("slide-*.png"), key=natural_key)
        if not rendered_temp:
            raise RuntimeError("No slide PNGs were produced")
        digits = max(2, len(str(len(rendered_temp))))
        final_images: list[Path] = []
        for index, source in enumerate(rendered_temp, 1):
            destination = output_dir / f"slide-{index:0{digits}d}.png"
            shutil.copyfile(source, destination)
            final_images.append(destination)

    contact_path = None
    if contact:
        contact_path = output_dir / "contact-sheet.png"
        make_contact_sheet(final_images, contact_path)

    return {
        "pptx": str(pptx.resolve()),
        "output_dir": str(output_dir.resolve()),
        "slides": len(final_images),
        "dpi": dpi,
        "images": [str(path.resolve()) for path in final_images],
        "contact_sheet": str(contact_path.resolve()) if contact_path else None,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path)
    parser.add_argument("output_dir", type=Path)
    parser.add_argument("--dpi", type=int, default=144)
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--no-contact-sheet", action="store_true")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    try:
        result = render(
            args.pptx,
            args.output_dir,
            max(72, min(300, args.dpi)),
            args.overwrite,
            not args.no_contact_sheet,
        )
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"Rendered {result['slides']} slides to {result['output_dir']}")
        if result["contact_sheet"]:
            print(f"Contact sheet: {result['contact_sheet']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
