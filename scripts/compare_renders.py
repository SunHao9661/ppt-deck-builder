#!/usr/bin/env python3
"""Create side-by-side, overlay, and difference images for PPT slide renders."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

from PIL import Image, ImageChops, ImageDraw, ImageFont, ImageOps


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp"}


def natural_key(path: Path) -> list[object]:
    return [
        int(piece) if piece.isdigit() else piece.lower()
        for piece in re.split(r"(\d+)", path.name)
    ]


def list_images(path: Path) -> list[Path]:
    if not path.is_dir():
        raise NotADirectoryError(f"Not a directory: {path}")
    images = [
        item
        for item in path.iterdir()
        if item.is_file() and item.suffix.lower() in IMAGE_EXTENSIONS
    ]
    return sorted(images, key=natural_key)


def load_font(size: int) -> ImageFont.ImageFont:
    candidates = [
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        r"C:\Windows\Fonts\arial.ttf",
    ]
    for candidate in candidates:
        if Path(candidate).exists():
            return ImageFont.truetype(candidate, size)
    return ImageFont.load_default()


def fit_on_canvas(image: Image.Image, size: tuple[int, int], background: str) -> Image.Image:
    image = image.convert("RGB")
    image.thumbnail(size, Image.Resampling.LANCZOS)
    canvas = Image.new("RGB", size, background)
    canvas.paste(
        image,
        ((size[0] - image.width) // 2, (size[1] - image.height) // 2),
    )
    return canvas


def common_size(left: Image.Image, right: Image.Image) -> tuple[int, int]:
    target_width = min(1280, max(left.width, right.width))
    scales = (
        target_width / max(1, image.width)
        for image in (left, right)
    )
    target_height = max(
        round(image.height * scale)
        for image, scale in zip((left, right), scales)
    )
    return target_width, max(1, target_height)


def normalized_pair(left_path: Path, right_path: Path) -> tuple[Image.Image, Image.Image]:
    with Image.open(left_path) as raw_left, Image.open(right_path) as raw_right:
        left = raw_left.convert("RGB")
        right = raw_right.convert("RGB")
        size = common_size(left, right)
        return (
            fit_on_canvas(left, size, "#161A20"),
            fit_on_canvas(right, size, "#161A20"),
        )


def image_metrics(left: Image.Image, right: Image.Image, threshold: int) -> dict[str, Any]:
    diff = ImageChops.difference(left, right)
    gray = diff.convert("L")
    hist = gray.histogram()
    total = max(1, left.width * left.height)
    changed = sum(hist[threshold + 1 :])
    weighted = sum(value * count for value, count in enumerate(hist))
    bbox_value = diff.getbbox()
    return {
        "mean_abs_diff_0_255": round(weighted / total, 4),
        "changed_pixel_ratio": round(changed / total, 6),
        "changed_percent": round(changed / total * 100, 3),
        "difference_bbox": list(bbox_value) if bbox_value else None,
        "threshold": threshold,
        "size": {"w": left.width, "h": left.height},
    }


def labeled_side_by_side(
    left: Image.Image,
    right: Image.Image,
    left_label: str,
    right_label: str,
) -> Image.Image:
    gap = 20
    outer = 24
    label_h = 34
    width = outer * 2 + left.width + right.width + gap
    height = outer * 2 + label_h + max(left.height, right.height)
    canvas = Image.new("RGB", (width, height), "#171A1F")
    draw = ImageDraw.Draw(canvas)
    font = load_font(18)
    left_x = outer
    right_x = outer + left.width + gap
    y = outer + label_h
    draw.text((left_x, outer), left_label, fill="#E6EEF5", font=font)
    draw.text((right_x, outer), right_label, fill="#E6EEF5", font=font)
    canvas.paste(left, (left_x, y))
    canvas.paste(right, (right_x, y))
    return canvas


def make_comparison(
    left_path: Path,
    right_path: Path,
    out_dir: Path,
    index: int,
    threshold: int,
    left_label: str,
    right_label: str,
) -> dict[str, Any]:
    left, right = normalized_pair(left_path, right_path)
    metrics = image_metrics(left, right, threshold)
    prefix = out_dir / f"compare-{index:02d}"
    side = labeled_side_by_side(left, right, left_label, right_label)
    overlay = Image.blend(left, right, 0.5)
    diff = ImageChops.difference(left, right)
    diff_map = ImageOps.colorize(
        diff.convert("L"),
        black="#000000",
        white="#FF3B30",
    )
    side_path = prefix.with_name(prefix.name + "-side-by-side.png")
    overlay_path = prefix.with_name(prefix.name + "-overlay.png")
    diff_path = prefix.with_name(prefix.name + "-diff.png")
    side.save(side_path)
    overlay.save(overlay_path)
    diff_map.save(diff_path)
    return {
        "left": str(left_path.resolve()),
        "right": str(right_path.resolve()),
        "side_by_side": str(side_path.resolve()),
        "overlay": str(overlay_path.resolve()),
        "difference_map": str(diff_path.resolve()),
        "metrics": metrics,
        "_side_image": side_path,
    }


def comparison_sheet(paths: list[Path], destination: Path) -> None:
    if not paths:
        return
    tile_w = 900
    gap = 18
    outer = 24
    label_h = 28
    tiles: list[Image.Image] = []
    for path in paths:
        with Image.open(path) as image:
            image = image.convert("RGB")
            ratio = image.height / image.width
            image = image.resize(
                (tile_w, round(tile_w * ratio)),
                Image.Resampling.LANCZOS,
            )
            tiles.append(image)
    cols = 1 if len(tiles) <= 2 else 2
    rows = (len(tiles) + cols - 1) // cols
    tile_h = max(image.height for image in tiles)
    width = outer * 2 + cols * tile_w + (cols - 1) * gap
    height = outer * 2 + rows * (tile_h + label_h) + (rows - 1) * gap
    canvas = Image.new("RGB", (width, height), "#171A1F")
    draw = ImageDraw.Draw(canvas)
    font = load_font(16)
    for index, image in enumerate(tiles):
        row = index // cols
        col = index % cols
        x = outer + col * (tile_w + gap)
        y = outer + row * (tile_h + label_h + gap)
        canvas.paste(image, (x, y + label_h))
        draw.text((x, y + 4), f"Comparison {index + 1:02d}", fill="#E6EEF5", font=font)
    canvas.save(destination)


def compare(
    left: Path,
    right: Path,
    out_dir: Path,
    threshold: int,
    left_label: str,
    right_label: str,
) -> dict[str, Any]:
    out_dir.mkdir(parents=True, exist_ok=True)
    if left.is_file() and right.is_file():
        pairs = [(left, right)]
    elif left.is_dir() and right.is_dir():
        left_images = list_images(left)
        right_images = list_images(right)
        if len(left_images) != len(right_images):
            raise ValueError(
                f"Image count mismatch: {len(left_images)} reference vs "
                f"{len(right_images)} output"
            )
        pairs = list(zip(left_images, right_images))
    else:
        raise ValueError("Pass two image files or two directories")
    if not pairs:
        raise ValueError("No images found to compare")

    comparisons: list[dict[str, Any]] = []
    side_paths: list[Path] = []
    for index, (left_path, right_path) in enumerate(pairs, 1):
        row = make_comparison(
            left_path,
            right_path,
            out_dir,
            index,
            threshold,
            left_label,
            right_label,
        )
        side_paths.append(row.pop("_side_image"))
        comparisons.append(row)

    sheet_path = out_dir / "comparison-contact-sheet.png"
    comparison_sheet(side_paths, sheet_path)
    report = {
        "reference": str(left.resolve()),
        "output": str(right.resolve()),
        "pairs": len(comparisons),
        "threshold": threshold,
        "contact_sheet": str(sheet_path.resolve()),
        "comparisons": comparisons,
    }
    report_path = out_dir / "comparison-report.json"
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    report["report"] = str(report_path.resolve())
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("reference", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--threshold", type=int, default=8)
    parser.add_argument("--left-label", default="Reference")
    parser.add_argument("--right-label", default="Output")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    try:
        result = compare(
            args.reference,
            args.output,
            args.out_dir,
            max(0, min(255, args.threshold)),
            args.left_label,
            args.right_label,
        )
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        means = [
            row["metrics"]["mean_abs_diff_0_255"] for row in result["comparisons"]
        ]
        print(f"Compared {result['pairs']} page(s)")
        print(f"Mean difference range: {min(means):.2f}–{max(means):.2f} / 255")
        print(f"Contact sheet: {result['contact_sheet']}")
        print(f"Report: {result['report']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
