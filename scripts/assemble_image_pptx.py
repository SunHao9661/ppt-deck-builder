#!/usr/bin/env python3
"""Assemble one final image per planned slide into an image-based PPTX."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp"}


def parse_hex(value: str) -> RGBColor:
    value = str(value).strip().lstrip("#")
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", value):
        raise ValueError(f"Invalid RGB color: {value!r}")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def set_canvas(prs: Presentation, ratio_text: str) -> tuple[int, int]:
    match = re.fullmatch(r"\s*(\d+)\s*:\s*(\d+)\s*", ratio_text)
    if not match:
        raise ValueError("visual.aspect_ratio must look like 16:9")
    left, right = int(match.group(1)), int(match.group(2))
    if left <= 0 or right <= 0:
        raise ValueError("Aspect ratio values must be positive")
    base = 7.5
    if left >= right:
        height = base
        width = base * left / right
    else:
        width = base
        height = base * right / left
    prs.slide_width = Inches(width)
    prs.slide_height = Inches(height)
    return int(prs.slide_width), int(prs.slide_height)


def image_candidates(images_dir: Path, index: int) -> list[Path]:
    patterns = [
        f"{index:02d}.*",
        f"{index}.*",
        f"slide-{index:02d}.*",
        f"slide-{index}.*",
        f"slide_{index:02d}.*",
        f"slide_{index}.*",
        f"*{index:02d}.*",
    ]
    found: list[Path] = []
    seen: set[str] = set()
    for pattern in patterns:
        for path in sorted(images_dir.glob(pattern)):
            if path.suffix.lower() not in IMAGE_EXTENSIONS:
                continue
            key = str(path.resolve())
            if key not in seen:
                found.append(path)
                seen.add(key)
    return found


def resolve_image(
    plan_dir: Path,
    images_dir: Path | None,
    slide: dict[str, Any],
    index: int,
) -> Path:
    declared = str(slide.get("image", "")).strip()
    if declared:
        path = Path(declared).expanduser()
        if not path.is_absolute():
            path = plan_dir / path
        path = path.resolve()
        if not path.exists():
            raise FileNotFoundError(f"Slide {index} image not found: {path}")
        return path

    if images_dir is None:
        raise FileNotFoundError(
            f"Slide {index} has no image field and no --images-dir was supplied"
        )
    candidates = image_candidates(images_dir, index)
    if not candidates:
        raise FileNotFoundError(
            f"Slide {index}: no matching image in {images_dir.resolve()}"
        )
    if len(candidates) > 1:
        names = ", ".join(path.name for path in candidates[:5])
        raise ValueError(
            f"Slide {index}: ambiguous image match ({names}). "
            "Set slide.image explicitly or rename the selected file."
        )
    return candidates[0].resolve()


def prepare_pptx_image(image_path: Path, assets_dir: Path, index: int) -> Path:
    if image_path.suffix.lower() != ".webp":
        return image_path
    assets_dir.mkdir(parents=True, exist_ok=True)
    destination = assets_dir / f"slide-{index:02d}.png"
    with Image.open(image_path) as image:
        image.save(destination, "PNG")
    return destination.resolve()


def add_picture_contained(
    slide: Any,
    image_path: Path,
    slide_width: int,
    slide_height: int,
    background: RGBColor,
) -> tuple[Any, bool]:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = background
    with Image.open(image_path) as image:
        image_width, image_height = image.size
    if image_width <= 0 or image_height <= 0:
        raise ValueError(f"Invalid image dimensions: {image_path}")

    source_ratio = image_width / image_height
    target_ratio = slide_width / slide_height
    exact = abs(source_ratio - target_ratio) <= 0.002
    if exact:
        left = 0
        top = 0
        width = slide_width
        height = slide_height
    elif source_ratio >= target_ratio:
        width = slide_width
        height = round(slide_width / source_ratio)
        left = 0
        top = (slide_height - height) // 2
    else:
        height = slide_height
        width = round(slide_height * source_ratio)
        left = (slide_width - width) // 2
        top = 0
    picture = slide.shapes.add_picture(
        str(image_path),
        int(left),
        int(top),
        int(width),
        int(height),
    )
    return picture, exact


def assemble(
    plan_path: Path,
    output_path: Path,
    images_dir: Path | None,
    force: bool,
) -> dict[str, Any]:
    if output_path.exists() and not force:
        raise FileExistsError(
            f"Output already exists: {output_path}. Pass --force to replace it."
        )
    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    slides = plan.get("slides")
    if not isinstance(slides, list) or not slides:
        raise ValueError("Deck plan must contain a non-empty slides array")
    visual = plan.get("visual") if isinstance(plan.get("visual"), dict) else {}
    ratio = str(visual.get("aspect_ratio", "16:9"))
    background_value = str(visual.get("background", "F4F4F4"))
    background = parse_hex(background_value)
    if images_dir is not None:
        images_dir = images_dir.expanduser().resolve()
        if not images_dir.is_dir():
            raise NotADirectoryError(f"Images directory not found: {images_dir}")

    prs = Presentation()
    slide_width, slide_height = set_canvas(prs, ratio)
    blank = prs.slide_layouts[6]
    warnings: list[str] = []
    used_images: list[str] = []
    plan_dir = plan_path.resolve().parent
    asset_dir = output_path.parent / f"{output_path.stem}-assets"

    for index, slide_spec in enumerate(slides, 1):
        if not isinstance(slide_spec, dict):
            raise ValueError(f"Slide {index} must be an object")
        image_path = resolve_image(plan_dir, images_dir, slide_spec, index)
        image_path = prepare_pptx_image(image_path, asset_dir, index)
        used_images.append(str(image_path))
        slide = prs.slides.add_slide(blank)
        picture, exact = add_picture_contained(
            slide,
            image_path,
            slide_width,
            slide_height,
            background,
        )
        title = " ".join(str(slide_spec.get("title", "")).split())
        picture.name = f"Slide image {index:02d}"
        picture._element.nvPicPr.cNvPr.set("descr", title or f"Slide {index}")
        if not exact:
            warnings.append(
                f"Slide {index}: image ratio differs from {ratio}; fitted with background"
            )
        notes = str(slide_spec.get("notes", "")).strip()
        if notes:
            slide.notes_slide.notes_text_frame.text = notes
        else:
            warnings.append(f"Slide {index}: no speaker notes")

    prs.core_properties.title = str(plan.get("title", ""))
    prs.core_properties.subject = str(plan.get("core_message", ""))
    prs.core_properties.comments = "Image-based deck assembled by ppt-deck-builder"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    return {
        "output": str(output_path.resolve()),
        "slides": len(slides),
        "aspect_ratio": ratio,
        "images": used_images,
        "warnings": warnings,
        "editability": "image-only",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("plan", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--images-dir", type=Path)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    try:
        result = assemble(args.plan, args.output, args.images_dir, args.force)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"Created {result['output']}")
        print(
            f"{result['slides']} image slides | {result['aspect_ratio']} | "
            "image-only"
        )
        for warning in result["warnings"]:
            print(f"WARN: {warning}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
