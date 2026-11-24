#!/usr/bin/env python3
"""Inspect PPTX slide and shape structure without modifying the deck."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def rgb_hex(color: Any) -> str | None:
    try:
        rgb = color.rgb
        if rgb is not None:
            return str(rgb)
    except Exception:
        return None
    return None


def bbox(shape: Any, slide_width: int, slide_height: int) -> dict[str, Any]:
    try:
        left = int(shape.left)
        top = int(shape.top)
        width = int(shape.width)
        height = int(shape.height)
    except (TypeError, ValueError):
        return {}
    return {
        "emu": {"x": left, "y": top, "w": width, "h": height},
        "inches": {
            "x": round(left / 914400, 4),
            "y": round(top / 914400, 4),
            "w": round(width / 914400, 4),
            "h": round(height / 914400, 4),
        },
        "fraction": {
            "x": round(left / slide_width, 6),
            "y": round(top / slide_height, 6),
            "w": round(width / slide_width, 6),
            "h": round(height / slide_height, 6),
        },
    }


def text_details(shape: Any, max_text: int) -> dict[str, Any] | None:
    if not getattr(shape, "has_text_frame", False):
        return None
    paragraphs: list[dict[str, Any]] = []
    for p_index, paragraph in enumerate(shape.text_frame.paragraphs):
        runs: list[dict[str, Any]] = []
        for r_index, run in enumerate(paragraph.runs):
            font = run.font
            runs.append(
                {
                    "run": r_index,
                    "text": run.text,
                    "font": font.name,
                    "size_pt": round(float(font.size.pt), 2) if font.size else None,
                    "bold": font.bold,
                    "italic": font.italic,
                    "color": rgb_hex(font.color),
                }
            )
        text = paragraph.text
        if max_text > 0 and len(text) > max_text:
            text = text[:max_text] + "…"
        paragraphs.append(
            {
                "paragraph": p_index,
                "level": paragraph.level,
                "alignment": str(paragraph.alignment) if paragraph.alignment else None,
                "text": text,
                "runs": runs,
            }
        )
    return {
        "full_text": shape.text_frame.text,
        "paragraphs": paragraphs,
    }


def walk_shape(
    shape: Any,
    slide_width: int,
    slide_height: int,
    max_text: int,
    depth: int = 0,
) -> dict[str, Any]:
    try:
        shape_type = shape.shape_type.name
    except Exception:
        shape_type = str(getattr(shape, "shape_type", "UNKNOWN"))
    row: dict[str, Any] = {
        "shape_id": getattr(shape, "shape_id", None),
        "name": getattr(shape, "name", ""),
        "type": shape_type,
        "depth": depth,
        "bbox": bbox(shape, slide_width, slide_height),
    }
    if getattr(shape, "has_chart", False):
        try:
            row["chart_type"] = str(shape.chart.chart_type)
            row["chart_series"] = len(shape.chart.series)
        except Exception:
            pass
    if getattr(shape, "has_table", False):
        try:
            table = shape.table
            row["table"] = {
                "rows": len(table.rows),
                "columns": len(table.columns),
                "cells": [
                    [
                        table.cell(row_index, column_index).text
                        for column_index in range(len(table.columns))
                    ]
                    for row_index in range(len(table.rows))
                ],
            }
        except Exception:
            pass
    try:
        alt = shape._element.nvPicPr.cNvPr.get("descr", "")
        if alt:
            row["alt_text"] = alt
    except Exception:
        pass
    details = text_details(shape, max_text)
    if details:
        row.update(details)
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            row["children"] = [
                walk_shape(
                    child,
                    slide_width,
                    slide_height,
                    max_text,
                    depth + 1,
                )
                for child in shape.shapes
            ]
    except Exception:
        pass
    return row


def inspect(path: Path, max_text: int, include_notes: bool) -> dict[str, Any]:
    prs = Presentation(path)
    slide_width = int(prs.slide_width)
    slide_height = int(prs.slide_height)
    slides: list[dict[str, Any]] = []
    for index, slide in enumerate(prs.slides, 1):
        row: dict[str, Any] = {
            "slide": index,
            "shapes": [
                walk_shape(shape, slide_width, slide_height, max_text)
                for shape in slide.shapes
            ],
        }
        if include_notes:
            try:
                row["notes"] = slide.notes_slide.notes_text_frame.text
            except Exception:
                row["notes"] = ""
        slides.append(row)
    return {
        "file": str(path.resolve()),
        "slide_count": len(prs.slides),
        "slide_size": {
            "width_emu": slide_width,
            "height_emu": slide_height,
            "width_in": round(slide_width / 914400, 4),
            "height_in": round(slide_height / 914400, 4),
            "ratio": round(slide_width / slide_height, 4),
        },
        "slides": slides,
    }


def flatten_shapes(shapes: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for shape in shapes:
        rows.append(shape)
        rows.extend(flatten_shapes(shape.get("children", [])))
    return rows


def print_human(result: dict[str, Any]) -> None:
    size = result["slide_size"]
    print(
        f"{result['file']} | {result['slide_count']} slides | "
        f"{size['width_in']} × {size['height_in']} in"
    )
    for slide in result["slides"]:
        print(f"\nSlide {slide['slide']}")
        for shape in flatten_shapes(slide["shapes"]):
            indent = "  " * int(shape.get("depth", 0))
            text = " ".join(shape.get("full_text", "").split())
            if len(text) > 80:
                text = text[:77] + "..."
            bbox_value = shape.get("bbox", {}).get("inches", {})
            location = ""
            if bbox_value:
                location = (
                    f" [{bbox_value['x']:.2f},{bbox_value['y']:.2f} "
                    f"{bbox_value['w']:.2f}×{bbox_value['h']:.2f}]"
                )
            print(
                f"{indent}#{shape.get('shape_id')} {shape['type']} "
                f"{shape['name']!r}{location}"
                + (f" | {text}" if text else "")
            )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--notes", action="store_true")
    parser.add_argument("--max-text", type=int, default=180)
    args = parser.parse_args()
    try:
        result = inspect(args.pptx, max(0, args.max_text), args.notes)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print_human(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
