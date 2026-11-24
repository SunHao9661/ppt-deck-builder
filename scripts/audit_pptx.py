#!/usr/bin/env python3
"""Audit an editable PPTX for common content, structure, and layout defects."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


PLACEHOLDER_RE = re.compile(
    r"(lorem\s+ipsum|click\s+to\s+add|your\s+text\s+here|"
    r"\[insert|slide\s+title|untitled\s+presentation|"
    r"\btodo\b|\btbd\b|placeholder)",
    re.IGNORECASE,
)


def text_of(shape: Any) -> str:
    if not getattr(shape, "has_text_frame", False):
        return ""
    return shape.text_frame.text.strip()


def title_of(slide: Any) -> str:
    named = [shape for shape in slide.shapes if shape.name == "Title"]
    for shape in named:
        text = text_of(shape)
        if text:
            return text
    for shape in slide.shapes:
        try:
            alt = shape._element.nvPicPr.cNvPr.get("descr", "").strip()
        except Exception:
            alt = ""
        if alt:
            return alt
    best = ""
    best_area = -1
    for shape in slide.shapes:
        text = text_of(shape)
        if not text or len(text) > 180:
            continue
        try:
            area = int(shape.width) * int(shape.height)
        except (TypeError, ValueError):
            area = 0
        if area > best_area:
            best_area = area
            best = text
    return best


def notes_of(slide: Any) -> str:
    try:
        return slide.notes_slide.notes_text_frame.text.strip()
    except Exception:
        return ""


def slide_text(slide: Any) -> str:
    parts: list[str] = []
    for shape in slide.shapes:
        text = text_of(shape)
        if text:
            parts.append(text)
    return "\n".join(parts)


def visible_char_count(text: str) -> int:
    return sum(1 for char in text if not char.isspace())


def audit(path: Path) -> dict[str, Any]:
    prs = Presentation(path)
    width = int(prs.slide_width)
    height = int(prs.slide_height)
    issues: list[dict[str, Any]] = []

    def add(severity: str, code: str, message: str, slide: int | None = None) -> None:
        entry: dict[str, Any] = {"severity": severity, "code": code, "message": message}
        if slide is not None:
            entry["slide"] = slide
        issues.append(entry)

    if not prs.slides:
        add("ERROR", "NO_SLIDES", "The presentation contains no slides")
        return {
            "file": str(path.resolve()),
            "slide_count": 0,
            "slide_size": {"width_emu": width, "height_emu": height},
            "ghost_deck": [],
            "issues": issues,
            "notes_covered": 0,
        }

    ratio = width / height
    if abs(ratio - 16 / 9) > 0.015 and abs(ratio - 4 / 3) > 0.015:
        add("WARN", "SLIDE_RATIO", f"Unusual slide aspect ratio: {ratio:.3f}")

    ghost: list[dict[str, Any]] = []
    notes_covered = 0
    title_sizes: list[float] = []
    previous_title = ""
    tolerance = int(0.02 * 914400)

    for index, slide in enumerate(prs.slides, 1):
        title = title_of(slide)
        notes = notes_of(slide)
        text = slide_text(slide)
        line_title = {"slide": index, "title": " ".join(title.split())}
        ghost.append(line_title)

        if notes:
            notes_covered += 1
        elif index not in {1, len(prs.slides)}:
            add("WARN", "MISSING_NOTES", "No speaker notes on this slide", index)

        if title and title == previous_title:
            add("WARN", "DUPLICATE_TITLE", f"Consecutive duplicate title: {title!r}", index)
        if title:
            previous_title = title
        else:
            add("WARN", "NO_TITLE", "No title-like text was found", index)

        match = PLACEHOLDER_RE.search(text)
        if match:
            add(
                "ERROR",
                "PLACEHOLDER_TEXT",
                f"Possible placeholder text: {match.group(0)!r}",
                index,
            )

        chars = visible_char_count(text)
        if chars > 420:
            add(
                "WARN",
                "TEXT_DENSITY",
                f"Very high visible text volume ({chars} non-space characters)",
                index,
            )

        has_visual = False
        body_sizes: list[float] = []
        for shape in slide.shapes:
            if getattr(shape, "has_chart", False):
                has_visual = True
            try:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    has_visual = True
                    alt = shape._element.nvPicPr.cNvPr.get("descr", "").strip()
                    if not alt:
                        add(
                            "WARN",
                            "MISSING_ALT_TEXT",
                            f"Picture {shape.name!r} has no alternative text",
                            index,
                        )
            except Exception:
                pass

            try:
                left = int(shape.left)
                top = int(shape.top)
                right = left + int(shape.width)
                bottom = top + int(shape.height)
            except (TypeError, ValueError):
                left = top = right = bottom = 0
            if left < -tolerance or top < -tolerance or right > width + tolerance or bottom > height + tolerance:
                add(
                    "ERROR",
                    "OUT_OF_BOUNDS",
                    f"Object {shape.name!r} extends outside the slide",
                    index,
                )

            if not getattr(shape, "has_text_frame", False):
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if run.font.size is None:
                        continue
                    size = float(run.font.size.pt)
                    if shape.name == "Title":
                        title_sizes.append(size)
                    elif shape.name != "Footer" and size < 8.5:
                        body_sizes.append(size)
        if body_sizes:
            add(
                "WARN",
                "SMALL_TEXT",
                f"Readable text below 8.5 pt: min {min(body_sizes):.1f} pt",
                index,
            )
        if not text and not has_visual:
            add("ERROR", "BLANK_SLIDE", "Slide has no visible text, picture, or chart", index)

    if len(set(round(size, 1) for size in title_sizes)) > 5:
        add(
            "WARN",
            "TITLE_SCALE_DRIFT",
            "Title sizes vary widely; confirm that section/cover roles justify the differences",
        )

    if len(prs.slides) >= 5 and notes_covered == 0:
        add("WARN", "NO_NOTES", "No speaker notes were found in a multi-slide deck")

    return {
        "file": str(path.resolve()),
        "slide_count": len(prs.slides),
        "slide_size": {
            "width_emu": width,
            "height_emu": height,
            "width_in": round(width / 914400, 3),
            "height_in": round(height / 914400, 3),
            "ratio": round(ratio, 4),
        },
        "ghost_deck": ghost,
        "notes_covered": notes_covered,
        "issues": issues,
    }


def print_human(result: dict[str, Any], show_ghost: bool) -> None:
    size = result["slide_size"]
    print(f"File: {result['file']}")
    print(
        f"Slides: {result['slide_count']} | "
        f"Size: {size['width_in']} × {size['height_in']} in | "
        f"Notes: {result['notes_covered']}/{result['slide_count']}"
    )
    if show_ghost:
        print("\nGhost deck:")
        for entry in result["ghost_deck"]:
            print(f"{entry['slide']:02d}  {entry['title'] or '[no title]'}")
    issues = result["issues"]
    if not issues:
        print("\nNo automated issues found.")
        return
    print("\nIssues:")
    for issue in issues:
        slide = f" slide {issue['slide']}" if "slide" in issue else ""
        print(f"{issue['severity']:5}  {issue['code']}{slide}: {issue['message']}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--ghost", action="store_true", help="Print title-only deck order")
    parser.add_argument("--json", action="store_true", help="Print a JSON report")
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero when any warning or error is present",
    )
    args = parser.parse_args()
    if not args.pptx.exists():
        print(f"ERROR: file not found: {args.pptx}", file=sys.stderr)
        return 2
    try:
        result = audit(args.pptx)
    except Exception as exc:
        print(f"ERROR: could not read presentation: {exc}", file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print_human(result, args.ghost)
    severities = {issue["severity"] for issue in result["issues"]}
    if "ERROR" in severities:
        return 1
    if args.strict and "WARN" in severities:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
