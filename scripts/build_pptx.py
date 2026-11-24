#!/usr/bin/env python3
"""Build a polished, editable 16:9 or 4:3 PPTX from a small JSON deck spec."""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn


THEMES: dict[str, dict[str, str]] = {
    "editorial": {
        "bg": "F7F3EA",
        "ink": "1D1B18",
        "muted": "6B6257",
        "surface": "EEE6D9",
        "accent": "B9472F",
        "on_accent": "FFFFFF",
        "dark": "181A1A",
        "on_dark": "F8F3E9",
        "surface_dark": "252827",
    },
    "business": {
        "bg": "FFFFFF",
        "ink": "172033",
        "muted": "667085",
        "surface": "F2F5FA",
        "accent": "2F6BFF",
        "on_accent": "FFFFFF",
        "dark": "101828",
        "on_dark": "F8FAFC",
        "surface_dark": "1D2939",
    },
    "tech": {
        "bg": "0A0F14",
        "ink": "F3F7F5",
        "muted": "91A29C",
        "surface": "131C22",
        "accent": "78F0AC",
        "on_accent": "07110C",
        "dark": "0A0F14",
        "on_dark": "F3F7F5",
        "surface_dark": "18232A",
    },
    "academic": {
        "bg": "FFFFFF",
        "ink": "1F3242",
        "muted": "667085",
        "surface": "F3F7FA",
        "accent": "2B6B97",
        "on_accent": "FFFFFF",
        "dark": "162A3A",
        "on_dark": "F8FAFC",
        "surface_dark": "213D52",
    },
}

SLIDE_TYPES = {
    "cover",
    "section",
    "agenda",
    "statement",
    "cards",
    "split",
    "comparison",
    "timeline",
    "metrics",
    "chart",
    "quote",
    "closing",
}

SLIDE_WIDTH_EMU = 0
SLIDE_HEIGHT_EMU = 0


def parse_hex(value: str) -> RGBColor:
    value = value.strip().lstrip("#")
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", value):
        raise ValueError(f"Invalid RGB hex color: {value!r}")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def mix_hex(a: str, b: str, amount: float) -> str:
    amount = max(0.0, min(1.0, amount))
    aa = [int(a[i : i + 2], 16) for i in (0, 2, 4)]
    bb = [int(b[i : i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * amount):02X}" for x, y in zip(aa, bb))


def installed_font_families() -> set[str]:
    families: set[str] = set()
    if shutil.which("fc-list"):
        try:
            raw = subprocess.check_output(
                ["fc-list", ":", "family"], text=True, errors="replace", timeout=8
            )
            for line in raw.splitlines():
                for family in line.split(","):
                    if family.strip():
                        families.add(family.strip())
        except (OSError, subprocess.SubprocessError):
            pass

    known_paths = {
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf": "Arial Unicode MS",
        "/System/Library/Fonts/Hiragino Sans GB.ttc": "Hiragino Sans GB",
        "/System/Library/Fonts/STHeiti Medium.ttc": "Heiti SC",
        "/System/Library/Fonts/Supplemental/Songti.ttc": "Songti SC",
    }
    for path, family in known_paths.items():
        if Path(path).exists():
            families.add(family)

    windows_fonts = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
    windows_known = {
        "msyh.ttc": "Microsoft YaHei",
        "msyh.ttf": "Microsoft YaHei",
        "msyhbd.ttc": "Microsoft YaHei",
    }
    for filename, family in windows_known.items():
        if (windows_fonts / filename).exists():
            families.add(family)
    return families


def choose_font(candidates: list[str], available: set[str]) -> str:
    for candidate in candidates:
        if not available or candidate in available:
            return candidate
    return candidates[-1]


def font_plan(language: str, overrides: dict[str, str] | None = None) -> dict[str, str]:
    available = installed_font_families()
    if language.lower().startswith(("zh", "ja", "ko")):
        ea_default = choose_font(
            [
                "Microsoft YaHei",
                "Hiragino Sans GB",
                "Noto Sans CJK SC",
                "Noto Sans SC",
                "Arial Unicode MS",
                "Heiti SC",
                "PingFang SC",
            ],
            available,
        )
    else:
        ea_default = choose_font(["Arial", "Arial Unicode MS"], available)

    plan = {
        "heading_latin": "Cambria",
        "body_latin": "Arial",
        "heading_ea": ea_default,
        "body_ea": ea_default,
    }
    if overrides:
        for key in plan:
            if overrides.get(key):
                plan[key] = str(overrides[key])
    return plan


def display_units(text: str) -> float:
    units = 0.0
    for ch in text:
        if ch.isspace():
            units += 0.35
        elif ord(ch) > 255:
            units += 1.0
        elif ch.isupper():
            units += 0.67
        else:
            units += 0.54
    return units


def balanced_text_break(text: str, threshold: float = 24.0) -> str:
    text = str(text).strip()
    if "\n" in text or display_units(text) <= threshold:
        return text
    target = len(text) / 2
    candidates = [
        index + 1
        for index, char in enumerate(text)
        if char in "，。；：,;"
    ]
    if not candidates:
        return text
    split = min(candidates, key=lambda index: abs(index - target))
    if split <= 4 or len(text) - split <= 4:
        return text
    return text[:split].rstrip() + "\n" + text[split:].lstrip()


def is_cjk_character(char: str) -> bool:
    code = ord(char)
    return (
        0x3400 <= code <= 0x4DBF
        or 0x4E00 <= code <= 0x9FFF
        or 0xF900 <= code <= 0xFAFF
        or 0x3040 <= code <= 0x30FF
        or 0xAC00 <= code <= 0xD7AF
        or 0x3000 <= code <= 0x303F
        or 0xFF00 <= code <= 0xFFEF
    )


def font_segments(text: str) -> list[tuple[str, bool]]:
    if not text:
        return []
    segments: list[tuple[str, bool]] = []
    buffer: list[str] = []
    current_kind = is_cjk_character(text[0])
    for char in text:
        kind = is_cjk_character(char)
        if buffer and kind != current_kind:
            segments.append(("".join(buffer), current_kind))
            buffer = []
        buffer.append(char)
        current_kind = kind
    if buffer:
        segments.append(("".join(buffer), current_kind))
    return segments


def fitted_size(
    text: str,
    *,
    base: float,
    minimum: float,
    one_line_units: float,
    scale: float,
    line_factor: float = 1.0,
) -> float:
    units = display_units(str(text))
    size = base if units <= one_line_units else base * one_line_units / units
    size *= max(0.84, min(1.0, 1.2 - 0.08 * line_factor))
    return round(max(minimum, size) * scale, 1)


def set_margins(shape: Any, left: float = 0.0, right: float = 0.0, top: float = 0.0, bottom: float = 0.0) -> None:
    shape.text_frame.margin_left = Inches(left)
    shape.text_frame.margin_right = Inches(right)
    shape.text_frame.margin_top = Inches(top)
    shape.text_frame.margin_bottom = Inches(bottom)


def set_run_font(
    run: Any,
    *,
    font_name: str,
    ea_font: str,
    size: float,
    color: str,
    bold: bool = False,
    italic: bool = False,
) -> None:
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = parse_hex(color)
    r_pr = run._r.get_or_add_rPr()
    ea = r_pr.find(qn("a:ea"))
    if ea is None:
        ea = OxmlElement("a:ea")
        r_pr.append(ea)
    ea.set("typeface", ea_font)


def add_markdown_runs(
    paragraph: Any,
    text: str,
    *,
    latin_font: str,
    ea_font: str,
    size: float,
    color: str,
    bold: bool = False,
) -> None:
    parts = re.split(r"(\*\*.*?\*\*)", str(text))
    for part in parts:
        if not part:
            continue
        is_bold = part.startswith("**") and part.endswith("**") and len(part) >= 4
        visible = part[2:-2] if is_bold else part
        for segment, has_cjk in font_segments(visible):
            run = paragraph.add_run()
            run.text = segment
            set_run_font(
                run,
                font_name=ea_font if has_cjk else latin_font,
                ea_font=ea_font,
                size=size,
                color=color,
                bold=bold or is_bold,
            )


def add_text(
    slide: Any,
    text: str,
    *,
    x: float,
    y: float,
    w: float,
    h: float,
    size: float,
    color: str,
    fonts: dict[str, str],
    use_heading: bool = False,
    bold: bool = False,
    italic: bool = False,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    valign: MSO_ANCHOR = MSO_ANCHOR.TOP,
    name: str | None = None,
) -> Any:
    size = max(8.5, float(size))
    left = int(SLIDE_WIDTH_EMU * x)
    top = int(SLIDE_HEIGHT_EMU * y)
    width = int(SLIDE_WIDTH_EMU * w)
    height = int(SLIDE_HEIGHT_EMU * h)
    box = slide.shapes.add_textbox(left, top, width, height)
    if name:
        box.name = name
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = valign
    set_margins(box)
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_after = Pt(0)
    p.line_spacing = 1.04
    latin = fonts["heading_latin"] if use_heading else fonts["body_latin"]
    ea = fonts["heading_ea"] if use_heading else fonts["body_ea"]
    add_markdown_runs(
        p,
        str(text),
        latin_font=latin,
        ea_font=ea,
        size=size,
        color=color,
        bold=bold,
    )
    if italic:
        for run in p.runs:
            run.font.italic = True
    return box


def add_bullets(
    slide: Any,
    items: list[Any],
    *,
    x: float,
    y: float,
    w: float,
    h: float,
    size: float,
    color: str,
    muted: str,
    fonts: dict[str, str],
    space_after: float = 8,
) -> Any:
    size = max(8.5, float(size))
    left = int(SLIDE_WIDTH_EMU * x)
    top = int(SLIDE_HEIGHT_EMU * y)
    width = int(SLIDE_WIDTH_EMU * w)
    height = int(SLIDE_HEIGHT_EMU * h)
    box = slide.shapes.add_textbox(left, top, width, height)
    box.name = "Body"
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = MSO_ANCHOR.TOP
    set_margins(box)
    for index, raw in enumerate(items):
        p = tf.paragraphs[0] if index == 0 else tf.add_paragraph()
        p.space_after = Pt(space_after)
        p.line_spacing = 1.08
        add_markdown_runs(
            p,
            "•  " + str(raw),
            latin_font=fonts["body_latin"],
            ea_font=fonts["body_ea"],
            size=size,
            color=color,
        )
    return box


def add_rect(
    slide: Any,
    *,
    x: float,
    y: float,
    w: float,
    h: float,
    fill: str,
    radius: bool = False,
    name: str | None = None,
) -> Any:
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    left = int(SLIDE_WIDTH_EMU * x)
    top = int(SLIDE_HEIGHT_EMU * y)
    width = int(SLIDE_WIDTH_EMU * w)
    height = int(SLIDE_HEIGHT_EMU * h)
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    if name:
        shape.name = name
    shape.fill.solid()
    shape.fill.fore_color.rgb = parse_hex(fill)
    shape.line.fill.background()
    shape.shadow.inherit = False
    if radius:
        try:
            shape.adjustments[0] = 0.08
        except (IndexError, ValueError):
            pass
    return shape


def add_rule(slide: Any, *, x: float, y: float, w: float, color: str, thickness: float = 0.002) -> Any:
    return add_rect(slide, x=x, y=y, w=w, h=thickness, fill=color, name="Rule")


def fill_background(slide: Any, color: str) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = parse_hex(color)


def resolve_asset(base_dir: Path, value: str) -> Path:
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = base_dir / path
    path = path.resolve()
    if not path.exists():
        raise FileNotFoundError(f"Image not found: {path}")
    return path


def add_image(
    slide: Any,
    path: Path,
    *,
    x: float,
    y: float,
    w: float,
    h: float,
    fit: str,
    backing: str,
) -> Any:
    slide_w = SLIDE_WIDTH_EMU
    slide_h = SLIDE_HEIGHT_EMU
    box_left = int(slide_w * x)
    box_top = int(slide_h * y)
    box_w = int(slide_w * w)
    box_h = int(slide_h * h)
    with Image.open(path) as image:
        image_w, image_h = image.size
    if image_w <= 0 or image_h <= 0:
        raise ValueError(f"Invalid image size: {path}")

    if fit == "contain":
        add_rect(slide, x=x, y=y, w=w, h=h, fill=backing, name="Image backing")
        source_ratio = image_w / image_h
        target_ratio = box_w / box_h
        if source_ratio >= target_ratio:
            image_width = box_w
            image_height = round(box_w / source_ratio)
        else:
            image_height = box_h
            image_width = round(box_h * source_ratio)
        left = box_left + (box_w - image_width) // 2
        top = box_top + (box_h - image_height) // 2
        return slide.shapes.add_picture(str(path), left, top, image_width, image_height)

    picture = slide.shapes.add_picture(str(path), box_left, box_top, box_w, box_h)
    source_ratio = image_w / image_h
    target_ratio = box_w / box_h
    if source_ratio > target_ratio:
        crop = (1.0 - target_ratio / source_ratio) / 2.0
        picture.crop_left = crop
        picture.crop_right = crop
    elif source_ratio < target_ratio:
        crop = (1.0 - source_ratio / target_ratio) / 2.0
        picture.crop_top = crop
        picture.crop_bottom = crop
    return picture


def add_notes(slide: Any, text: str | None) -> None:
    if text:
        slide.notes_slide.notes_text_frame.text = str(text).strip()


def style_chart(
    chart: Any,
    *,
    theme: dict[str, str],
    fonts: dict[str, str],
    font_scale: float,
) -> None:
    chart.font.name = fonts["body_latin"]
    chart.font.size = Pt(10 * font_scale)
    chart.font.color.rgb = parse_hex(theme["muted"])
    try:
        chart.value_axis.has_major_gridlines = True
        chart.value_axis.major_gridlines.format.line.color.rgb = parse_hex(
            mix_hex(theme["muted"], theme["bg"], 0.72)
        )
        chart.value_axis.major_gridlines.format.line.width = Pt(0.5)
    except Exception:
        pass
    try:
        chart.category_axis.tick_labels.font.name = fonts["body_latin"]
        chart.category_axis.tick_labels.font.size = Pt(10 * font_scale)
        chart.category_axis.tick_labels.font.color.rgb = parse_hex(theme["muted"])
        chart.value_axis.tick_labels.font.name = fonts["body_latin"]
        chart.value_axis.tick_labels.font.size = Pt(10 * font_scale)
        chart.value_axis.tick_labels.font.color.rgb = parse_hex(theme["muted"])
    except Exception:
        pass


def set_series_colors(
    plot: Any,
    *,
    theme: dict[str, str],
    is_pie: bool = False,
    is_line: bool = False,
) -> None:
    accent = theme["accent"]
    colors = [
        accent,
        mix_hex(accent, theme["ink"], 0.35),
        mix_hex(accent, theme["bg"], 0.35),
        mix_hex(theme["muted"], theme["bg"], 0.2),
    ]
    for index, series in enumerate(plot.series):
        if is_pie:
            try:
                for point_index, point in enumerate(series.points):
                    point.format.fill.solid()
                    point.format.fill.fore_color.rgb = parse_hex(
                        colors[point_index % len(colors)]
                    )
            except Exception:
                pass
        else:
            color = colors[index % len(colors)]
            if is_line:
                try:
                    series.format.line.color.rgb = parse_hex(color)
                    series.format.line.width = Pt(2.2)
                    continue
                except Exception:
                    pass
            try:
                series.format.fill.solid()
                series.format.fill.fore_color.rgb = parse_hex(color)
            except Exception:
                pass


def page_palette(theme: dict[str, str], tone: str = "light") -> dict[str, str]:
    if tone == "dark":
        return {
            "bg": theme["dark"],
            "ink": theme["on_dark"],
            "muted": mix_hex(theme["on_dark"], theme["dark"], 0.42),
            "surface": theme["surface_dark"],
        }
    return {
        "bg": theme["bg"],
        "ink": theme["ink"],
        "muted": theme["muted"],
        "surface": theme["surface"],
    }


def add_footer(
    slide: Any,
    *,
    spec: dict[str, Any],
    palette: dict[str, str],
    fonts: dict[str, str],
    page_number: int,
    total: int,
    font_scale: float,
    show_number: bool = True,
) -> None:
    source = str(spec.get("source", "")).strip()
    if source:
        add_text(
            slide,
            source,
            x=0.06,
            y=0.935,
            w=0.72,
            h=0.035,
            size=8.5 * font_scale,
            color=palette["muted"],
            fonts=fonts,
            name="Footer",
        )
    if show_number:
        add_text(
            slide,
            f"{page_number:02d} / {total:02d}",
            x=0.84,
            y=0.934,
            w=0.10,
            h=0.035,
            size=8.5 * font_scale,
            color=palette["muted"],
            fonts=fonts,
            align=PP_ALIGN.RIGHT,
            name="Footer",
        )


def add_page_title(
    slide: Any,
    *,
    title: str,
    kicker: str,
    palette: dict[str, str],
    fonts: dict[str, str],
    font_scale: float,
) -> None:
    if kicker:
        add_text(
            slide,
            kicker.upper(),
            x=0.06,
            y=0.055,
            w=0.70,
            h=0.035,
            size=9 * font_scale,
            color=palette["muted"],
            fonts=fonts,
        )
    size = fitted_size(
        title,
        base=29 * font_scale,
        minimum=21 * font_scale,
        one_line_units=48,
        scale=1.0,
    )
    add_text(
        slide,
        title,
        x=0.06,
        y=0.09 if kicker else 0.065,
        w=0.86,
        h=0.105,
        size=size,
        color=palette["ink"],
        fonts=fonts,
        use_heading=True,
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
        name="Title",
    )


def section_items(value: Any) -> tuple[str, list[Any]]:
    if isinstance(value, dict):
        return str(value.get("heading", "")), list(value.get("items", []))
    if isinstance(value, list):
        return "", list(value)
    if value is None:
        return "", []
    return "", [str(value)]


def layout_cover(
    slide: Any,
    spec: dict[str, Any],
    deck: dict[str, Any],
    theme: dict[str, str],
    fonts: dict[str, str],
    font_scale: float,
) -> dict[str, str]:
    palette = page_palette(theme, "light")
    fill_background(slide, palette["bg"])
    panel = theme["accent"]
    add_rect(slide, x=0.70, y=0.0, w=0.30, h=1.0, fill=panel, name="Cover panel")
    add_text(
        slide,
        str(spec.get("eyebrow") or deck.get("subtitle") or "PRESENTATION").upper(),
        x=0.06,
        y=0.075,
        w=0.55,
        h=0.04,
        size=9 * font_scale,
        color=theme["accent"],
        fonts=fonts,
    )
    title = str(spec.get("title") or deck.get("title") or "Untitled presentation")
    title_size = fitted_size(
        title,
        base=46 * font_scale,
        minimum=29 * font_scale,
        one_line_units=30,
        scale=1.0,
        line_factor=1.2,
    )
    add_text(
        slide,
        title,
        x=0.06,
        y=0.18,
        w=0.58,
        h=0.37,
        size=title_size,
        color=palette["ink"],
        fonts=fonts,
        use_heading=True,
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
        name="Title",
    )
    subtitle = str(spec.get("subtitle") or deck.get("subtitle") or "")
    if subtitle:
        add_text(
            slide,
            subtitle,
            x=0.06,
            y=0.61,
            w=0.55,
            h=0.12,
            size=16 * font_scale,
            color=palette["muted"],
            fonts=fonts,
            valign=MSO_ANCHOR.TOP,
        )
    meta_parts = [
        str(spec.get("meta", "")).strip(),
        str(deck.get("author", "")).strip(),
        str(deck.get("date", "")).strip(),
    ]
    meta = "  ·  ".join(part for part in meta_parts if part)
    if meta:
        add_text(
            slide,
            meta,
            x=0.06,
            y=0.84,
            w=0.56,
            h=0.05,
            size=9 * font_scale,
            color=palette["muted"],
            fonts=fonts,
        )
    add_text(
        slide,
        f"{len(deck.get('slides', [])):02d}",
        x=0.74,
        y=0.31,
        w=0.20,
        h=0.23,
        size=72 * font_scale,
        color=theme["on_accent"],
        fonts=fonts,
        use_heading=True,
        bold=True,
        align=PP_ALIGN.RIGHT,
        name="Slide count",
    )
    add_text(
        slide,
        "SLIDES",
        x=0.74,
        y=0.57,
        w=0.20,
        h=0.04,
        size=9 * font_scale,
        color=theme["on_accent"],
        fonts=fonts,
        align=PP_ALIGN.RIGHT,
    )
    return palette


def layout_section(
    slide: Any,
    spec: dict[str, Any],
    deck: dict[str, Any],
    theme: dict[str, str],
    fonts: dict[str, str],
    font_scale: float,
) -> dict[str, str]:
    palette = {
        "bg": theme["accent"],
        "ink": theme["on_accent"],
        "muted": mix_hex(theme["on_accent"], theme["accent"], 0.28),
        "surface": mix_hex(theme["accent"], theme["on_accent"], 0.10),
    }
    fill_background(slide, palette["bg"])
    number = str(spec.get("number") or spec.get("kicker") or "")
    if number:
        add_text(
            slide,
            number,
            x=0.67,
            y=0.08,
            w=0.27,
            h=0.50,
            size=126 * font_scale,
            color=mix_hex(theme["accent"], theme["on_accent"], 0.35),
            fonts=fonts,
            use_heading=True,
            bold=False,
            align=PP_ALIGN.RIGHT,
            name="Section number",
        )
    title = str(spec.get("title", "Section"))
    add_text(
        slide,
        title,
        x=0.07,
        y=0.33,
        w=0.68,
        h=0.22,
        size=fitted_size(
            title,
            base=48 * font_scale,
            minimum=28 * font_scale,
            one_line_units=26,
            scale=1.0,
            line_factor=1.2,
        ),
        color=palette["ink"],
        fonts=fonts,
        use_heading=True,
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
        name="Title",
    )
    subtitle = str(spec.get("subtitle", ""))
    if subtitle:
        add_text(
            slide,
            subtitle,
            x=0.07,
            y=0.60,
            w=0.57,
            h=0.12,
            size=16 * font_scale,
            color=palette["muted"],
            fonts=fonts,
        )
    return palette


def layout_agenda(
    slide: Any,
    spec: dict[str, Any],
    deck: dict[str, Any],
    theme: dict[str, str],
    fonts: dict[str, str],
    font_scale: float,
) -> dict[str, str]:
    palette = page_palette(theme, str(spec.get("tone", "light")))
    fill_background(slide, palette["bg"])
    title = str(spec.get("title", "Agenda"))
    add_text(
        slide,
        title,
        x=0.06,
        y=0.12,
        w=0.28,
        h=0.30,
        size=fitted_size(title, base=36 * font_scale, minimum=25 * font_scale, one_line_units=18, scale=1.0),
        color=palette["ink"],
        fonts=fonts,
        use_heading=True,
        bold=True,
        name="Title",
    )
    items = list(spec.get("items", []))
    if len(items) > 6:
        raise ValueError("Agenda supports at most 6 items")
    row_top = 0.11
    row_h = 0.12 if len(items) <= 4 else 0.095
    for index, item in enumerate(items):
        if isinstance(item, dict):
            label = str(item.get("title") or item.get("label") or "")
            body = str(item.get("body") or item.get("description") or "")
        else:
            label = str(item)
            body = ""
        y = row_top + index * row_h
        add_text(
            slide,
            f"{index + 1:02d}",
            x=0.40,
            y=y,
            w=0.06,
            h=0.05,
            size=10 * font_scale,
            color=theme["accent"],
            fonts=fonts,
            bold=True,
        )
        add_text(
            slide,
            label,
            x=0.49,
            y=y - 0.008,
            w=0.41,
            h=0.05,
            size=fitted_size(label, base=18 * font_scale, minimum=13 * font_scale, one_line_units=44, scale=1.0),
            color=palette["ink"],
            fonts=fonts,
            use_heading=True,
            bold=True,
        )
        if body:
            add_text(
                slide,
                body,
                x=0.49,
                y=y + 0.042,
                w=0.41,
                h=0.035,
                size=10.5 * font_scale,
                color=palette["muted"],
                fonts=fonts,
            )
        if index < len(items) - 1:
            add_rule(
                slide,
                x=0.40,
                y=y + row_h - 0.014,
                w=0.50,
                color=mix_hex(palette["muted"], palette["bg"], 0.73),
            )
    return palette


def layout_statement(
    slide: Any,
    spec: dict[str, Any],
    deck: dict[str, Any],
    theme: dict[str, str],
    fonts: dict[str, str],
    font_scale: float,
) -> dict[str, str]:
    palette = page_palette(theme, str(spec.get("tone", "light")))
    fill_background(slide, palette["bg"])
    kicker = str(spec.get("kicker", ""))
    if kicker:
        add_text(
            slide,
            kicker.upper(),
            x=0.07,
            y=0.10,
            w=0.55,
            h=0.04,
            size=9 * font_scale,
            color=theme["accent"],
            fonts=fonts,
        )
    title = str(spec.get("title", ""))
    add_text(
        slide,
        title,
        x=0.07,
        y=0.23,
        w=0.78,
        h=0.36,
        size=fitted_size(
            title,
            base=46 * font_scale,
            minimum=28 * font_scale,
            one_line_units=25,
            scale=1.0,
            line_factor=1.1,
        ),
        color=palette["ink"],
        fonts=fonts,
        use_heading=True,
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
        name="Title",
    )
    body = str(spec.get("body", ""))
    if body:
        add_text(
            slide,
            body,
            x=0.07,
            y=0.64,
            w=0.63,
            h=0.13,
            size=fitted_size(body, base=16 * font_scale, minimum=12 * font_scale, one_line_units=105, scale=1.0),
            color=palette["muted"],
            fonts=fonts,
        )
    return palette


def layout_cards(
    slide: Any,
    spec: dict[str, Any],
    deck: dict[str, Any],
    theme: dict[str, str],
    fonts: dict[str, str],
    font_scale: float,
) -> dict[str, str]:
    palette = page_palette(theme, str(spec.get("tone", "light")))
    fill_background(slide, palette["bg"])
    add_page_title(
        slide,
        title=str(spec.get("title", "")),
        kicker=str(spec.get("kicker", "")),
        palette=palette,
        fonts=fonts,
        font_scale=font_scale,
    )
    cards = list(spec.get("cards", []))
    if not 1 <= len(cards) <= 6:
        raise ValueError("cards supports 1–6 items")
    if len(cards) <= 4:
        cols = len(cards)
        rows = 1
    else:
        cols = 3
        rows = 2
    gap = 0.025
    x0 = 0.06
    y0 = 0.24
    area_w = 0.88
    area_h = 0.58
    card_w = (area_w - gap * (cols - 1)) / cols
    card_h = (area_h - gap * (rows - 1)) / rows
    for index, card in enumerate(cards):
        row = index // cols
        col = index % cols
        x = x0 + col * (card_w + gap)
        y = y0 + row * (card_h + gap)
        add_rect(slide, x=x, y=y, w=card_w, h=card_h, fill=palette["surface"], radius=True, name="Card")
        label = str(card.get("label") or f"{index + 1:02d}")
        title_text = str(card.get("title") or "")
        body_text = str(card.get("body") or "")
        add_text(
            slide,
            label,
            x=x + 0.025,
            y=y + 0.035,
            w=max(0.03, card_w - 0.05),
            h=0.035,
            size=9.5 * font_scale,
            color=theme["accent"],
            fonts=fonts,
            bold=True,
        )
        add_text(
            slide,
            title_text,
            x=x + 0.025,
            y=y + 0.09,
            w=card_w - 0.05,
            h=0.10,
            size=fitted_size(
                title_text,
                base=18 * font_scale,
                minimum=13 * font_scale,
                one_line_units=max(14, card_w * 66),
                scale=1.0,
            ),
            color=palette["ink"],
            fonts=fonts,
            use_heading=True,
            bold=True,
        )
        if body_text:
            add_text(
                slide,
                body_text,
                x=x + 0.025,
                y=y + 0.24,
                w=card_w - 0.05,
                h=max(0.08, card_h - 0.27),
                size=fitted_size(
                    body_text,
                    base=13.5 * font_scale,
                    minimum=10.5 * font_scale,
                    one_line_units=max(30, card_w * 72),
                    scale=1.0,
                ),
                color=palette["muted"],
                fonts=fonts,
            )
    return palette


def layout_split(
    slide: Any,
    spec: dict[str, Any],
    deck: dict[str, Any],
    theme: dict[str, str],
    fonts: dict[str, str],
    font_scale: float,
    base_dir: Path,
) -> dict[str, str]:
    palette = page_palette(theme, str(spec.get("tone", "light")))
    fill_background(slide, palette["bg"])
    add_page_title(
        slide,
        title=str(spec.get("title", "")),
        kicker=str(spec.get("kicker", "")),
        palette=palette,
        fonts=fonts,
        font_scale=font_scale,
    )
    left_heading, left_items = section_items(spec.get("left"))
    right_heading, right_items = section_items(spec.get("right"))
    add_text(
        slide,
        left_heading.upper(),
        x=0.06,
        y=0.235,
        w=0.39,
        h=0.04,
        size=10 * font_scale,
        color=theme["accent"],
        fonts=fonts,
        bold=True,
    )
    add_bullets(
        slide,
        left_items,
        x=0.06,
        y=0.29,
        w=0.39,
        h=0.50,
        size=fitted_size(
            "\n".join(map(str, left_items)),
            base=16 * font_scale,
            minimum=12 * font_scale,
            one_line_units=80,
            scale=1.0,
        ),
        color=palette["ink"],
        muted=palette["muted"],
        fonts=fonts,
    )
    image_value = str(spec.get("image", "")).strip()
    if image_value:
        image_path = resolve_asset(base_dir, image_value)
        add_image(
            slide,
            image_path,
            x=0.51,
            y=0.235,
            w=0.43,
            h=0.53,
            fit=str(spec.get("image_fit", "cover")),
            backing=palette["surface"],
        )
        caption = str(spec.get("image_caption", ""))
        if caption:
            add_text(
                slide,
                caption,
                x=0.51,
                y=0.79,
                w=0.43,
                h=0.04,
                size=9 * font_scale,
                color=palette["muted"],
                fonts=fonts,
            )
    else:
        add_rect(slide, x=0.51, y=0.225, w=0.43, h=0.57, fill=palette["surface"], radius=True, name="Content panel")
        add_text(
            slide,
            right_heading.upper(),
            x=0.54,
            y=0.255,
            w=0.37,
            h=0.04,
            size=10 * font_scale,
            color=theme["accent"],
            fonts=fonts,
            bold=True,
        )
        add_bullets(
            slide,
            right_items,
            x=0.54,
            y=0.31,
            w=0.37,
            h=0.44,
            size=fitted_size(
                "\n".join(map(str, right_items)),
                base=15 * font_scale,
                minimum=11.5 * font_scale,
                one_line_units=75,
                scale=1.0,
            ),
            color=palette["ink"],
            muted=palette["muted"],
            fonts=fonts,
        )
    return palette


def layout_comparison(
    slide: Any,
    spec: dict[str, Any],
    deck: dict[str, Any],
    theme: dict[str, str],
    fonts: dict[str, str],
    font_scale: float,
) -> dict[str, str]:
    palette = page_palette(theme, str(spec.get("tone", "light")))
    fill_background(slide, palette["bg"])
    add_page_title(
        slide,
        title=str(spec.get("title", "")),
        kicker=str(spec.get("kicker", "")),
        palette=palette,
        fonts=fonts,
        font_scale=font_scale,
    )
    left_heading, left_items = section_items(spec.get("left"))
    right_heading, right_items = section_items(spec.get("right"))
    add_text(
        slide,
        left_heading,
        x=0.06,
        y=0.235,
        w=0.39,
        h=0.06,
        size=20 * font_scale,
        color=palette["ink"],
        fonts=fonts,
        use_heading=True,
        bold=True,
    )
    add_text(
        slide,
        right_heading,
        x=0.55,
        y=0.235,
        w=0.39,
        h=0.06,
        size=20 * font_scale,
        color=theme["accent"],
        fonts=fonts,
        use_heading=True,
        bold=True,
    )
    add_bullets(
        slide,
        left_items,
        x=0.06,
        y=0.34,
        w=0.39,
        h=0.39,
        size=fitted_size(
            "\n".join(map(str, left_items)),
            base=15 * font_scale,
            minimum=11 * font_scale,
            one_line_units=70,
            scale=1.0,
        ),
        color=palette["ink"],
        muted=palette["muted"],
        fonts=fonts,
    )
    add_bullets(
        slide,
        right_items,
        x=0.55,
        y=0.34,
        w=0.39,
        h=0.39,
        size=fitted_size(
            "\n".join(map(str, right_items)),
            base=15 * font_scale,
            minimum=11 * font_scale,
            one_line_units=70,
            scale=1.0,
        ),
        color=palette["ink"],
        muted=palette["muted"],
        fonts=fonts,
    )
    add_rect(
        slide,
        x=0.495,
        y=0.23,
        w=0.002,
        h=0.52,
        fill=mix_hex(palette["muted"], palette["bg"], 0.76),
        name="Comparison divider",
    )
    takeaway = str(spec.get("body") or spec.get("takeaway") or "")
    if takeaway:
        add_rect(slide, x=0.06, y=0.76, w=0.88, h=0.11, fill=palette["surface"], name="Takeaway band")
        add_text(
            slide,
            takeaway,
            x=0.08,
            y=0.777,
            w=0.84,
            h=0.07,
            size=fitted_size(takeaway, base=14 * font_scale, minimum=11 * font_scale, one_line_units=120, scale=1.0),
            color=palette["ink"],
            fonts=fonts,
            bold=True,
            valign=MSO_ANCHOR.MIDDLE,
        )
    return palette


def layout_timeline(
    slide: Any,
    spec: dict[str, Any],
    deck: dict[str, Any],
    theme: dict[str, str],
    fonts: dict[str, str],
    font_scale: float,
) -> dict[str, str]:
    palette = page_palette(theme, str(spec.get("tone", "light")))
    fill_background(slide, palette["bg"])
    add_page_title(
        slide,
        title=str(spec.get("title", "")),
        kicker=str(spec.get("kicker", "")),
        palette=palette,
        fonts=fonts,
        font_scale=font_scale,
    )
    steps = list(spec.get("steps", []))
    if not 2 <= len(steps) <= 5:
        raise ValueError("timeline supports 2–5 steps")
    x0, x1 = 0.08, 0.86
    line_y = 0.39
    add_rect(slide, x=x0, y=line_y, w=x1 - x0, h=0.002, fill=mix_hex(palette["muted"], palette["bg"], 0.72), name="Timeline")
    slot = (x1 - x0) / max(1, len(steps) - 1)
    for index, step in enumerate(steps):
        cx = x0 + slot * index
        marker = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            int(SLIDE_WIDTH_EMU * cx - Inches(0.12)),
            int(SLIDE_HEIGHT_EMU * line_y - Inches(0.12)),
            Inches(0.24),
            Inches(0.24),
        )
        marker.fill.solid()
        marker.fill.fore_color.rgb = parse_hex(theme["accent"])
        marker.line.fill.background()
        marker.shadow.inherit = False
        marker.name = f"Timeline marker {index + 1}"
        add_text(
            slide,
            str(step.get("label") or f"{index + 1:02d}"),
            x=cx - 0.06,
            y=line_y - 0.017,
            w=0.12,
            h=0.034,
            size=max(8.5, 9 * font_scale),
            color=theme["on_accent"],
            fonts=fonts,
            bold=True,
            align=PP_ALIGN.CENTER,
            valign=MSO_ANCHOR.MIDDLE,
        )
        step_title = str(step.get("title") or "")
        step_body = str(step.get("body") or "")
        col_w = 0.19
        add_text(
            slide,
            step_title,
            x=max(0.05, cx - col_w / 2),
            y=0.52,
            w=col_w,
            h=0.08,
            size=fitted_size(step_title, base=15 * font_scale, minimum=11.5 * font_scale, one_line_units=18, scale=1.0),
            color=palette["ink"],
            fonts=fonts,
            use_heading=True,
            bold=True,
            align=PP_ALIGN.CENTER,
        )
        if step_body:
            add_text(
                slide,
                step_body,
                x=max(0.05, cx - col_w / 2),
                y=0.62,
                w=col_w,
                h=0.16,
                size=fitted_size(step_body, base=11.5 * font_scale, minimum=9.5 * font_scale, one_line_units=44, scale=1.0),
                color=palette["muted"],
                fonts=fonts,
                align=PP_ALIGN.CENTER,
            )
    return palette


def layout_metrics(
    slide: Any,
    spec: dict[str, Any],
    deck: dict[str, Any],
    theme: dict[str, str],
    fonts: dict[str, str],
    font_scale: float,
) -> dict[str, str]:
    palette = page_palette(theme, str(spec.get("tone", "light")))
    fill_background(slide, palette["bg"])
    add_page_title(
        slide,
        title=str(spec.get("title", "")),
        kicker=str(spec.get("kicker", "")),
        palette=palette,
        fonts=fonts,
        font_scale=font_scale,
    )
    metrics = list(spec.get("metrics", []))
    if not 2 <= len(metrics) <= 4:
        raise ValueError("metrics supports 2–4 items")
    gap = 0.03
    x0 = 0.07
    width = (0.86 - gap * (len(metrics) - 1)) / len(metrics)
    for index, metric in enumerate(metrics):
        x = x0 + index * (width + gap)
        add_rect(slide, x=x, y=0.25, w=width, h=0.48, fill=palette["surface"], radius=True, name="Metric")
        value = str(metric.get("value") or "")
        label = str(metric.get("label") or "")
        note = str(metric.get("note") or "")
        add_text(
            slide,
            value,
            x=x + 0.025,
            y=0.31,
            w=width - 0.05,
            h=0.16,
            size=fitted_size(value, base=48 * font_scale, minimum=27 * font_scale, one_line_units=9, scale=1.0),
            color=theme["accent"],
            fonts=fonts,
            use_heading=True,
            bold=True,
            valign=MSO_ANCHOR.MIDDLE,
        )
        add_text(
            slide,
            label,
            x=x + 0.025,
            y=0.50,
            w=width - 0.05,
            h=0.07,
            size=fitted_size(label, base=14 * font_scale, minimum=10.5 * font_scale, one_line_units=max(18, width * 55), scale=1.0),
            color=palette["ink"],
            fonts=fonts,
            bold=True,
        )
        if note:
            add_text(
                slide,
                note,
                x=x + 0.025,
                y=0.58,
                w=width - 0.05,
                h=0.11,
                size=fitted_size(note, base=10.5 * font_scale, minimum=8.5 * font_scale, one_line_units=max(26, width * 50), scale=1.0),
                color=palette["muted"],
                fonts=fonts,
            )
    return palette


def layout_chart(
    slide: Any,
    spec: dict[str, Any],
    deck: dict[str, Any],
    theme: dict[str, str],
    fonts: dict[str, str],
    font_scale: float,
) -> dict[str, str]:
    palette = page_palette(theme, str(spec.get("tone", "light")))
    fill_background(slide, palette["bg"])
    add_page_title(
        slide,
        title=str(spec.get("title", "")),
        kicker=str(spec.get("kicker", "")),
        palette=palette,
        fonts=fonts,
        font_scale=font_scale,
    )
    chart_spec = dict(spec.get("chart", {}))
    categories = list(chart_spec.get("categories", []))
    series_specs = list(chart_spec.get("series", []))
    if not categories or not series_specs:
        raise ValueError("chart requires categories and at least one series")

    chart_data = CategoryChartData()
    chart_data.categories = [str(item) for item in categories]
    for series in series_specs:
        chart_data.add_series(str(series.get("name", "Series")), list(series.get("values", [])))

    chart_type_name = str(chart_spec.get("chart_type", "column")).lower()
    chart_types = {
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "bar": XL_CHART_TYPE.BAR_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE,
    }
    if chart_type_name not in chart_types:
        raise ValueError("chart_type must be column, bar, line, or pie")
    is_pie = chart_type_name == "pie"
    chart_x, chart_y, chart_w, chart_h = 0.06, 0.24, 0.62, 0.57
    chart = slide.shapes.add_chart(
        chart_types[chart_type_name],
        int(SLIDE_WIDTH_EMU * chart_x),
        int(SLIDE_HEIGHT_EMU * chart_y),
        int(SLIDE_WIDTH_EMU * chart_w),
        int(SLIDE_HEIGHT_EMU * chart_h),
        chart_data,
    ).chart
    chart.has_title = False
    chart.has_legend = len(series_specs) > 1 or is_pie
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.name = fonts["body_latin"]
        chart.legend.font.size = Pt(9.5 * font_scale)
        chart.legend.font.color.rgb = parse_hex(palette["muted"])
    plot = chart.plots[0]
    plot.has_data_labels = False if is_pie else True
    if not is_pie:
        try:
            plot.data_labels.show_value = True
            plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
            plot.data_labels.font.size = Pt(9 * font_scale)
            plot.data_labels.font.color.rgb = parse_hex(palette["ink"])
        except Exception:
            pass
    need_axes = not is_pie
    try:
        chart.value_axis.has_major_gridlines = False
        chart.category_axis.has_major_gridlines = False
        chart.value_axis.visible = need_axes
        chart.category_axis.visible = need_axes
    except Exception:
        pass
    style_chart(chart, theme=palette, fonts=fonts, font_scale=font_scale)
    set_series_colors(
        plot,
        theme=theme,
        is_pie=is_pie,
        is_line=chart_type_name == "line",
    )

    unit = str(chart_spec.get("unit", "")).strip()
    if unit:
        add_text(
            slide,
            unit,
            x=0.61,
            y=0.205,
            w=0.07,
            h=0.03,
            size=8.5 * font_scale,
            color=palette["muted"],
            fonts=fonts,
            align=PP_ALIGN.RIGHT,
        )

    commentary = spec.get("commentary", [])
    if isinstance(commentary, str):
        commentary = [commentary]
    add_text(
        slide,
        "WHAT IT MEANS",
        x=0.72,
        y=0.25,
        w=0.20,
        h=0.035,
        size=9 * font_scale,
        color=theme["accent"],
        fonts=fonts,
        bold=True,
    )
    add_bullets(
        slide,
        list(commentary),
        x=0.72,
        y=0.31,
        w=0.20,
        h=0.40,
        size=fitted_size(
            "\n".join(map(str, commentary)),
            base=13 * font_scale,
            minimum=10 * font_scale,
            one_line_units=48,
            scale=1.0,
        ),
        color=palette["ink"],
        muted=palette["muted"],
        fonts=fonts,
        space_after=6,
    )
    return palette


def layout_quote(
    slide: Any,
    spec: dict[str, Any],
    deck: dict[str, Any],
    theme: dict[str, str],
    fonts: dict[str, str],
    font_scale: float,
) -> dict[str, str]:
    tone = str(spec.get("tone", "dark"))
    palette = page_palette(theme, tone)
    fill_background(slide, palette["bg"])
    add_text(
        slide,
        "“",
        x=0.05,
        y=0.02,
        w=0.20,
        h=0.27,
        size=122 * font_scale,
        color=mix_hex(palette["bg"], theme["accent"], 0.56),
        fonts=fonts,
        use_heading=True,
    )
    quote = balanced_text_break(str(spec.get("quote") or spec.get("title") or ""))
    add_text(
        slide,
        quote,
        x=0.12,
        y=0.25,
        w=0.75,
        h=0.32,
        size=fitted_size(
            quote,
            base=29 * font_scale,
            minimum=18 * font_scale,
            one_line_units=48,
            scale=1.0,
            line_factor=1.2,
        ),
        color=palette["ink"],
        fonts=fonts,
        use_heading=True,
        valign=MSO_ANCHOR.MIDDLE,
        name="Title",
    )
    attribution = str(spec.get("attribution", ""))
    if attribution:
        add_text(
            slide,
            attribution,
            x=0.12,
            y=0.62,
            w=0.55,
            h=0.06,
            size=11 * font_scale,
            color=palette["muted"],
            fonts=fonts,
            bold=True,
        )
    context = str(spec.get("context", ""))
    if context:
        add_text(
            slide,
            context,
            x=0.70,
            y=0.77,
            w=0.23,
            h=0.10,
            size=9.5 * font_scale,
            color=palette["muted"],
            fonts=fonts,
            align=PP_ALIGN.RIGHT,
        )
    return palette


def layout_closing(
    slide: Any,
    spec: dict[str, Any],
    deck: dict[str, Any],
    theme: dict[str, str],
    fonts: dict[str, str],
    font_scale: float,
) -> dict[str, str]:
    tone = str(spec.get("tone", "dark"))
    palette = page_palette(theme, tone)
    fill_background(slide, palette["bg"])
    title = str(spec.get("title", "Next step"))
    if font_scale < 0.95:
        title = balanced_text_break(title, threshold=16)
    add_text(
        slide,
        title,
        x=0.07,
        y=0.14,
        w=0.76,
        h=0.25,
        size=fitted_size(title, base=42 * font_scale, minimum=25 * font_scale, one_line_units=25, scale=1.0),
        color=palette["ink"],
        fonts=fonts,
        use_heading=True,
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
        name="Title",
    )
    subtitle = str(spec.get("subtitle", ""))
    if subtitle:
        add_text(
            slide,
            subtitle,
            x=0.07,
            y=0.42,
            w=0.58,
            h=0.10,
            size=15 * font_scale,
            color=palette["muted"],
            fonts=fonts,
        )
    actions = list(spec.get("actions", []))
    if actions:
        gap = 0.02
        y = 0.64
        width = (0.86 - gap * (len(actions) - 1)) / len(actions)
        for index, action in enumerate(actions):
            if isinstance(action, dict):
                label = str(action.get("label") or action.get("title") or "")
                body = str(action.get("body") or "")
            else:
                label = str(action)
                body = ""
            x = 0.07 + index * (width + gap)
            add_rect(slide, x=x, y=y, w=width, h=0.15, fill=palette["surface"], radius=True, name="Action")
            add_text(
                slide,
                label,
                x=x + 0.02,
                y=y + 0.025,
                w=width - 0.04,
                h=0.04,
                size=fitted_size(label, base=12.5 * font_scale, minimum=9.5 * font_scale, one_line_units=max(20, width * 55), scale=1.0),
                color=palette["ink"],
                fonts=fonts,
                bold=True,
            )
            if body:
                add_text(
                    slide,
                    body,
                    x=x + 0.02,
                    y=y + 0.075,
                    w=width - 0.04,
                    h=0.05,
                    size=9 * font_scale,
                    color=palette["muted"],
                    fonts=fonts,
                )
    return palette


LAYOUT_HANDLERS = {
    "cover": layout_cover,
    "section": layout_section,
    "agenda": layout_agenda,
    "statement": layout_statement,
    "cards": layout_cards,
    "split": layout_split,
    "comparison": layout_comparison,
    "timeline": layout_timeline,
    "metrics": layout_metrics,
    "chart": layout_chart,
    "quote": layout_quote,
    "closing": layout_closing,
}


def validate_deck(deck: dict[str, Any]) -> None:
    if not isinstance(deck, dict):
        raise ValueError("Deck spec must be a JSON object")
    slides = deck.get("slides")
    if not isinstance(slides, list) or not slides:
        raise ValueError("Deck spec must contain a non-empty slides array")
    for index, slide in enumerate(slides, 1):
        if not isinstance(slide, dict):
            raise ValueError(f"Slide {index} must be an object")
        slide_type = str(slide.get("type", ""))
        if slide_type not in SLIDE_TYPES:
            raise ValueError(f"Slide {index}: unsupported type {slide_type!r}")
        required = {
            "cover": ("title",),
            "section": ("title",),
            "agenda": ("title", "items"),
            "statement": ("title",),
            "cards": ("title", "cards"),
            "split": ("title",),
            "comparison": ("title", "left", "right"),
            "timeline": ("title", "steps"),
            "metrics": ("title", "metrics"),
            "chart": ("title", "chart"),
            "quote": ("quote",),
            "closing": ("title",),
        }[slide_type]
        for field in required:
            if field not in slide:
                raise ValueError(f"Slide {index} ({slide_type}) requires {field!r}")


def build_deck(spec_path: Path, output_path: Path, force: bool = False) -> dict[str, Any]:
    global SLIDE_WIDTH_EMU, SLIDE_HEIGHT_EMU
    if output_path.exists() and not force:
        raise FileExistsError(
            f"Output already exists: {output_path}. Pass --force to replace it."
        )
    base_dir = spec_path.resolve().parent
    deck = json.loads(spec_path.read_text(encoding="utf-8"))
    validate_deck(deck)
    theme_name = str(deck.get("theme", "business"))
    if theme_name not in THEMES:
        raise ValueError(f"theme must be one of {', '.join(THEMES)}")
    theme = dict(THEMES[theme_name])
    if isinstance(deck.get("palette"), dict):
        for key, value in deck["palette"].items():
            if key in theme and isinstance(value, str):
                theme[key] = value.lstrip("#").upper()
    language = str(deck.get("language", "zh-CN"))
    fonts = font_plan(language, deck.get("fonts") if isinstance(deck.get("fonts"), dict) else None)
    aspect = str(deck.get("aspect_ratio", "16:9"))
    prs = Presentation()
    if aspect == "4:3":
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)
        font_scale = 0.86
    elif aspect == "16:9":
        prs.slide_width = Inches(13.333333)
        prs.slide_height = Inches(7.5)
        font_scale = 1.0
    else:
        raise ValueError("aspect_ratio must be 16:9 or 4:3")
    SLIDE_WIDTH_EMU = int(prs.slide_width)
    SLIDE_HEIGHT_EMU = int(prs.slide_height)
    blank_layout = prs.slide_layouts[6]
    total = len(deck["slides"])
    warnings: list[str] = []

    for index, slide_spec in enumerate(deck["slides"], 1):
        slide = prs.slides.add_slide(blank_layout)
        slide_type = str(slide_spec["type"])
        handler = LAYOUT_HANDLERS[slide_type]
        kwargs = {
            "slide": slide,
            "spec": slide_spec,
            "deck": deck,
            "theme": theme,
            "fonts": fonts,
            "font_scale": font_scale,
        }
        if slide_type == "split":
            palette = handler(**kwargs, base_dir=base_dir)
        else:
            palette = handler(**kwargs)
        add_notes(slide, slide_spec.get("notes"))
        add_footer(
            slide,
            spec=slide_spec,
            palette=palette,
            fonts=fonts,
            page_number=index,
            total=total,
            font_scale=font_scale,
            show_number=slide_type not in {"cover", "quote"},
        )
        notes = str(slide_spec.get("notes", "")).strip()
        if slide_type not in {"cover", "section", "closing"} and not notes:
            warnings.append(f"Slide {index}: no speaker notes")

    prs.core_properties.title = str(deck.get("title", ""))
    prs.core_properties.subject = str(deck.get("subtitle", ""))
    prs.core_properties.author = str(deck.get("author", ""))
    prs.core_properties.comments = "Built with ppt-deck-builder"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    return {
        "output": str(output_path.resolve()),
        "slides": total,
        "theme": theme_name,
        "aspect_ratio": aspect,
        "fonts": fonts,
        "warnings": warnings,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("spec", type=Path, help="UTF-8 JSON deck specification")
    parser.add_argument("output", type=Path, help="Destination .pptx path")
    parser.add_argument("--force", action="store_true", help="Replace an existing output file")
    parser.add_argument("--json", action="store_true", help="Print a JSON result")
    args = parser.parse_args()
    try:
        result = build_deck(args.spec, args.output, args.force)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"Created {result['output']}")
        print(
            f"{result['slides']} slides | {result['aspect_ratio']} | "
            f"theme={result['theme']} | font={result['fonts']['body_ea']}"
        )
        for warning in result["warnings"]:
            print(f"WARN: {warning}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
