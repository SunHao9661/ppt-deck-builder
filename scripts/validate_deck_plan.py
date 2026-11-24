#!/usr/bin/env python3
"""Validate a route-aware PPT deck plan and print its ghost deck."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any


ROUTES = {"native", "template", "image", "hybrid", "html", "motion", "fidelity"}
DENSITIES = {"presentation", "balanced", "read-close", "high-density"}
GENERIC_TITLES = {
    "overview",
    "background",
    "current situation",
    "analysis",
    "results",
    "discussion",
    "recommendations",
    "next steps",
    "thank you",
    "thanks",
    "q&a",
    "概览",
    "背景",
    "现状",
    "分析",
    "结果",
    "讨论",
    "建议",
    "下一步",
    "谢谢",
    "感谢聆听",
}
EXEMPT_GENERIC_ROLES = {"agenda", "section", "appendix-index", "divider"}
EVIDENCE_ROLES = {
    "evidence",
    "metrics",
    "chart",
    "comparison",
    "analysis",
    "finding",
    "data",
}
ACTION_ROLES = {"recommendation", "decision", "conclusion", "closing"}


def as_list(value: Any) -> list[Any]:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    return [value]


def validate(plan: dict[str, Any]) -> dict[str, Any]:
    issues: list[dict[str, Any]] = []

    def add(severity: str, code: str, message: str, slide: str | None = None) -> None:
        row: dict[str, Any] = {"severity": severity, "code": code, "message": message}
        if slide is not None:
            row["slide"] = slide
        issues.append(row)

    if not isinstance(plan, dict):
        raise ValueError("Deck plan must be a JSON object")

    for field in ("title", "audience", "outcome", "core_message", "route", "visual", "slides"):
        if field not in plan:
            add("ERROR", "MISSING_FIELD", f"Missing required top-level field: {field}")

    route = str(plan.get("route", "")).strip()
    if route and route not in ROUTES:
        add(
            "ERROR",
            "INVALID_ROUTE",
            f"route must be one of: {', '.join(sorted(ROUTES))}",
        )

    visual = plan.get("visual")
    if not isinstance(visual, dict):
        visual = {}
        add("ERROR", "VISUAL_OBJECT", "visual must be an object")
    for field in ("style", "density", "aspect_ratio"):
        if not str(visual.get(field, "")).strip():
            add("ERROR", "VISUAL_MISSING", f"visual.{field} is required")
    density = str(visual.get("density", ""))
    if density and density not in DENSITIES:
        add("WARN", "UNKNOWN_DENSITY", f"Unknown density: {density}")

    if not str(plan.get("language", "")).strip():
        add("WARN", "MISSING_LANGUAGE", "language is not set")
    duration = plan.get("duration_minutes")
    if duration is None:
        add("WARN", "MISSING_DURATION", "duration_minutes is not set")
    elif not isinstance(duration, (int, float)) or duration <= 0:
        add("ERROR", "INVALID_DURATION", "duration_minutes must be a positive number")

    slides = plan.get("slides")
    if not isinstance(slides, list) or not slides:
        add("ERROR", "NO_SLIDES", "slides must be a non-empty array")
        slides = []

    ids: set[str] = set()
    titles: list[str] = []
    roles: list[str] = []
    visuals: list[str] = []
    content_count = 0
    previous_title = ""
    previous_visual = ""
    visual_run = 0

    for index, slide in enumerate(slides, 1):
        if not isinstance(slide, dict):
            add("ERROR", "SLIDE_OBJECT", f"Slide {index} must be an object")
            continue
        slide_id = str(slide.get("id", "")).strip()
        role = str(slide.get("role", "")).strip()
        title = " ".join(str(slide.get("title", "")).split())
        purpose = str(slide.get("purpose", "")).strip()
        visual_form = " ".join(str(slide.get("visual", "")).split())
        if not slide_id:
            add("ERROR", "MISSING_SLIDE_ID", "Slide ID is required", str(index))
        elif slide_id in ids:
            add("ERROR", "DUPLICATE_SLIDE_ID", f"Duplicate slide ID: {slide_id}", slide_id)
        ids.add(slide_id)
        if not role:
            add("ERROR", "MISSING_ROLE", "Slide role is required", slide_id or str(index))
        if not title:
            add("ERROR", "MISSING_TITLE", "Slide title is required", slide_id or str(index))
        if not purpose and role not in {"cover", "section", "divider", "agenda"}:
            add("WARN", "MISSING_PURPOSE", "Content slide has no purpose", slide_id or str(index))

        normalized_title = re.sub(r"[\s:：|·—–-]+$", "", title).strip()
        if (
            normalized_title.lower() in GENERIC_TITLES
            and role.lower() not in EXEMPT_GENERIC_ROLES
        ):
            add(
                "WARN",
                "GENERIC_TITLE",
                f"Title {title!r} is a topic label; state the takeaway",
                slide_id or str(index),
            )
        if title and title == previous_title:
            add(
                "WARN",
                "DUPLICATE_TITLE",
                f"Consecutive duplicate title: {title!r}",
                slide_id or str(index),
            )
        previous_title = title
        titles.append(title)
        roles.append(role)

        if role not in {"cover", "section", "divider", "agenda", "appendix", "appendix-index"}:
            content_count += 1

        points = as_list(slide.get("points"))
        evidence = as_list(slide.get("evidence"))
        if role in EVIDENCE_ROLES:
            if not evidence and not points:
                add(
                    "WARN",
                    "NO_EVIDENCE",
                    "Evidence/data page has no evidence or supporting points",
                    slide_id or str(index),
                )
            if evidence and not str(slide.get("source", "")).strip():
                add(
                    "WARN",
                    "MISSING_SOURCE",
                    "Evidence exists but source/attribution is empty",
                    slide_id or str(index),
                )
            if not str(slide.get("so_what", "")).strip():
                add(
                    "WARN",
                    "MISSING_SO_WHAT",
                    "Evidence/data page has no so-what or implication",
                    slide_id or str(index),
                )
        if role in ACTION_ROLES and not points and not str(slide.get("so_what", "")).strip():
            add(
                "WARN",
                "WEAK_ACTION",
                "Decision/conclusion page has no explicit action or takeaway",
                slide_id or str(index),
            )

        if visual_form:
            if visual_form == previous_visual:
                visual_run += 1
            else:
                visual_run = 1
            if visual_run >= 3:
                add(
                    "WARN",
                    "VISUAL_REPEAT",
                    "Three consecutive slides use the same visual form",
                    slide_id or str(index),
                )
            previous_visual = visual_form
        visuals.append(visual_form)

        if route in {"image", "hybrid", "fidelity"}:
            if not visual_form:
                add(
                    "WARN",
                    "MISSING_VISUAL",
                    f"{route} route page has no visual direction",
                    slide_id or str(index),
                )
        if route == "image":
            if not str(slide.get("image_prompt", "")).strip() and not str(slide.get("image", "")).strip():
                add(
                    "ERROR",
                    "IMAGE_SOURCE_REQUIRED",
                    "Image route page requires image_prompt or image",
                    slide_id or str(index),
                )
        if route == "motion":
            if index > 1 and not str(slide.get("motion", "")).strip():
                add(
                    "WARN",
                    "MISSING_MOTION",
                    "Motion route page has no transition/motion note",
                    slide_id or str(index),
                )

    if slides:
        first_role = str(slides[0].get("role", "")) if isinstance(slides[0], dict) else ""
        last_role = str(slides[-1].get("role", "")) if isinstance(slides[-1], dict) else ""
        artifact = str(plan.get("artifact", "deck")).lower()
        if artifact not in {"fragment", "section", "handout", "appendix"}:
            if first_role not in {"cover", "title"}:
                add("WARN", "NO_COVER", "Deck does not start with a cover/title page")
            if last_role not in {"conclusion", "closing", "decision", "recommendation", "next-steps"}:
                add(
                    "WARN",
                    "WEAK_ENDING",
                    "Deck does not end on a conclusion, decision, or action",
                )

    if route == "template" and not str(plan.get("template", "")).strip():
        add("ERROR", "TEMPLATE_REQUIRED", "Template route requires a template path")
    if route in {"hybrid", "fidelity"} and not str(plan.get("editability", "")).strip():
        add(
            "ERROR",
            "EDITABILITY_REQUIRED",
            f"{route} route requires an editability statement",
        )
    if route == "motion" and not plan.get("motion"):
        add("ERROR", "MOTION_POLICY_REQUIRED", "Motion route requires a top-level motion policy")
    if route == "html" and str(visual.get("aspect_ratio", "")) not in {"16:9", "4:3"}:
        add("WARN", "HTML_STAGE", "HTML deck should use a stable 16:9 or 4:3 stage")

    if isinstance(duration, (int, float)) and duration > 0:
        maximum_practical = duration * 2.0
        if content_count > maximum_practical:
            add(
                "WARN",
                "TIME_RISK",
                f"{content_count} content pages for {duration:g} minutes is likely too dense",
            )

    return {
        "title": str(plan.get("title", "")),
        "route": route,
        "slide_count": len(slides),
        "content_slide_count": content_count,
        "ghost_deck": [
            {"slide": index, "title": title}
            for index, title in enumerate(titles, 1)
        ],
        "roles": roles,
        "issues": issues,
    }


def print_human(result: dict[str, Any], ghost: bool) -> None:
    print(
        f"{result['title'] or '[untitled]'} | route={result['route'] or '[missing]'} | "
        f"{result['slide_count']} slides"
    )
    if ghost:
        print("\nGhost deck:")
        for row in result["ghost_deck"]:
            print(f"{row['slide']:02d}  {row['title'] or '[no title]'}")
    if not result["issues"]:
        print("\nNo automated issues found.")
        return
    print("\nIssues:")
    for issue in result["issues"]:
        slide = f" {issue['slide']}" if "slide" in issue else ""
        print(f"{issue['severity']:5}  {issue['code']}{slide}: {issue['message']}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("plan", type=Path)
    parser.add_argument("--ghost", action="store_true")
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--strict", action="store_true")
    args = parser.parse_args()
    try:
        plan = json.loads(args.plan.read_text(encoding="utf-8"))
        result = validate(plan)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
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
