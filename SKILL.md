---
name: ppt-deck-builder
description: Create, revise, or restyle presentation decks across native PPTX, template-faithful, image-first, hybrid editable, and HTML routes. Use for PPT, PPTX, slides, decks, 汇报, 路演, 课件, 答辩, 培训, 提案, or requests to improve an existing presentation; exclude static posters, one-off social images, and Word documents unless the request is specifically to turn them into a deck.
metadata:
  short-description: Story-first editable PPT decks with rendered QA
---

# PPT Deck Builder

Build the deck as a communication artifact, not as a pile of decorated slides.
The default deliverable is a native, editable `.pptx`; use HTML only when the
user explicitly wants a browser-first or motion-led presentation.

## Route First

Choose the simplest production route that satisfies the actual handoff:

| Need | Route |
|---|---|
| New editable deck, charts, notes, standard business/academic design | Native PPTX |
| Supplied corporate/user template must stay intact | Template-faithful PPTX |
| Visual richness matters most; each slide may be one final image | Image-first PPTX |
| Images must become a presentation with editable key text | Hybrid reconstruction |
| Browser-first deck, live URL, presenter mode, rich motion | HTML deck |
| Video transitions or a self-running visual sequence are required | Motion/video deck |
| “1:1,” “match the blueprint,” or complex native graphics are demanded | High-fidelity blueprint workflow |

Read [references/production-routes.md](references/production-routes.md) before
choosing a route. If the request already matches one row, do not present a
route menu to the user.

Rules:

- Do not call image-only slides “editable.”
- Do not generate a full image deck when colleagues must edit paragraphs,
  charts, or shapes.
- Do not rebuild a supplied template from scratch when the template itself is
  the visual authority.
- Do not add video transitions just because an image deck exists; motion is a
  separate delivery choice.
- For image-first, hybrid, or 1:1 reconstruction, approve a representative
  content sample before full production.

If the host already provides a dedicated presentation or PPTX tool, use it for
file mechanics. This skill governs the deck's communication, design, and QA.

## 1. Define the Communication Contract

Extract these decisions from the request and source material:

- audience and what they already know
- desired outcome: understand, believe, choose, approve, act, or learn
- presentation setting and duration
- reading mode: live presentation, balanced, or read-close handout
- language, brand/template constraints, and required facts
- final file format and whether someone else must edit it

Ask only when a missing answer would materially change the story or file format.
Otherwise make a reasonable assumption and state it in the handoff.

Read [references/story-system.md](references/story-system.md) when planning the
outline, choosing a narrative mode, writing slide titles, or producing speaker
notes.

## 2. Lock the Story Before the Visuals

Create a slide-by-slide plan with a stable ID, page job, action title, evidence,
visual form, source, and optional speaker note. Then run the ghost-deck test:
read only the titles in sequence. They must form a complete, causal argument.

For route-based production, create the structured plan in
[references/deck-plan.md](references/deck-plan.md) and validate it:

```bash
python3 scripts/validate_deck_plan.py deck-plan.json --ghost
```

The plan records the route, audience outcome, visual direction, editability,
slide evidence, and source assets. Do not freeze a visual blueprint or start
full-deck generation before it passes.

Rules that matter most:

- one primary claim per slide
- titles state the takeaway, not only the topic
- evidence follows the claim; explanation is not merely repeated on the slide
- vary page roles; do not repeat the same card grid
- finish with a conclusion, decision, or next action, not an empty thank-you page
- every external fact, figure, and borrowed visual keeps its source

Use a loose slide budget: roughly one content slide per minute, fewer for live
presentation and more for read-close material. Treat it as a planning cue, not
a hard rule.

## 3. Choose One Visual System

Choose one visual direction based on audience, subject, and venue. When the user
has no existing brand/template:

- editorial: human, cultural, narrative, research, or thought-leadership themes
- business: reports, proposals, reviews, and decision meetings
- tech: product, engineering, AI, data, or launch narratives
- academic: papers, seminars, defenses, and grant briefings

Use a single type hierarchy, one dominant palette, and one recurring visual
device. Style consistency means the same visual grammar across pages, not the
same composition on every page.

Read [references/visual-system.md](references/visual-system.md) before choosing
layouts, charts, imagery, typography, density, or a template.

## 4. Build a Real Deck

Read [references/production-routes.md](references/production-routes.md) and
[references/pptx-production.md](references/pptx-production.md) before building.

### Native PPTX

For a new native deck, prefer the bundled builder:

```bash
python3 scripts/build_pptx.py deck.json deck.pptx
```

Use [references/example-deck.json](references/example-deck.json) as the starting
shape for the JSON input. The builder supports cover, section, agenda, statement,
cards, split, comparison, timeline, metrics, chart, quote, and closing pages.
Text, shapes, charts, speaker notes, and page numbers remain editable.

### Template-Faithful PPTX

Never overwrite the supplied file. Inspect its rendered pages and object
structure before editing:

```bash
python3 scripts/inspect_pptx.py source.pptx --json > template-inventory.json
```

Preserve the template's layout, type scale, colors, master/layout structure,
and chart relationships. Replace text at paragraph/run level. Remove unused
template groups rather than leaving orphaned art.

### Image-First PPTX

Use raster generation only when the user accepts image-based slides. Keep the
prompt, generated source, selected image, and speaker notes for each page.
Assemble final selected images with:

```bash
python3 scripts/assemble_image_pptx.py deck-plan.json image-deck.pptx --images-dir slides
```

### Hybrid or High-Fidelity Reconstruction

Use two layers:

- native editable text, key numbers, labels, and simple charts
- tightly cropped non-text images or vector geometry only where native
  reconstruction would materially destroy the approved visual

Freeze the approved plan and representative sample before producing the
remaining pages. Compare rendered output against the approved reference:

```bash
python3 scripts/compare_renders.py approved/ work/renders --out-dir work/compare
```

### HTML or Motion

For browser-first, presenter-mode, or motion-led delivery, read
[references/html-motion.md](references/html-motion.md). Preserve a fixed slide
stage and keep speaker notes separate from audience-visible content.

Keep the source artifacts:

- `deck-plan.json`
- final `.pptx`
- generating script or JSON
- outline or slide plan
- images and other linked assets
- speaker notes
- approved samples and comparison reports when fidelity matters

## 5. Verify Rendered Pixels

Do not judge a deck from code or object coordinates alone. Render it, inspect
every slide at readable size, fix the owning layer, and re-render changed slides.

```bash
python3 scripts/audit_pptx.py deck.pptx --ghost
python3 scripts/render_pptx.py deck.pptx work/renders --dpi 144
```

The renderer creates one PNG per slide plus a labeled contact sheet. Inspect the
contact sheet for rhythm and every full-size slide for overflow, overlap,
alignment, cropping, and readability.

For template fidelity, image-to-editable reconstruction, or an approved visual
blueprint, compare reference and output:

```bash
python3 scripts/compare_renders.py approved/ work/renders --out-dir work/compare
```

Review the side-by-side, 50/50 overlay, and difference heatmap. A clean
structural audit does not prove visual fidelity.

Use at most two repair passes unless the user explicitly asks for exhaustive
visual polish. If a defect remains after the repair budget, report the slide and
the concrete limitation instead of claiming completion.

Read [references/qa-and-delivery.md](references/qa-and-delivery.md) for the
required checks and handoff format.

## Non-Negotiable Presentation Checks

- No clipped, overlapping, or off-canvas text.
- Body copy remains readable at the target viewing distance.
- Important meaning does not depend on color alone.
- Images keep their subject and evidence legible; screenshots use contain, not
  destructive cover crops.
- Charts have a clear takeaway, labeled series, units, and source.
- Similar page roles share title, margin, and typography positions.
- Speaker notes add explanation and transitions; they do not read the slide.
- Slides with borrowed material include citations.
- Image-first output is explicitly labelled image-based; it is not presented as
  an editable-text deck.
- Template-faithful output keeps the supplied layout and removes all unrelated
  template promotion/placeholders.
- Hybrid/fidelity pages keep key information native while preserving approved
  non-text visual semantics.
- The final file opens successfully, contains the expected number of slides,
  and has no placeholder or demo copy.

## Delivery

Report:

- final file path and preview/contact-sheet path
- slide count and route used
- source-plan path for future edits
- assumptions made without user input
- any slide or capability that did not pass verification

Keep the final response concise. Lead with the finished file, not the process.
