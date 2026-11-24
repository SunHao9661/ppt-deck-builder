# PPTX Production

Use this reference for file mechanics, template edits, and the bundled JSON
builder.

## Required Output Properties

For an editable presentation:

- real PowerPoint text boxes, shapes, charts, and tables
- speaker notes in PowerPoint notes, not visible slide text
- consistent slide dimensions and orientation
- no leftover template or demo copy
- no object past the slide boundary
- image assets retained with the deck
- source file or generator retained for future revision

## Route-Specific Output

| Route | Required extra evidence |
|---|---|
| Native | deck spec, notes, render/contact sheet |
| Template | source template, template inventory, render comparison |
| Image | prompt/source record, selected image per page, image manifest |
| Hybrid | native/visual layer manifest, approved sample, comparison |
| Fidelity | approved reference per page, reconstruction notes, render comparison |

The route contract is in [production-routes.md](production-routes.md). A clean
PPTX file is not enough if the route requires visual comparison or source
provenance.

## New Deck with the Bundled Builder

Run:

```bash
python3 scripts/build_pptx.py deck.json output/deck.pptx
```

The input is UTF-8 JSON. Paths are resolved relative to the JSON file.

Top-level fields:

```json
{
  "title": "Deck title",
  "subtitle": "Optional subtitle",
  "author": "Author",
  "date": "2026-11-24",
  "language": "zh-CN",
  "theme": "editorial | business | tech | academic",
  "aspect_ratio": "16:9 | 4:3",
  "slides": []
}
```

Every slide may include:

- `type`
- `id`
- `title`
- `kicker`
- `source`
- `notes`
- `tone`: `light` or `dark` for applicable page types

Supported slide types:

| Type | Key fields |
|---|---|
| `cover` | title, subtitle, eyebrow, meta |
| `section` | number, title, subtitle |
| `agenda` | title, items |
| `statement` | title, body, kicker |
| `cards` | title, cards: label/title/body |
| `split` | title, left/right headings and items, optional image |
| `comparison` | title, left/right headings and items |
| `timeline` | title, steps: label/title/body |
| `metrics` | title, metrics: value/label/note |
| `chart` | title, commentary, chart |
| `quote` | quote, attribution, context |
| `closing` | title, subtitle, actions |

Chart shape:

```json
{
  "chart_type": "column | bar | line | pie",
  "categories": ["Q1", "Q2", "Q3"],
  "series": [{"name": "Revenue", "values": [12, 18, 24]}],
  "unit": "USDm"
}
```

Image path fields accept an absolute path or a path relative to the JSON file.
Use `image_fit: "contain"` for screenshots and diagrams; use `cover` for
photographs.

## Existing Deck or Corporate Template

If the host provides a dedicated presentation tool, use it first.

Otherwise:

1. Work on a copy; never overwrite the source deck.
2. Render thumbnail grids and inspect representative slides.
3. Inspect object structure and inventory slide roles/placeholders:

```bash
python3 scripts/inspect_pptx.py source.pptx --json > template-inventory.json
```

4. Map content to the smallest set of layouts that preserve the template.
5. Replace content without collapsing run formatting.
6. Remove unused template groups, not just the text inside them.
7. Keep images, charts, notes, and source timing in their native locations.
8. Validate that the final file opens.

For complex OOXML edits, manipulate object and run properties directly rather
than replacing an entire text frame with plain text.

When a template has placeholder pages with no caption role, skip them. Do not
invent a cover, agenda, section divider, or “Thank you” page from an unrelated
layout just to satisfy a conventional structure.

### Template Text Replacement

- Match the exact shape/paragraph/run before editing.
- Keep run font size, weight, color, spacing, and emphasis.
- Use paragraph-level replacement only when the whole paragraph is the
  template slot.
- Replace all editable placeholders; remove the whole orphan group when a
  visual item no longer has content.
- Keep agenda labels synchronized with section dividers and breadcrumb text.
- For charts, update native chart data rather than replacing the chart with an
  image.
- Never modify the supplied source in place.

## Image-First Assembly

Use:

```bash
python3 scripts/assemble_image_pptx.py deck-plan.json image-deck.pptx --images-dir slides
```

The plan supplies page order, titles/alt text, notes, and aspect ratio. The
script fits images without stretching and records warnings when source ratios
differ from the slide.

Image-first output is not considered text-editable. Keep the prompts and source
images so a page can be regenerated or later reconstructed.

For video or interactions, see [html-motion.md](html-motion.md).

## Hybrid Layer Rules

Use two layers:

- **native information layer:** title, body, key numbers, labels, legends,
  citations, page number, simple chart/table, simple dividers
- **visual asset layer:** photos, official logos, product UI, complex quality
  effects, complex custom geometry, and non-text illustration that would be
  degraded by a rough native approximation

Do not put main text in the visual layer to hide misalignment. Do not replace an
important curve, flow, or custom diagram with a few preset shapes merely because
the shapes are editable.

If the source image contains text that must become editable:

1. reconstruct the background/visual without the original text
2. place native text on top
3. render and inspect for ghost text or double text

Keep tight non-text crops in the project and record why each one is needed.

## Fonts

PowerPoint renders fonts on the viewer's machine. The authoring environment may
substitute fonts during PDF/PNG preview, which changes text width and overflow.

Use widely available fonts when portability matters. For Chinese decks,
`Microsoft YaHei` is a safe Windows-oriented choice; on macOS, check the rendered
result because preview substitution can differ. When a client supplies a brand
font, verify the audience can install or embed it, or provide a safe fallback.

Do not silently reduce important body text to 8–10 pt to fit. Split the slide or
shorten the copy.

## Charts and Tables

Use native chart objects for editable business data. A raster chart is acceptable
only when no native chart type can express the relationship or the chart is a
captured result that must remain pixel-identical. State the limitation.

For tables:

- keep headers clear
- avoid more than several rows/columns on a projected slide
- move detailed tables into an appendix or separate handout
- keep units and source visible
- align numbers by decimal or right edge

## Images

- Preserve the original image file in the project.
- Use contain crops for screenshots, code, diagrams, and evidence with text.
- Use cover crops for photos when the subject remains intact.
- Do not upscale small screenshots until they become unreadable.
- Check rights for supplied and web images; record the source where relevant.

## Speaker Notes

Place the delivery script in the notes pane, not on the slide. Notes should
explain and transition, not mirror visible text.

## Troubleshooting

- **File will not open:** regenerate from the last valid source, then render.
- **Text moves in preview:** check font availability/substitution, set explicit
  margins, and leave width slack for non-safe fonts.
- **Chart missing after edit:** keep native chart structure intact; avoid
  deleting or replacing chart relationships manually.
- **Image missing:** use relative paths inside the project and keep the assets
  folder with the deck.
- **Template edit breaks:** patch run-level text and preserve paragraph
  properties; do not flatten an entire text frame.
