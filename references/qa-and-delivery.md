# QA and Delivery

Technical validity is the baseline. The deck is finished only when the rendered
slides communicate clearly.

## Required QA Order

1. **Story:** read the title-only ghost deck.
2. **Content:** verify facts, numbers, order, citations, assumptions.
3. **Structure:** ensure the expected slides exist and no demo/placeholder copy
   remains.
4. **File:** open the deck and audit shape bounds, notes, and slide dimensions.
5. **Render:** export every slide to PNG.
6. **Reference compare:** if a template, image-to-editable reconstruction, or
   approved blueprint exists, compare source and output.
7. **Visual:** inspect every slide full-size; inspect the contact sheet for rhythm.
8. **Repair:** fix the owning layer and re-render changed slides.

Use the bundled tools:

```bash
python3 scripts/audit_pptx.py deck.pptx --ghost
python3 scripts/render_pptx.py deck.pptx work/renders --dpi 144
python3 scripts/compare_renders.py approved/ work/renders --out-dir work/compare
```

The comparison command creates:

- side-by-side reference/output images
- 50/50 overlays
- difference heatmaps
- a numeric report with changed-pixel ratios

## Route Gates

### Native

- text and key charts are native
- no repeated page grammar without a narrative reason
- page roles include evidence and conclusion, not only agenda/cards

### Template

- result matches rendered template roles, not just the theme colors
- no template promotion or placeholder pages
- same-role titles retain original position and scale
- unused template graphics are removed with their content
- chart data remains native

### Image

- image text is correctly spelled and readable
- page density and style are coherent
- no unsupported factual claim or invented chart value
- output is described as image-based, not editable
- notes and alt text exist where appropriate

### Hybrid / Fidelity

Two gates must pass together:

1. **Information editability:** key text, figures, labels, sources, and simple
   charts can be edited as native objects.
2. **Visual semantics:** layout, surface system, hierarchy, density, chart
   geometry, and spatial relationships still match the approved reference.

Neither `all text is editable` nor `the page looks similar` passes by itself.

### HTML

- every slide stays inside a fixed 16:9 stage
- keyboard, overview, deep links, and presenter mode work
- speaker notes are hidden from the audience
- print/PDF produces one page per slide when requested

### Motion

- static pages were approved first
- transitions start and end on the correct pages
- text remains stable
- codec/resolution/fps/audio are consistent
- final playback is inspected, not only individual frames

## Visual Hard Failures

Fix these before delivery:

- text clipped by a shape or slide edge
- text overlapping another text block or meaningful visual
- image subject, label, axis, code, or screenshot content cropped away
- chart without readable labels or with missing units/source
- unreadable contrast
- lost title alignment between equivalent page roles
- broken or missing image
- visibly orphaned template graphics
- accidental blank page
- stray placeholder, demo sentence, or internal note on the slide
- whole-slide reference/blueprint used as a content-page background without an
  explicit image-only route
- simple chart or conclusion block rasterized while the route promises editable
  information
- visual semantics downgraded from an approved reference into a generic card
  stack or default chart

## High-Fidelity Repairs

Repair at the owning layer:

1. missing or wrong data/content
2. canvas, container, font size, and overflow
3. editable/native layer
4. cropped visual assets or custom geometry
5. interpretation and bottom-line message
6. alignment, spacing, and cross-page rhythm

After every page repair, render that page again. After merging single-page
outputs, render the merged deck again to catch relationship, theme, or ordering
regressions.

## Soft Issues

Act only when clearly visible:

- content cluster too close to the edge or footer
- inconsistent column gaps
- title/body spacing that changes without semantic reason
- accidental one-word line or word split
- decorative element competing with the main evidence
- content container too large for its content, creating accidental emptiness

Do not fix soft issues by oscillating between two versions. Choose the clearer
version and stop.

## Repair Discipline

- Fix content in the source plan or generator, not by typing over generated
  slides unless the user supplied a final deck.
- Preserve the template and visual system while fixing layout.
- Do not shrink everything to make overflow disappear.
- Do not remove citations or caveats to simplify a slide.
- Re-render after a fix; object coordinates alone are not enough.
- Use at most two repair passes for ordinary production. Report unresolved
  issues instead of hiding them.

## Final Handoff

Report:

- final `.pptx` path
- contact-sheet and render-directory path
- slide count and selected route/theme
- source JSON, script, outline, or template path
- speaker-notes status
- comparison report path when a reference existed
- deliverable format limitation (`image-based`, `HTML`, `video`, etc.)
- any remaining limitation or unverified area

When delivering to a non-technical user, lead with the file and one sentence
describing the deck. Keep process detail secondary.
