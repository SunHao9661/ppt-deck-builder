# Visual System

Use this reference to create a coherent deck rather than a collection of
individually decorated pages.

## Route and Fidelity

| Delivery requirement | Best route |
|---|---|
| Editable text/charts/shapes in PowerPoint | Native PPTX |
| Corporate template or strict brand system | Native PPTX built from the template |
| Rich web animation, custom interaction, live URL | HTML |
| Visually exact slide images | Image-led deck; state that objects are not separately editable |
| Approved visual reference plus editable key text | Hybrid/fidelity reconstruction |
| Recorded or self-running presentation | Motion/video after static approval |

Never promise editable objects when the output is a full-slide image.
Use [production-routes.md](production-routes.md) for route selection.

## Visual Thesis

Choose five anchors and keep them stable:

1. palette
2. typography hierarchy
3. page margins and grid
4. one recurring visual device
5. density policy

The deck can vary composition by page role while these anchors remain stable.

## Density Modes

Choose one:

| Mode | Use | Behavior |
|---|---|---|
| Presentation | live keynote, pitch, classroom | one claim, sparse copy, large evidence/visual |
| Balanced | normal business/academic deck | claim + concise support |
| Read-close | pre-read, handout, due diligence | self-contained explanations, details, sources |
| High-density | consulting/report/analysis | multiple evidence zones, caveats, so-what, footnotes |

High density comes from structure and evidence, not from shrinking text. A
high-density page normally has:

- one action title
- one primary exhibit
- 2–4 supporting evidence blocks
- a visible so-what/decision line
- compact source and caveat treatment

If the page cannot support that without unreadable type, split it.

### Editorial

- warm paper or white field
- serif/display headline paired with a clean sans body
- assertive scale contrast and generous margins
- one restrained accent color
- images treated as evidence or editorial crops

Use for narratives, culture, research stories, thought leadership, and talks
that need a human point of view.

### Business

- white or near-white field
- navy, graphite, or content-relevant accent
- clear sans-serif hierarchy
- structured evidence pages, metrics, and decision framing
- restrained shape system

Use for reviews, proposals, operating plans, and executive communication.

### Tech

- dark field or high-contrast light field
- cool neutral text with one signal accent
- strong numeric scale, diagrams, and code/data captures
- avoid neon decoration unless the subject supports it

Use for product, engineering, AI, infrastructure, and launch narratives.

### Academic

- white field and one institutional or subject-derived accent
- conservative sans or serif pairing
- action titles, generous evidence area, clear citations
- no decorative icon clusters or stock imagery

Use for conferences, seminars, defenses, and grant presentations.

## Grid and Spacing

- Use a consistent outer margin; 0.5 in is a safe baseline for 16:9.
- Align repeated titles, body starts, chart frames, and citations.
- Use a small spacing set, such as 0.16, 0.32, 0.5, and 0.75 in.
- Prefer grouping through alignment and negative space before adding containers.
- Cards are useful when items are genuinely parallel. Do not put every idea in
  a rounded rectangle.
- Keep primary content away from slide edges and bottom chrome.

## Typography

Use two roles at most:

- heading/display
- body/annotation

Additional roles such as mono-data, quote, or citation are acceptable when they
perform a real semantic job.

Baseline guidance for 16:9:

- cover title: 36–54 pt
- content title: 24–32 pt
- lead statement: 28–42 pt
- body: 16–20 pt
- labels/captions: 11–14 pt

These are starting points, not a universal floor. Readability depends on room
size, viewing distance, projection, and language. When text no longer fits at a
readable size, cut, split, or restructure before shrinking.

Use one font family for Chinese and Latin text when possible. If not, pair them
deliberately and verify both render. Do not use decorative fonts for body copy.

## Color

- One color should dominate the visual field.
- Use one accent for emphasis; reserve semantic colors for data/status.
- Body text needs strong contrast against its actual background.
- Avoid relying on color alone to encode a decision or category.
- Do not choose a generic blue palette by default; tie the palette to the topic,
  audience, or existing brand.
- Dark and light sections can create rhythm, but they must share the same type,
  spacing, and graphic language.

## Page Grammar

Choose composition from the page job:

| Page job | Useful composition |
|---|---|
| Opening | one hook, title, attribution, restrained context |
| Section change | large section number/word, one sentence of orientation |
| Agenda | sequence with no more than 5 meaningful sections |
| Main claim | oversized statement plus one supporting line |
| Parallel points | 3–5 equal cards or rows |
| Comparison | side-by-side criteria with a clear conclusion |
| Process | timeline or step sequence, not an icon parade |
| Metrics | 2–4 large numbers with labels and scope |
| Evidence | one exhibit plus one interpretation |
| Quote | one human statement, author, and relevance |
| Closing | decision, next step, or takeaway |

Do not repeat the same “title + 3 cards” structure on consecutive pages.
Do not use a card grid when a table, timeline, matrix, or custom diagram better
preserves the relationship.

## Images

Decide what the image proves:

- mood or context
- real subject/product/place
- evidence screenshot
- process or architecture
- comparison

Use cover crops for photographs and contain/fit for screenshots, diagrams,
charts, code, or anything with essential text. Preserve aspect ratio unless the
user explicitly allows a crop. Give every image a purpose; decorative stock
should not displace useful evidence.

Generated imagery must follow the same palette, lighting, density, and subject
constraints as the rest of the deck. Do not generate slide titles, page numbers,
logos, watermarks, or other chrome inside the image.

## Charts

- Use a chart when the visual relationship matters: trend, comparison,
  distribution, composition, or correlation.
- Use a table when exact values and lookup matter more than shape.
- Give each chart one takeaway.
- Label units, time range, source, and highlighted points.
- Keep axis and label typography consistent across the deck.
- Remove unnecessary gridlines and chart clutter.
- Keep charts native when users may need to edit data.

## Template Use

When using a supplied template:

1. Inspect real rendered slides, not only XML.
2. Identify the available cover, section, content, image, chart, and closing
   page roles.
3. Preserve the template's typography scale and geometry.
4. Delete unused template groups and placeholder objects rather than leaving
   orphaned visuals.
5. Keep a copy of the original template unchanged.

When no template exists, a simple original system usually looks better than a
mix of several borrowed styles.

## Image-First Use

Image-first pages can be visually richer than native layouts, but the tradeoff
is explicit:

- text is baked into the image
- charts cannot be edited
- corrections mean regenerating or reconstructing the page
- accessibility/searchability requires alt text or a separate transcript

Use a representative content page as the sample, not only the cover. Check:

- exact wording and numbers
- hierarchy and readability
- repeated visual system
- chart/data honesty
- safe margins for the final aspect ratio
- no watermark, fake logo, or internal prompt artifacts

## HTML and Motion

For HTML:

- keep a fixed slide stage
- use one theme contract
- separate audience and presenter content
- keep motion purposeful and reduced-motion aware

For motion:

- fades/cuts are preferable when they communicate the sequence
- generative transitions should preserve titles and data readability
- transition clips must normalize to the same resolution/fps/audio
- inspect playback, not only still frames

See [html-motion.md](html-motion.md).

## Common Failure Modes

- generic title + bullet wall
- large empty header band or decorative stripe taking space from evidence
- different title positions on same-role slides
- all content centered
- too many accent colors
- every page using the same card grid
- icon clusters that add no meaning
- charts without units, source, or takeaway
- screenshots cropped until labels disappear
- full-slide image output described as editable
- a closing page that says only “Thank you”
