# Production Routes

Use this reference to choose the production system before designing pages. The
route controls editability, fidelity, artifacts, and QA.

## Route Matrix

| Route | Output | Editability | Best use | Main risk |
|---|---|---|---|---|
| Native | Editable PPTX | High | Business, academic, teaching, standard decks | Generic visual language |
| Template | PPTX using supplied design | High within template | Corporate decks, branded reports | Breaking the template |
| Image-first | Image-per-slide PPTX | None or annotation only | Visual concepts, campaigns, cinematic stories | Typos and expensive rework |
| Hybrid | Native text + controlled visual assets | Key text high | Template/image reproduction, consulting visuals | Misalignment between layers |
| HTML | Static web deck | Source-level | Presenter view, motion, links, responsive delivery | Poor PowerPoint handoff |
| Motion/video | Images/video/self-running sequence | Low | Launch reels, recorded presentations | Slow, costly, harder to revise |
| Fidelity | Mixed native/vector/image with locked plan | High for information layer | “1:1,” blueprint matching, complex diagrams | Cost and quality-control overhead |

## Route Selection

Answer these in order:

1. **Must another person edit the file in PowerPoint?**
   - Yes, paragraphs/charts: Native, Template, Hybrid, or Fidelity.
   - No, browser/video is acceptable: HTML or Motion.
2. **Is a PPTX template the visual authority?**
   - Yes: Template.
3. **Are source pages already images and must they be reconstructed?**
   - Yes: Hybrid or Fidelity.
4. **Does visual quality outweigh editability?**
   - Yes: Image-first.
5. **Is motion part of the deliverable?**
   - Yes: HTML or Motion.
6. **Is “exact,” “1:1,” or an approved reference frame mandatory?**
   - Yes: Fidelity.

Otherwise default to Native. Do not ask the user to choose among technical route
names when their requirements already determine the route.

## Shared Planning Gate

Before any route:

- define audience, outcome, setting, duration, reading mode, and final format
- write a slide roster in the structured plan
- run the ghost-deck test
- establish evidence and source gaps
- choose one visual system and one density mode
- state any route limitation that affects the deliverable

Validate the plan:

```bash
python3 scripts/validate_deck_plan.py deck-plan.json --ghost
```

## Native

Use when the deck can be authored from a clean design system.

Build:

```bash
python3 scripts/build_pptx.py deck.json deck.pptx
```

Artifacts:

- `deck-plan.json`
- `deck.json`
- final PPTX
- render directory and contact sheet

Gate:

- no image-only text where editability is required
- no repeated page grammar merely because it is easy to generate
- every chart has units, labels, takeaway, and source

## Template

Use when a supplied `.pptx` or `.potx` defines the visual authority.

Preconditions:

- source file exists
- user permits using selected pages from the template
- output remains a copy

Workflow:

1. Render the source deck.
2. Inspect object structure.
3. Map template pages to cover, agenda, section, content, data, and closing
   roles.
4. Identify promotional or unusable pages.
5. Replace text and chart data without flattening paragraph formatting.
6. Remove orphaned placeholder groups.
7. Render and compare.

Inspection:

```bash
python3 scripts/inspect_pptx.py source.pptx --json > template-inventory.json
```

Hard rules:

- preserve the template's positions, font scale, colors, and master/layout
- replace run text rather than assigning a whole text frame when formatting
  must survive
- remove unused visual groups when the corresponding content is removed
- do not leave sample names, placeholder copy, sections, or logos that do not
  belong to the new deck

## Image-First

Use when every page can be a final raster image and the user accepts that the
text is not separately editable.

Required artifacts:

- `deck-plan.json`
- one prompt/source record per slide
- final selected image per slide
- speaker notes
- image-based PPTX

Recommended gate:

1. outline
2. one representative content-page prompt
3. one sample image
4. inspect text accuracy, hierarchy, density, and visual system
5. generate remaining pages
6. inspect every image before assembly
7. assemble

Assembly:

```bash
python3 scripts/assemble_image_pptx.py deck-plan.json image-deck.pptx --images-dir slides
```

Hard rules:

- do not claim editable text
- do not use a cover alone as the style sample; use a real content page
- do not let the model invent data or citations
- do not accept illegible or misspelled text merely because the layout is
  attractive
- keep the same aspect ratio and visual system across the deck

## Hybrid

Use when visual fidelity and text editability both matter.

Layer model:

- **information layer:** native text, native key numbers, simple charts,
  citations, page numbers
- **visual layer:** tightly cropped non-text images, official logos/product
  imagery, vector paths, complex textures, or other assets that cannot be
  reproduced faithfully as ordinary shapes

Do not put main titles, paragraphs, chart labels, sources, or conclusions in
the image layer.

The approved plan must state:

- which elements are native
- which visual assets are allowed
- why each asset is necessary
- how text will be masked/replaced if the reference already contains text
- which page is the representative sample

## HTML

Use when browser-first presentation, presenter mode, navigation, responsive
fitting, or sophisticated motion matters more than PowerPoint editability.

Read [html-motion.md](html-motion.md).

Hard rules:

- fixed 16:9 or user-specified slide stage
- one theme system and one layout family per deck
- speaker notes hidden from the audience
- URL/asset paths verified
- print/PDF output checked if the deck will be shared
- no claim that HTML is directly editable in PowerPoint

## Motion / Video

Use only when the user asks for recorded/self-running playback, animated
transitions, or a video deliverable.

The minimum pipeline is:

1. final slides approved as images or HTML frames
2. motion plan per transition
3. generated or coded transition clips
4. consistent resolution/fps/audio
5. interactive player or full MP4
6. playback inspection, not only frame inspection

Do not begin motion production until the static deck is accepted. A broken
static slide cannot be saved by a transition.

## Fidelity

Use for “match this blueprint,” complex diagrams, image-to-editable conversion,
or consulting pages with strong visual semantics.

This route has two equal gates:

1. **Information editability:** key text, figures, labels, and simple charts
   remain native.
2. **Visual semantics:** approved layout, surface system, hierarchy, geometry,
   density, and relationships remain recognizable.

Neither gate wins by default.

### Gate Sequence

1. Plan and ghost deck.
2. Sample content page.
3. Approve sample and freeze style.
4. Create scene reconstruction notes for every page.
5. Produce pages one at a time or in small controlled batches.
6. Render.
7. Compare with the approved reference.
8. Repair the current page.
9. Merge only after pages pass.
10. Render the merged deck again.

### Page Record

Keep a compact page record:

```json
{
  "slide": 1,
  "role": "evidence",
  "reference": "approved/01.png",
  "native": ["title", "labels", "chart", "source"],
  "image_assets": ["complex logo crop"],
  "tolerances": {"position_px": 4, "text_box_px": 5},
  "status": "passed"
}
```

Avoid a giant mandatory schema for ordinary decks. Use the full page record only
when the user requests exact fidelity or a visual blueprint governs production.

### Visual Comparison

```bash
python3 scripts/compare_renders.py approved/ work/renders --out-dir work/compare
```

Inspect:

- background and surface system
- title position and scale
- chart geometry and labels
- icon/node/label anchors
- text wrapping and container fit
- image crops
- bottom conclusion/SO-WHAT area
- density and reading order

Structural QA can pass while fidelity still fails.

## Stop Conditions

Stop and report a blocker when:

- required template or source image is unavailable
- the selected model/provider cannot be used and no approved local route exists
- a required image-to-editable conversion would not preserve key text
- a page cannot fit at the agreed density without becoming unreadable
- a visual comparison shows unresolved High/Critical differences
- the final file does not open or render

Do not silently downgrade the route, substitute a different output type, or
claim success from code generation alone.
