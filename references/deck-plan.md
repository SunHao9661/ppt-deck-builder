# Deck Plan

Use one plan before native, template, image, hybrid, HTML, or motion production.
The plan contains decisions, not implementation details such as coordinates or
shape IDs.

## Minimal Shape

```json
{
  "title": "AI adoption review",
  "audience": "Executive committee",
  "outcome": "Approve the staged rollout plan",
  "core_message": "Adoption is strong enough to expand, but support capacity is the limiting factor",
  "route": "native",
  "duration_minutes": 15,
  "language": "en",
  "visual": {
    "style": "restrained business editorial",
    "density": "balanced",
    "aspect_ratio": "16:9",
    "motif": "large section numbers and evidence-first charts"
  },
  "slides": [
    {
      "id": "s01",
      "role": "cover",
      "title": "Expansion is justified, but support capacity must come first",
      "purpose": "Establish the decision",
      "visual": "quiet cover with the decision as the main line",
      "notes": "Open with the decision, not the project history."
    }
  ]
}
```

## Top-Level Fields

| Field | Required | Meaning |
|---|---|---|
| `title` | Yes | Deck or working title |
| `audience` | Yes | Specific audience |
| `outcome` | Yes | What they should understand, decide, or do |
| `core_message` | Yes | One sentence that must survive |
| `route` | Yes | `native`, `template`, `image`, `hybrid`, `html`, `motion`, or `fidelity` |
| `duration_minutes` | Recommended | Presentation or review time |
| `language` | Recommended | BCP 47 code or readable language |
| `visual` | Yes | Shared visual direction |
| `template` | Template route | Source `.pptx`, `.potx`, or design workspace |
| `editability` | Hybrid/fidelity | `native-full`, `key-text`, `image-only`, or equivalent plain description |
| `motion` | Motion route | Delivery, duration, transition policy |
| `slides` | Yes | Ordered page roster |

## Visual Fields

```json
{
  "style": "Swiss editorial with warm neutral surfaces",
  "density": "presentation | balanced | read-close | high-density",
  "aspect_ratio": "16:9",
  "palette": ["#F6F1E8", "#1D1B18", "#C4492F"],
  "motif": "one recurring visual device",
  "image_treatment": "documentary photo | clean product crop | diagram | none"
}
```

Use palette values only when they are known or approved. A descriptive style is
enough when color is not fixed.

## Slide Fields

| Field | Required | Meaning |
|---|---|---|
| `id` | Yes | Stable unique slide ID |
| `role` | Yes | cover, context, problem, evidence, comparison, process, decision, conclusion, etc. |
| `title` | Yes | Audience-facing action title |
| `purpose` | Yes | Job the slide performs in the argument |
| `points` | Content-dependent | Essential supporting statements |
| `evidence` | Data/evidence pages | Data, exhibit, or support |
| `caveat` | As needed | Limitation, uncertainty, or source risk |
| `so_what` | Recommended | Meaning or decision implication |
| `visual` | Recommended | Composition or visual form |
| `source` | As needed | Citation/source label |
| `notes` | Recommended | Speaker note, not visible slide copy |
| `image_prompt` | Image/hybrid | Self-contained image-generation brief |
| `image` | Image/hybrid | Final selected image path |
| `motion` | Motion | Transition or playback note |
| `editability` | Hybrid/fidelity | Page-specific exception or layer rule |

## Roles

Use roles to check rhythm and completeness:

- `cover`
- `context`
- `problem`
- `question`
- `method`
- `evidence`
- `comparison`
- `process`
- `timeline`
- `metrics`
- `case`
- `recommendation`
- `decision`
- `conclusion`
- `appendix`

A deck can omit cover or conclusion when the requested artifact is a fragment,
appendix, or single-page handoff. Otherwise their absence is a warning.

## Action Titles

Titles should express the claim. These are too weak for content pages:

- Overview
- Background
- Current situation
- Analysis
- Results
- Discussion
- Recommendations
- Next steps
- Thank you

A technical or navigational page may use one of these only when the page itself
is an agenda, section divider, appendix index, or deliberate neutral reference
page.

## Density

Choose one deck-level density:

- `presentation`: one claim, sparse copy, strong visual
- `balanced`: claim + concise support
- `read-close`: fuller evidence and explanation
- `high-density`: consulting/report pages with multiple evidence blocks,
  caveats, source notes, and a visible so-what

High density is not permission for tiny type. Density comes from structure,
not from shrinking copy.

## Route-Specific Requirements

### Template

Set `template` to the source deck or design workspace. Do not place detailed
shape IDs in the plan; discover them after rendering and inspection.

### Image

Every non-trivial page should have a self-contained `image_prompt`. The prompt
must include:

- exact aspect ratio
- palette/typography treatment
- required content and labels
- chart/diagram type
- safe areas
- exclusions: no watermark, no fake logo, no extra page chrome

Keep the prompt short enough to edit, specific enough to reproduce.

### Hybrid / Fidelity

Set `editability` explicitly. For each page, identify:

- native text and simple charts
- non-text visual assets
- why each visual asset cannot be represented faithfully as ordinary native
  shapes

### HTML

Specify the fixed stage ratio, presenter-note needs, navigation needs, and
whether print/PDF is required.

### Motion

Set the motion frame rate, resolution, duration, audio policy, and transition
intent. Approve static pages first.

## Validation

Run:

```bash
python3 scripts/validate_deck_plan.py deck-plan.json --ghost
```

The validator checks required fields, duplicate IDs, route-specific gaps,
generic titles, repeated visual grammar, evidence/source warnings, and timing
plausibility. It does not prove factual correctness or visual quality.

Use `--strict` only when warnings should block the workflow.
