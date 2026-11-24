# HTML and Motion Decks

Use this route when browser delivery, presenter mode, motion, or video is part
of the requested experience.

## HTML Deck

### Use When

- the deck will be presented in a browser
- speaker view, timers, overview, links, or navigation matter
- rich CSS/WebGL/canvas motion is part of the presentation
- a static file or hosted URL is a valid handoff

Do not use HTML when the client must edit individual slides in PowerPoint.

### Architecture

Use a fixed slide stage:

```text
viewport wrapper
└── fixed 1920×1080 stage
    └── one slide section per logical page
```

Scale the whole stage uniformly. Do not reflow slide content at mobile
breakpoints. Letterboxing is acceptable; a changed composition is not.

Keep the deck source modular:

- one theme/tokens layer
- one slide/layout layer
- one runtime for navigation, overview, presenter mode, and deep links
- local image/font assets with verified relative paths

### Visual System

- one theme across the deck
- one type scale
- 2–4 layout families, varied by page role
- one or two motion recipes used with intent
- no animated background that competes with reading
- `prefers-reduced-motion` support

### Speaker Mode

Speaker-only material must never be visible on the audience slide.

Provide:

- current/next preview or stable page notes
- script/notes panel
- elapsed time and slide count
- keyboard navigation
- a clear recovery path if the presenter window is blocked

### Print/PDF

If the deck will be shared:

- each slide must become one page
- backgrounds and images must print
- no audience-hidden controls or notes should appear
- page size follows the stage aspect ratio

### HTML QA

- open the actual file/URL
- verify every asset loads
- use arrow keys, overview, deep links, fullscreen, and presenter mode
- inspect current and next slides at 16:9
- test one narrow viewport; the stage may letterbox but must not reflow
- print/PDF every page
- inspect text truncation and scrollbars

## Motion / Video Deck

### Use When

- the user asks for recorded, self-running, or video-first output
- transitions carry a narrative or product reveal
- image or HTML slides are already approved

### Preconditions

- static deck is approved
- motion adds meaning rather than hiding weak pages
- target duration, resolution, fps, aspect ratio, and audio policy are known

### Pipeline

1. Freeze the static slide sequence.
2. Mark which transitions deserve motion.
3. Write a transition brief for each cut:
   - what moves
   - what stays fixed
   - start/end state
   - duration
   - camera behavior
   - text stability requirement
4. Generate/code only those transitions.
5. Normalize all clips to the same codec, resolution, fps, color space, and
   audio layout.
6. Build an interactive player or full video.
7. Play the result end to end.

Simple fades and cuts do not require generative video. Use them when they are
the clearer solution.

### Transition Prompt

```text
Start from the approved slide image A and end on slide image B.
Keep the main title readable and stationary.
Move the central diagram from left to right while the background surface remains
stable.
Use a single continuous camera move, no cuts, no new text, no logos, no invented
objects, no warped typography.
Duration: 2.0 seconds.
```

### Video QA

Inspect playback, not only frames:

- transition begins/ends on the intended slides
- text remains readable and stable
- timing matches narration
- no flicker, banding, warping, or duplicated UI
- all clips use the same frame rate and audio sample rate
- final video duration matches the plan

### Output Contract

Report:

- static slide directory
- motion manifest/prompts
- raw transition clips
- final video or interactive player
- resolution, fps, duration, audio policy
- any transition not verified
