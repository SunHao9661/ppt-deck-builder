# Research Basis

This skill is an original synthesis. It does not copy templates, prose, or code
from the projects below.

## Projects Reviewed

Star counts were observed on 2026-11-24 via the GitHub API.

| Project | Stars | Route | Useful pattern adopted |
|---|---:|---|---|
| `anthropics/skills` (`skills/pptx`) | 199,599 repo stars | Native PPTX | required text/visual/file QA; template safety; native charts and notes |
| `hugohe3/ppt-master` | 79,687 | Native PPTX + HTML/export | deterministic route selection; native shapes/charts; template and edit separation; visual review |
| `zarazhangrui/frontend-slides` | 32,824 | HTML | show visual choices rather than asking users to describe taste; fixed stage |
| `op7418/guizang-ppt-skill` | 30,329 | HTML | one locked visual system; layout roles; rendered checks; presenter requirements |
| `chuspeeism/dashi-ppt-skill` | 11,053 | HTML + PPTX | constrained template filling and final accept/repair loop |
| `lewislulu/html-ppt-skill` | 9,200 | HTML | theme/layout separation and speaker-note discipline |
| `ningzimu/codex-ppt-skill` | 8,136 | Image-based PPTX | outline/style/sample gates and explicit image-led limitation |
| `GordenSun/GordenPPTSkill` | 4,222 | Native PPTX | do not break template layout; remove placeholders; render final file |
| `op7418/NanoBanana-PPT-Skills` | 3,382 | Image + video | image-first pages, transition prompts, interactive player, final video assembly |
| `GordenSun/GordenSuperPPTSkills` | 2,301 | Image → editable PPTX | image generation → background/frame/icon/text reconstruction pipeline |
| `crazyykhllc-bit/CyberPPT` | 1,937 | High-fidelity PPTX | SCR/evidence chain; content lock; editable text + complex visual assets; render comparison |
| `Gabberflast/academic-pptx-skill` | 1,671 | Native PPTX | action titles, ghost deck, evidence discipline, academic citations |

## Recurring Principles

The strongest projects independently converge on:

1. define audience and communication outcome first
2. plan the slide roster before styling
3. keep one visual system across the deck
4. vary composition by page role
5. keep native editability when the handoff requires it
6. render and inspect actual slide pixels
7. treat speaker notes as a separate delivery layer
8. fail visibly when a required asset/backend/layout gate is missing
9. choose production route from handoff and editability, not from fashion
10. keep an approved image or blueprint as the comparison baseline
11. treat native information and visual fidelity as separate, simultaneous gates
12. keep static approval before video transitions or hybrid reconstruction

## What This Skill Adds

- one route selector for native, template, image, hybrid, HTML, motion, and
  high-fidelity work
- a small, portable native PPTX builder that does not depend on one vendor
- a deck-plan validator, image-deck assembler, template inspector, structure
  audit, renderer, and reference-vs-output comparison
- high-density/SCR and evidence/source guidance without tying the skill to one
  model provider or one generation engine
- content and design rules separated into progressive references
- no bundled third-party templates or generated art, avoiding license coupling

The reviewed repositories use different and sometimes restrictive licenses.
This skill therefore incorporates high-level workflow principles and original
code only; it does not redistribute their templates, images, prompts, or source.

## Reference Links

- https://github.com/anthropics/skills
- https://github.com/hugohe3/ppt-master
- https://github.com/zarazhangrui/frontend-slides
- https://github.com/op7418/guizang-ppt-skill
- https://github.com/chuspeeism/dashi-ppt-skill
- https://github.com/lewislulu/html-ppt-skill
- https://github.com/ningzimu/codex-ppt-skill
- https://github.com/op7418/NanoBanana-PPT-Skills
- https://github.com/GordenSun/GordenPPTSkill
- https://github.com/GordenSun/GordenSuperPPTSkills
- https://github.com/crazyykhllc-bit/CyberPPT
- https://github.com/Gabberflast/academic-pptx-skill
