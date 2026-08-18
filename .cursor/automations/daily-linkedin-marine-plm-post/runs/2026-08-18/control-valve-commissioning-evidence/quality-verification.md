# Quality Verification — control-valve-commissioning-evidence (2026-08-18)

## Post QA (per claude-review-prompt.md Step 2)

| Check | Result |
|-------|--------|
| English-only, professional consultative tone | PASS |
| Strong hook/title with purposeful emojis (🔧 title, 🔎 points, 🎯 takeaway) | PASS |
| 3–5 short paragraphs with concrete execution context | PASS (3 body paragraphs) |
| 3 practical application points | PASS |
| 5–8 relevant hashtags | PASS (7 hashtags) |
| Not promotional — decision-oriented | PASS (pain: sea failure after quay stroke test; ownership: acceptance limits; proof gates: loop baseline → handover) |
| Concrete execution context | PASS (control-valve commissioning, HART configuration diff, fail-safe trip proof, class-witness handover evidence) |
| Embedded ARM rules | N/A (non-ARM topic; last ARM post 2026-08-13, 1-in-4 rotation satisfied) |

## Image QA (per prompt.md rules and claude-review-prompt.md Step 3)

| Gate | Result |
|------|--------|
| professional-grade | PASS — executive B2B engineering infographic matching commissioning-gates reference style |
| dimensions | PASS — 1080 × 1350 portrait (proportional 900×1350 scale + seamless deep-navy side padding, no distortion) |
| text-fit | PASS — all labels inside cards, chips, and panels; no clipped titles |
| overlap QA | PASS — no thumbnail/label collisions; pipeline strip, chips, and HUD labels clear of scene subjects |
| leader-lines | PASS — cyan TRACED LINK terminates on the HART positioner; amber dashed HOLD/ESCALATE terminates on the amber warning marker |
| solution-overview | PASS — 7-gate pipeline with bullets, photorealistic engine-room loop-check center scene, LOOP EVENTS legend, PYTHON AUTOMATION sidebar, 3 value pillars, amber footer takeaway |
| photo-diversity | N/A — single central photorealistic scene per commissioning-gates reference layout (no card thumbnails) |
| C2PA strip | PASS — `caBX`, `c2pa`, `jumb` markers absent from PNG bytes |

## Image source

- Built-in image generation tool (per Image generation rule), reference-styled on
  `runs/2026-08-13/cortex-m-peripheral-driver-gates/cortex-m-peripheral-driver-gates-infographic.png`.
- Deterministic post-processing: proportional resize to 900×1350 and pad to 1080×1350 with edge color `#000C1C`.

## Verdict

READY — artifacts staged for local Claude QA (09:20) and mirror+post (09:35).
