# Devlog

## Status

Branch `feat/review-round-1` (from main at ed50908): the fixes, workflow improvements and features from the
November 2026 review. Not pushed. Next: see the last line of the log.

## Bug register

| ID | Bug | Status |
|---|---|---|
| B01 | plugin.json still says 0.2.0, which pins installed users: two `feat` commits never reached them | fixed: 0.3.0 |
| B02 | `Add-FromLibrary` matches keys against text it has just written | open |
| B03 | `Find-Slide` and `Find-Shape` miss grouped shapes and table cells | open |
| B04 | PowerPoint lingers after `Close-Deck` (the "Busy PowerPoint" gotcha) | open |
| B05 | `pptx_media.py`: videos never scaled, exit 0 over the size limit, ICC profile dropped | open |
| B06 | No text-replace helper: assigning `.Text` wipes hand formatting | open |
| B07 | `Test-DeckLeftovers` flags source lines, has no exclude, misses review comments | open |
| B08 | Stale PNGs after `Export-SlidePngs`; header comment and version stamp; `icon` not in plugin.json | open (icon fixed) |

## Work items

| ID | Item | Wave | Status |
|---|---|---|---|
| W-A | Scripted review: `pptx_check.py` (file level), `Test-DeckLayout` (layout level) | 1, 2 | open |
| W-B | Hand edits detected from a record written by `Save-Deck` | 1 | open |
| W-C | Contact sheet of slides | 1 | open |
| W-D | Cold read by a fresh agent in the review step | docs | open |
| W-E | Checks that run in CI and without PowerPoint; local COM smoke test | 1 | open |
| F1 | Starter slide library for generated templates | 2 | open |
| F2 | Finish manifest from the Slides source (`slides_manifest.py`, `Invoke-FinishManifest`) | 1, 2 | open |
| F3 | Media sized by footprint, `--fit` | 1 | open |
| F4 | `Export-DeckPdf`, with notes pages as the read-alone version | 1 | open |

## Log

- 2026-11-10: review of the repo; plan written; branch created; B01 and the icon field fixed;
  `tests/leftover_cases.txt` added as the shared contract for the two leftover checks.
  NEXT: wave 1 (PS-1, PY-media, PY-check, PY-manifest), then merge, focused tests, checkpoint review.
