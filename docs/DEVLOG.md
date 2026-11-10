# Devlog

## Status

Branch `feat/review-round-1` (from main at ed50908): the fixes, workflow improvements and features from the
November 2026 review, all merged on the branch. Not pushed. Next: see the last line of the log.

## Bug register

| ID | Bug | Status |
|---|---|---|
| B01 | plugin.json still says 0.2.0, which pins installed users: two `feat` commits never reached them | fixed: 0.3.0 |
| B02 | `Add-FromLibrary` matches keys against text it has just written | fixed |
| B03 | `Find-Slide` and `Find-Shape` miss grouped shapes and table cells | fixed: one walker |
| B04 | PowerPoint lingers after `Close-Deck` (the "Busy PowerPoint" gotcha) | fixed: release on close, retry on open |
| B05 | `pptx_media.py`: videos never scaled, exit 0 over the size limit, ICC profile dropped | fixed |
| B06 | No text-replace helper: assigning `.Text` wipes hand formatting | fixed: `Edit-DeckText` |
| B07 | `Test-DeckLeftovers` flags source lines, has no exclude, misses review comments | fixed |
| B08 | Stale PNGs after `Export-SlidePngs`; header comment and version stamp; `icon` not in plugin.json | fixed |
| B09 | `Add-TextBox` with a whole-number size, then a fractional one, fails "Unable to cast" | fixed: `[single]` |
| B10 | A PowerPoint on its start screen counted as the script's own and was quit | fixed: ownership by process |
| B11 | `Stop-PowerPoint` and `Close-Deck` could end a PowerPoint a person had joined (single instance) | fixed: never with a visible window |
| B12 | Builds on multi-block elements never matched the export; the applier was not all-or-nothing | fixed |
| B13 | `Test-DeckLayout` missed overflow in shrink-on-overflow placeholders | fixed |

## Work items

| ID | Item | Status |
|---|---|---|
| W-A | Scripted review: `pptx_check.py` (file level), `Test-DeckLayout` (layout level) | done |
| W-B | Hand edits detected from a record written by `Save-Deck` (`Get-DeckHandEdits`) | done |
| W-C | Contact sheet of slides (`Export-ContactSheet`) | done |
| W-D | Cold read by a fresh agent in the review step (SKILL.md, step 6) | done |
| W-E | Checks that run in CI and without PowerPoint; local COM smoke test (`tests/com_smoke.ps1`) | done |
| F1 | Starter slide library for generated templates (`New-StarterLibrary`) | done |
| F2 | Finish manifest from the Slides source (`slides_manifest.py`, `Invoke-FinishManifest`) | done; not yet run on a real Slides deck |
| F3 | Media sized by footprint, `--fit` | done |
| F4 | `Export-DeckPdf`, with notes pages as the read-alone version | done |
| G | PowerPoint guards: never quit or stop an instance the helper did not start (`Stop-PowerPoint`) | done |

## Known limits

- `slides_manifest.py` and `Invoke-FinishManifest` are tested on fixture decks written from the Slides type's
  authoring guide, not on a real exported deck. The current Slides release does not write videos.
- `Export-DeckPdf -Notes` depends on the "Microsoft Print to PDF" printer and its paper size.
- Not covered by a test: the hidden-dialog branch of `Save-Deck`, and the `font` findings (no font at hand
  that cannot be embedded).
- Deck records (`<name>.dump.txt`) written before the rounding change show every shape line as edited once.

## Log

- 2026-11-10: review of the repo; plan written; branch created; B01 and the icon field fixed;
  `tests/leftover_cases.txt` added as the shared contract for the two leftover checks.
- 2026-11-10: wave 1 merged (PptxCom hardening, `pptx_media.py`, `pptx_check.py`, `slides_manifest.py`);
  checkpoint review: 8 must-fix and 6 minor findings, all reproduced by a failing test and fixed.
- 2026-11-10: wave 2 merged (guards, `Test-DeckLayout`, starter library, manifest applier); docs written.
- 2026-11-10: final whole-branch review: 4 must-fix and 7 minor findings, each reproduced by running it;
  one fix pass (B11-B13 and the minor ones); docs corrected where the review proved them false.
  NEXT: the user decides on merge to main, push and release (the marketplace name question is open).
