# Contributing

Thanks for helping. This plugin has a narrow aim, and knowing it saves you effort: **a general workflow and tools for making presentations, usable by anyone, with any company template.** A contribution is judged by fit to that aim first, execution second. If you are unsure, open an Idea in [Discussions](https://github.com/TheMattHub/decks/discussions) before building; it costs nothing.

## What gets merged

- **Fixes with the failing case shown**: what you ran, on what kind of deck, what happened.
- **General workflow improvements**: a step, check or reference that helps any kind of deck (consultancy, research, marketing, internal).
- **Helpers that remove manual PowerPoint work** many decks need, documented with a row in `references/powerpoint.md`.
- **Docs that close a real gap**: a setup step, a gotcha you hit, a stale reference.

## What gets declined

- **Company- or client-specific content**: templates, brand rules, colour sets or decks of a particular organisation. Register yours locally in `~/.claude/decks.json`; never commit it.
- **A second copy of the workflow** for another agent runtime. `SKILL.md` already uses the portable Agent Skills format; adapt it in your fork and point back here, so improvements reach you.
- **Speculative features**: argue from a deck problem you actually had.
- **Bundles**: one concern per pull request. Split bundles are reviewed faster.

## Claims get verified

Show the failure on the real path: the documented command or helper, run on a real-shaped deck (a fictional one is fine), not a hand-built value fed straight into a function. A test that fails before your fix and passes after it is necessary, but not sufficient.

## Setup and tests

- Windows with PowerPoint, to run the COM helpers (`skills/making-presentations/scripts/PptxCom.ps1`).
- [uv](https://docs.astral.sh/uv/) for the Python scripts and tests.

```powershell
uv run --with pytest --with python-pptx --with pillow --with imageio-ffmpeg --with latex2mathml --with lxml pytest tests
pwsh -NoProfile -File tests/check_ps.ps1
```

CI has no PowerPoint: test COM changes by hand on a copy of a deck and say in the pull request what you ran and what you saw (a slide render helps).

## Practical notes

- **Pull-request titles are the release notes**: write them for a user and start with a [Conventional Commits](https://www.conventionalcommits.org/) type (`feat:`, `fix:`, `docs:` ...).
- Keep `PptxCom.ps1` ASCII-only and compatible with Windows PowerShell 5.1.
- A new function gets a row in `references/powerpoint.md`; a new step a line in `SKILL.md`. Keep `SKILL.md` short: details go in the references.
- Test material must be fictional (see `examples/build_demo.ps1`). Never attach a company template or client deck to an issue or pull request.
