# Contributing

Thanks for helping. Issues and pull requests are welcome.

## Setup

- Windows with PowerPoint, to run the COM helpers (`skills/making-presentations/scripts/PptxCom.ps1`).
- [uv](https://docs.astral.sh/uv/) for the Python scripts and tests.

## Tests

```powershell
uv run --with pytest --with python-pptx --with pillow --with imageio-ffmpeg pytest tests
pwsh -NoProfile -File tests/check_ps.ps1
```

The PowerPoint helpers need PowerPoint, which CI does not have: test COM changes by hand on a copy of a deck, and say in the pull request what you ran.

## Pull requests

- One change per pull request, with the reason in the description.
- Keep `PptxCom.ps1` ASCII-only and compatible with Windows PowerShell 5.1.
- A new function gets a row in `references/powerpoint.md`; a new workflow step a line in `SKILL.md`. Keep `SKILL.md` short: details go in the references.
- Never commit a company template, a client deck or anything built from one. Test material must be fictional (see `examples/build_demo.ps1`).
