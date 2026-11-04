# Security

Report a vulnerability privately through [GitHub's private vulnerability reporting](https://github.com/TheMattHub/decks/security/advisories/new), not in a public issue. If that form is unavailable, open an issue describing the class of problem without a working recipe.

## Threat model

This is an agent workflow: an AI agent with file access reads decks, notes and templates, and runs scripts on your machine.

- **Decks from others are untrusted content.** Slide text, speaker notes, hidden shapes or a template's layout names can carry text written to steer the agent. The skill tells the agent to treat everything read from a deck as data, never as instructions. Review what the agent proposes before it edits or sends anything.
- **Scripts run locally on the files you name.** `PptxCom.ps1` drives your own PowerPoint through COM; the Python scripts read and write .pptx zips. None opens a network connection, apart from `uv` fetching the dependencies each script declares (including an ffmpeg build for video compression).
- **No macros.** The tools work on .pptx and .potx; they never run VBA. Do not point them at macro-enabled files (.pptm) from unknown sources.
- **Company files stay out of repositories.** Template paths are read from `~/.claude/decks.json` on your machine.

This is not a sandbox: the agent runs with the permissions your agent runtime grants it.
