# Security

Report a vulnerability privately through [GitHub's private vulnerability reporting](https://github.com/TheMattHub/decks/security/advisories/new), not in a public issue.

What to know about this project:

- The scripts run locally on your machine, on files you point them at. They drive your own PowerPoint through COM and open no network connections, apart from `uv` fetching the Python dependencies declared in each script.
- A .pptx is a zip of XML and media from whoever made it. The scripts read only the files you give them and never execute macros; still, treat decks from unknown sources as you would any document.
- Template paths are read from `~/.claude/decks.json` on your machine; keep company templates out of repositories.
