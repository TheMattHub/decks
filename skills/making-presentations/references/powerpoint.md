# Finishing and editing in PowerPoint

## Tools

- **PowerPoint through COM** (Windows, `<skill>/scripts/PptxCom.ps1`) for anything PowerPoint itself must
  produce: videos and their play settings, transitions including Morph, build-ins, embedded fonts, and
  true renders.
- **python-pptx** for reading a deck, quick dumps, setting document properties (COM late binding fails
  on `BuiltInDocumentProperties`), and decks built on a machine without PowerPoint.
- **zipfile** (Python) to swap a picture or video inside the .pptx under its own name. This keeps the
  shape's crop, opacity, trim and start point, which re-inserting the media would lose.

## PptxCom.ps1 functions

| Function | Does |
|---|---|
| `Open-Deck path [-ReadOnly]` / `Save-Deck s [path]` / `Close-Deck s` | Open without a window. Read-write mode refuses while another program holds the file, and backs it up first. Save embeds fonts. Close leaves a PowerPoint that was already running open |
| `Find-Slide deck regex`, `Find-Shape slide regex`, `Get-Picture slide` | Find slides and shapes by text |
| `Set-Notes`, `Add-TextBox`, `Set-FitInto`, `Set-SwapPicture` | Notes, a fixed text box (colour as "#RRGGBB"), fit into a frame, swap a picture in its box |
| `Add-PictureFill slide file l t w h` | A photo cropped to fill a box exactly (full-bleed covers, image cards); the crop stays editable |
| `Add-Video slide file poster l t w h` | Embed a video that plays when the slide opens, with a poster frame |
| `Set-Transition slide effect dur`, `Add-BuildIn slide shapes effect trigger` | Native transitions (`$PpFade`, `$PpPushUp`, `$PpMorph`) and build-in groups |
| `Export-SlidePngs deck dir [width]` | One PNG per slide, to look at |
| `Test-DeckLeftovers deck [extra regexes]` | List of leftovers (placeholders, TODO, option labels, file paths, "Source:") in shapes, groups, tables and notes, as Slide, Where, Match, Text. Prints nothing |
| `Export-LayoutSheet deck out.png [cols] [thumbWidth]` | One PNG of every layout as a labelled thumbnail with each slot named. Works on a template opened `-ReadOnly`: temporary slides are added and deleted, the deck is never saved. Read it with `Get-TemplateProfile` |
| `Write-DeckDump deck file` | Text dump of shapes, boxes, text, notes and effects, to diff |
| `Get-DeckTemplatePath name [-Library]` | Path of a template registered in `~/.claude/decks.json`; `-Library` gives its registered slide library, and throws when there is none |
| `Get-TemplateProfile deck` | Markdown list of layouts, placeholders, theme colours and fonts |
| `New-DeckFromTemplate file [-KeepSlides]`, `Add-LayoutSlide deck regex` | A new, unsaved deck from a .pptx or .potx, without its sample slides (`-KeepSlides` keeps them, for a slide library); a slide on a named layout |
| `Set-SlideRole slide role`, `Get-SlideRole slide` | Tag a slide with a role in its COM `Name` (`role:<name>`); read it back (`$null` when untagged) |
| `Get-SlideLibrary deck` | One `{Role; Slide; Title; Shapes}` per tagged slide; `Shapes` lists each text shape's name and first 40 characters, so a script sees what to replace |
| `Add-FromLibrary deck role texts` | Duplicates the first slide of that role, puts the copy before the first pattern, untags it, and sets the text of each shape matching a key regex of the `$texts` hashtable. Throws, leaving no stray copy, when a key matches nothing. Returns the slide |
| `Remove-LibrarySlides deck` | Deletes every slide still tagged with a role (the unused patterns) |
| `Get-Placeholder`, `Set-PlaceholderText slide kind text [n]`, `Remove-EmptyPlaceholders` | Fill title, body, subtitle and picture placeholders, counted left to right |
| `New-TemplateFromTheme out dark light accents heading body` | A .potx from colours and fonts |

## Slide library

A library deck holds designed pattern slides, each tagged with a role (cover, case-study, big-number,
comparison, section, closing). Generating from it edits designed slides instead of laying out new ones.

- **Tag once**: open a copy of the reference deck, run `Set-SlideRole` on each pattern slide, delete the
  slides that are not patterns, save as `<name> library.pptx`. Register it in `~/.claude/decks.json` as
  `"library": "<path>"` beside `"path"` for that template; `Get-DeckTemplatePath name -Library` reads it.
- **Build**: `New-DeckFromTemplate <library> -KeepSlides`, then `Get-SlideLibrary` to see each pattern's shape names and text,
  then `Add-FromLibrary $deck "case-study" @{ '^Optimization' = "New title"; '^Goal' = "Goal:`n..." }` per
  slide (the keys are regexes on the pattern's current text), then `Remove-LibrarySlides`, then `Save-Deck`.
- The role lives in the slide's `Name`, so it survives hand edits and moves (checked across save and
  reopen). Built slides sit in order ahead of the patterns.

## A finishing script, in outline

```powershell
$ErrorActionPreference = "Stop"
. "$PSScriptRoot\PptxCom.ps1"
$s = Open-Deck "$PSScriptRoot\export.pptx" -ReadOnly          # the export stays unchanged
try {
    if ($s.Deck.Slides.Count -ne 12) { throw "expected the 12-slide export" }   # guards against a new export
    $answer = Find-Slide $s.Deck '^THE ANSWER$'
    Set-Transition $answer $PpMorph 0.7
    $film = $s.Deck.Slides.AddSlide($answer.SlideIndex + 1, $answer.CustomLayout)
    foreach ($sh in @($film.Shapes)) { $sh.Delete() }
    [void](Add-Video $film "..\video\intro.mp4" "..\video\intro.png" 0 0 $s.Deck.PageSetup.SlideWidth $s.Deck.PageSetup.SlideHeight)
    Set-Notes $film "5:00. Let the film run."
    Save-Deck $s "$PSScriptRoot\Talk.pptx"
} finally { Close-Deck $s }
```

Get handles to every slide you will touch before moving any slide: `MoveTo` shifts the indexes.

## Gotchas

- **Colours**: `.RGB` takes BGR. Use `ConvertTo-Bgr "#E8692C"`. Text in theme colours cannot be read
  back, so keep the palette as hex values in the script.
- **Geometry**: setting Left, Top, Width or Height can fail with "Specified cast is not valid". Wrap the
  value in `[single](...)`.
- **PowerShell traps**: a function that returns a TextRange gets unrolled into characters (`return , $x`).
  A single COM match is not an array (wrap the pipeline in `@(...)`). `"$n:"` does not parse (`"${n}:"`).
- **Encoding**: Windows PowerShell 5.1 reads a .ps1 without a BOM as ANSI. Keep scripts ASCII and
  write `[char]0x00B7` for non-ASCII characters.
- **Busy PowerPoint**: `RPC_E_CALL_REJECTED` or "RPC server unavailable" means PowerPoint is still
  closing. Wait until `Get-Process POWERPNT` is gone, then retry.
- **Hidden dialogs**: a font-embedding dialog can block a windowless save. Check the saved file's
  timestamp after each run.
- **Fonts**: install them for the user by copying them to `%LOCALAPPDATA%\Microsoft\Windows\Fonts`
  and adding an `HKCU\Software\Microsoft\Windows NT\CurrentVersion\Fonts` entry. Then save with fonts
  embedded and check that `ppt/fonts/` exists inside the .pptx. Embedding needs fonts whose licence
  allows it (OFL fonts do). For a glyph the font lacks (IBM Plex draws a check mark as a root sign),
  use Segoe UI Symbol.
- **Paths**: file names with apostrophes break `python -c` quoting. Put code in script files, and make
  exact text edits with the Edit tool rather than Python strings (`\a` and `\f` become control
  characters).

## Equations

Native, editable equations from LaTeX: write each equation as its own text-box paragraph, `$$...$$`
(other paragraphs in the box stay text), then run

```
uv run "<skill>/scripts/pptx_equations.py" IN.pptx OUT.pptx [--xsl path\to\MML2OMML.XSL]
```

It converts LaTeX through MathML to Office's equation format with Office's own `MML2OMML.XSL` (so
Office must be installed, or pass `--xsl`), in the structure PowerPoint itself writes; the original
`$$latex$$` shape stays as the fallback. Output opens without repair and stays editable.

Limits found: only slides are scanned, not layouts, masters or notes; the equation is centred in its box; a sum inside a fraction keeps its limits beside it;
the XSL drops `mspace`, so the script turns `\quad` and similar gaps into an em space; shapes already
in `mc:AlternateContent` are skipped, so a second run changes nothing. Style the boxes before the run
(size and colour are copied from the first run of text); a COM open and save afterwards keeps the
equations. If the XSL is missing or an equation will not convert, an image per equation is the fallback.

## Media and size

Size the deck for what it is for, and pick the goal from the brief: `present` when it is shown from
this file, `share` for Teams, OneDrive or a download link, `email` when it goes as an attachment.

```
uv run "<skill>/scripts/pptx_media.py" deck.pptx                            # report, largest first, files over 5 MB flagged
uv run "<skill>/scripts/pptx_media.py" deck.pptx small.pptx --goal share    # compressed copy for the goal
```

| Goal | Pictures | JPEG | Video CRF | Stays under |
|---|---|---|---|---|
| `present` (default) | 2560 px | 85 | 24 | no limit |
| `share` | 1920 px | 80 | 28 | 25 MB |
| `email` | 1280 px | 75 | 30 | 15 MB (mail encoding adds a third; Outlook stops at 20 MB) |

The copy keeps every part's name and format, so crops, trims and play settings hold. A file is replaced
only when it gets smaller; PNGs are optimised and videos re-encoded to H.264 with the same length.
`--max-px`, `--jpeg-quality` and `--crf` override the goal (`--crf 28` for small cards). If the copy is
still over the goal's size, the script says so: cut or shorten the largest files, or link long films.
ffmpeg comes with the script. Already-compressed decks shrink little at `present` (24.3 to 23.2 MB on an
already lean talk deck), so the report is the main use there: it shows what to cut.

- Compress videos before embedding them too: same settings, at their own aspect ratio, with a poster PNG.
- A video on a small card needs no full resolution: embed a 640 px copy of the same length.
- Long films (over about 30 s) get their own slide or go to the backup.
- On a deck with hand edits, write the copy, check it in PowerPoint, then replace the original.

## Hand edits

1. Ask whether the person has changed the file since the last script run.
2. Build a fresh copy to a scratch path. Dump both decks with `Write-DeckDump` and diff them. Report
   what was changed by hand.
3. Write the next numbered round (`edits_N.ps1`). It opens read-write, finds every target by its text,
   throws if a target is missing or the round has already been applied, edits and saves.
4. Record in the project's memory or README that the deck is now edited in place.

Small text or layout changes after the export can also be made by the user with the Claude in
PowerPoint add-in. That is safe with the edit scripts, because they find slides by text, but ask
whether the deck was changed before the next scripted round (step 1 already does). Reviewers report
drift on heavily customised templates.
