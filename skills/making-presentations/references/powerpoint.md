# Finishing and editing in PowerPoint

## Tools

- **PowerPoint through COM** (Windows, `<skill>/scripts/PptxCom.ps1`) for anything PowerPoint itself must
  produce: videos and their play settings, transitions including Morph, build-ins, embedded fonts, and
  true renders.
- **python-pptx** for reading a deck, quick dumps, setting document properties (COM late binding fails
  on `BuiltInDocumentProperties`), and decks built on a machine without PowerPoint.
- **`pptx_check.py`** (Python, no PowerPoint needed) for the checks the file alone can answer; see "Checks".
- **zipfile** (Python) to swap a picture or video inside the .pptx under its own name. This keeps the
  shape's crop, opacity, trim and start point, which re-inserting the media would lose.

## PptxCom.ps1 functions

| Function | Does |
|---|---|
| `Open-Deck path [-ReadOnly]` | Open without a window. Read-write mode refuses while another program holds the file, and backs it up first (same extension). Retries for up to 20 s while a previous PowerPoint is still closing |
| `Save-Deck s [path]` | Saves with fonts embedded. Throws when the file was not written (a hidden dialog), warns per font that cannot be embedded, and writes the record `<name>.dump.txt` (hash, then the dump) beside a .pptx |
| `Close-Deck s` | Closes and releases PowerPoint, so a windowless one exits within seconds. Quits PowerPoint only when this process started it and no presentation is left |
| `Stop-PowerPoint` | Ends a hung PowerPoint only when this process started it, or when it has no window and no unsaved presentation; otherwise throws, naming them. Never `Stop-Process POWERPNT` yourself |
| `Find-Slide deck regex`, `Find-Shape slide regex`, `Get-Picture slide` | Find slides and shapes by text, inside groups and table cells too |
| `Edit-DeckText deck-or-slide old new` | Replaces the one literal, case-sensitive occurrence and keeps the run formatting (assigning `.Text` loses a bold word). Throws when it is missing (already applied, or changed by hand) or occurs more than once (names the slides) |
| `Set-Notes`, `Add-TextBox`, `Set-FitInto`, `Set-SwapPicture` | Notes, a fixed text box (colour as "#RRGGBB"), fit into a frame, swap a picture in its box |
| `Add-PictureFill slide file l t w h` | A photo cropped to fill a box exactly (full-bleed covers, image cards); the crop stays editable |
| `Add-Video slide file poster l t w h` | Embed a video that plays when the slide opens, with a poster frame |
| `Set-Transition slide effect dur`, `Add-BuildIn slide shapes effect trigger` | Native transitions (`$PpFade`, `$PpPushUp`, `$PpMorph`) and build-in groups |
| `Export-SlidePngs deck dir [width]` | One PNG per slide, to look at; old `slide*.png` in the folder are deleted first |
| `Export-ContactSheet deck out.png [cols] [thumbWidth]` | Every slide as a numbered thumbnail on one PNG, hidden slides marked: one image to triage a deck |
| `Export-DeckPdf deck out.pdf [-Notes]` | PDF of the shown slides; `-Notes` gives notes pages (slide plus notes: the read-alone version from the same source) through the "Microsoft Print to PDF" printer. The deck is untouched |
| `Test-DeckLeftovers deck [extra regexes] [exclude regexes]` | List of leftovers (placeholders, TODO, option labels, file paths) in shapes, groups, tables and notes, plus one `comments` entry per slide with review comments, as Slide, Where, Match, Text. An exclude regex drops a hit whose text matches. Prints nothing |
| `Test-DeckLayout deck [minPt]` | Layout findings as Slide, Where, Check, Detail: `small-text` (under 12 pt on a 960 pt slide), `overflow` (text taller than its fixed box), `off-slide`, `font` (cannot be embedded). Prints nothing |
| `Export-LayoutSheet deck out.png [cols] [thumbWidth]` | One PNG of every layout as a labelled thumbnail with each slot named. Works on a template opened `-ReadOnly`: temporary slides are added and deleted, the deck is never saved. Read it with `Get-TemplateProfile` |
| `Get-DeckDump deck` / `Write-DeckDump deck file` | Text dump of shapes, boxes, text (with the first run's font and size), notes and effects |
| `Get-DeckHandEdits s` | What a person changed since the last `Save-Deck`: nothing when the file matches its record, else the dump lines that differ (`- ` recorded, `+ ` now). Throws when there is no record |
| `Invoke-FinishManifest deck manifest.json [mediaDir]` | Applies a `slides_manifest.py` manifest (transitions, build-ins, videos) to a Slides export. Checks everything first; throws once listing every problem and changes nothing. Run once per export |
| `Get-DeckTemplatePath name [-Library]` | Path of a template registered in `~/.claude/decks.json`; `-Library` gives its registered slide library, and throws when there is none |
| `Get-TemplateProfile deck` | Markdown list of layouts, placeholders, theme colours and fonts |
| `New-DeckFromTemplate file [-KeepSlides]`, `Add-LayoutSlide deck regex` | A new, unsaved deck from a .pptx or .potx, without its sample slides (`-KeepSlides` keeps them, for a slide library); a slide on a named layout |
| `Set-SlideRole slide role`, `Get-SlideRole slide` | Tag a slide with a role in its COM `Name` (`role:<name>`); read it back (`$null` when untagged) |
| `Get-SlideLibrary deck` | One `{Role; Slide; Title; Shapes}` per tagged slide; `Shapes` lists each text shape's name and first 40 characters, so a script sees what to replace |
| `Add-FromLibrary deck role texts` | Duplicates the first slide of that role, puts the copy before the first pattern, untags it, and sets the text of each shape matching a key regex of the `$texts` hashtable. Throws, leaving no stray copy, when a key matches nothing. Returns the slide |
| `Remove-LibrarySlides deck` | Deletes every slide still tagged with a role (the unused patterns) |
| `Get-Placeholder`, `Set-PlaceholderText slide kind text [n]`, `Remove-EmptyPlaceholders` | Fill title, body, subtitle and picture placeholders, counted left to right |
| `New-TemplateFromTheme out dark light accents heading body` | A .potx from colours and fonts |
| `New-StarterLibrary template out.pptx` | A slide library for such a template: cover, section, cards, big-numbers, quote, timeline, closing, in the theme's colours and fonts |

## Slide library

A library deck holds designed pattern slides, each tagged with a role (cover, case-study, big-number,
comparison, section, closing). Generating from it edits designed slides instead of laying out new ones.

- **Tag once**: open a copy of the reference deck, run `Set-SlideRole` on each pattern slide, delete the
  slides that are not patterns, save as `<name> library.pptx`. Register it in `~/.claude/decks.json` as
  `"library": "<path>"` beside `"path"` for that template; `Get-DeckTemplatePath name -Library` reads it.
- **Build**: `New-DeckFromTemplate <library> -KeepSlides`, then `Get-SlideLibrary` to see each pattern's shape names and text,
  then `Add-FromLibrary $deck "case-study" @{ '^Optimization' = "New title"; '^Goal' = "Goal:`n..." }` per
  slide (the keys are regexes on the pattern's current text), then `Remove-LibrarySlides`, then `Save-Deck`.
- **A generated template has no designed slides**: `New-StarterLibrary $tpl "<name> library.pptx"` makes them, and
  the file is registered as the template's `library`. Every sample text starts with `[placeholder]`
  (keys such as `'^\[placeholder\] Card one$'`), so one left unreplaced is caught by `Test-DeckLeftovers`.
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

Transitions and build-ins need not be retyped: `slides_manifest.py` reads them from the Slides deck's files
(`claude-slides.md`) and `Invoke-FinishManifest $s.Deck manifest.json` applies them. Videos are still added
with `Add-Video`: the current Slides release does not write them.

## Gotchas

- **Colours**: `.RGB` takes BGR. Use `ConvertTo-Bgr "#E8692C"`. Text in theme colours cannot be read
  back, so keep the palette as hex values in the script.
- **Geometry**: setting Left, Top, Width or Height can fail with "Specified cast is not valid". Wrap the
  value in `[single](...)`.
- **PowerShell traps**: a function that returns a TextRange gets unrolled into characters (`return , $x`).
  A single COM match is not an array (wrap the pipeline in `@(...)`). `"$n:"` does not parse (`"${n}:"`).
  PowerShell remembers the argument type of a COM property set: pass sizes as `[single]`, or a later
  fractional value fails with "Unable to cast". Under Windows PowerShell with `$ErrorActionPreference = "Stop"`,
  a native command that writes to stderr (such as `uv`) stops the script.
- **Encoding**: Windows PowerShell 5.1 reads a .ps1 without a BOM as ANSI. Keep scripts ASCII and
  write `[char]0x00B7` for non-ASCII characters.
- **Busy PowerPoint**: `RPC_E_CALL_REJECTED` or "RPC server unavailable" means PowerPoint is still
  closing. The functions that start PowerPoint retry for up to 20 s, so this should no longer reach a script.
- **The person's PowerPoint**: PowerPoint is one instance per Windows session, so a script shares it with
  the decks the person has open. The helpers never quit or stop an instance they did not start. For a hung
  one use `Stop-PowerPoint`, which refuses while there is unsaved work or a visible window; then tell the
  person instead of forcing it.
- **Hidden dialogs**: a font-embedding dialog can block a windowless save. `Save-Deck` throws when the file
  was not written.
- **PDF**: PowerPoint writes only the first PDF saved from an open deck (later saves silently write nothing)
  and ignores the notes-page print setting, so `Export-DeckPdf` works on a throwaway copy and prints notes
  pages to "Microsoft Print to PDF" at that printer's paper size.
- **Measured text**: `TextRange.BoundHeight` is measured on the slide, so a rotated box's text looks taller
  than the box. Changing the slide size through COM does not scale font sizes.
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

| Goal | Pictures | Videos | JPEG | Video CRF | Stays under |
|---|---|---|---|---|---|
| `present` (default) | 2560 px | 1920 px | 85 | 24 | no limit |
| `share` | 1920 px | 1280 px | 80 | 28 | 25 MB |
| `email` | 1280 px | 960 px | 75 | 30 | 15 MB (mail encoding adds a third; Outlook stops at 20 MB) |

The pixel values are for media shown across the whole slide. Each picture and video is sized to the pixels
its slides actually show (crops, groups and use on several slides counted, never upscaled), so a small card
gets a small file; media whose shown size is unknown (layout, master, background, fill) keep the full value,
and `--no-footprint` turns this off.

The copy keeps every part's name and format, so crops, trims and play settings hold. A file is replaced
only when it gets smaller; PNGs are optimised and videos re-encoded to H.264 with the same length.
`--max-px`, `--max-video-px`, `--jpeg-quality` and `--crf` override the goal. If the copy is still over the
goal's size, the script says so and exits 1. `--fit` then rewrites the copy from the original with smaller
settings, up to four steps (pixels x 0.8, JPEG quality -5, CRF +2 per step), and prints the settings it kept;
if that is not enough, cut or shorten the largest files, or link long films.
ffmpeg comes with the script. Already-compressed decks shrink little at `present` (24.3 to 23.2 MB on an
already lean talk deck), so the report is the main use there: it shows what to cut.

- Compress videos before embedding them too: same settings, at their own aspect ratio, with a poster PNG.
- Long films (over about 30 s) get their own slide or go to the backup.
- On a deck with hand edits, write the copy, check it in PowerPoint, then replace the original.

## Checks

Run both before presenting, and after every edit round; clear each finding or waive it with a reason.

```
uv run "<skill>/scripts/pptx_check.py" deck.pptx --numbers slides/key_numbers.md --slot 20 --budget 12
```

`pptx_check.py` needs only the file, so it also runs without PowerPoint. It prints the slide counts (main and
backup), hidden slides, embedded fonts and document properties (look at author and title before sending a deck
out), then one `slide N  check  detail` line per finding, and exits 1 when there is any:
`leftover` (drafting residue in shapes, groups, table cells and notes), `comment` (open review comments),
`number` (a number on a slide that is not in the numbers file; a bare single digit is skipped), `timing` (main
slides' notes start with `m:ss.`, the starts rise, the last one leaves a minute, and the words fit the time at
150 a minute) and `budget`. `--exclude REGEX` waives a hit whose text matches; `--pattern REGEX` adds a
leftover pattern. Main slides are those before the first slide titled Backup, Appendix or Annex
(`--backup-title`).

`Test-DeckLayout $s.Deck` needs PowerPoint, because only it lays the text out: small text, overflow, shapes
off the slide, fonts that cannot be embedded.

## Hand edits

1. `Get-DeckHandEdits $s` right after `Open-Deck`: it compares the deck with the record the last `Save-Deck`
   wrote, so nobody has to remember. Report what was changed by hand. (No record: the deck was never saved by
   `Save-Deck`; ask, and save once to start one.)
2. Read the changed slides before editing: the person's wording wins.
3. Write the next numbered round (`edits_N.ps1`). It opens read-write and changes text with `Edit-DeckText`,
   which throws if the old text is missing (the round was already applied, or a person changed it) or
   ambiguous; other targets are found by their text. Then it saves.
4. Record in the project's memory or README that the deck is now edited in place.

Small text or layout changes after the export can also be made by the user with the Claude in
PowerPoint add-in. That is safe with the edit scripts, because they find slides by text, but ask
whether the deck was changed before the next scripted round (step 1 already does). Reviewers report
drift on heavily customised templates.
