# /// script
# requires-python = ">=3.10"
# dependencies = ["python-pptx"]
# ///
"""Check a .pptx for what a script can judge from the file alone (no PowerPoint needed, so CI can run it).

    uv run pptx_check.py DECK.pptx [--numbers FILE] [--slot MINUTES] [--budget N]
                                   [--pattern REGEX ...] [--exclude REGEX ...] [--backup-title REGEX]

Prints info lines (slide counts, hidden slides, embedded fonts, document properties), then one line per
finding, `slide N  <check>  <detail>`, and exits 1 when there is any finding. Checks:

  leftover  drafting residue in shapes (groups too), table cells and notes: lorem ipsum, TODO/TBD/FIXME,
            [insert / [placeholder / [__], "click to add", "made with claude", "option A", file paths.
            --pattern adds a regex; --exclude drops a hit whose slide text matches it (waives a false hit).
  comment   a slide with review comments (legacy or modern), with their count.
  number    (--numbers FILE) a number in slide text (not notes) that FILE does not contain. Values are
            compared, not strings: 1,200 = 1 200 = 1200; 3,5 = 3.5; 38% = 38 %. A lone digit (Step 3) and
            digits inside a word (Q3, N01) are skipped. Years are checked like any number.
  timing    notes start with the slide's start time, `m:ss.` (e.g. `2:30.`). A main slide without one, one
            that starts no later than the previous, with --slot M a last main slide starting after M minus
            one minute, or notes that take over 1.3 times the slide's time to say at 150 words a minute.
  budget    (--budget N) more than N main slides (the cover counts).

Main slides are those before the backup divider: the first slide whose title (else first text) matches
--backup-title (default `^(backup|appendix|annex)\\b`, case-insensitive). Without a divider all are main.
"""
import argparse
import re
import sys
import zipfile
from decimal import Decimal

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

# Shared with Test-DeckLeftovers in PptxCom.ps1 and tests/leftover_cases.txt: change them together.
LEFTOVER_PATTERNS = [
    r"lorem|ipsum",
    r"\bTODO\b|\bTBD\b|\bFIXME\b",
    r"\[insert|\[placeholder|\[€__\]|\[[A-Z ]*__\]",
    r"click to add",
    r"made with claude",
    r"\b(option|variant) [A-D]\b",
    r"[A-Za-z]:\\|\.(py|ps1|csv|json|md)\b",
]
COMMENT_RELS = ("/relationships/comments",)  # legacy .../2006/relationships/comments and the 2018/10 modern one
START = re.compile(r"^\s*(\d{1,2}):(\d{2})\.")
# A separator followed by exactly three digits groups thousands; any other comma or dot is the decimal mark.
NUMBER = re.compile(r"(?<!\w)(\d{1,3}(?:[ \u00a0\u202f,.]\d{3})+(?!\d)(?:[.,]\d+)?|\d+(?:[.,]\d+)?)(?!\w)")
LINEBREAKS = re.compile(r"[\r\n\v]+")
GROUPED = re.compile(r"^\d{1,3}(?:[ \u00a0\u202f,.]\d{3})+$")


def numbers_in(text):
    """(token as written, value) for each number worth checking."""
    out = []
    for m in NUMBER.finditer(text):
        tok = m.group(1)
        if re.fullmatch(r"\d", tok):  # "Step 3"
            continue
        digits = re.sub(r"[   ,.]", "", tok) if GROUPED.match(tok) else tok.replace(",", ".")
        out.append((tok, Decimal(digits).normalize()))
    return out


def shape_texts(shapes):
    """(where, text) for shapes, groups (nested) and table cells; slide-number placeholders are skipped."""
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from shape_texts(sh.shapes)
        elif sh.has_table:
            for row in sh.table.rows:
                for cell in row.cells:
                    yield f"table {sh.name}", cell.text_frame.text
        elif sh.has_text_frame:
            if sh.is_placeholder and sh.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER:
                continue
            yield f"shape {sh.name}", sh.text_frame.text


def notes_of(slide):
    return slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ""


def title_of(slide):
    t = slide.shapes.title
    if t is not None and t.has_text_frame and t.text_frame.text.strip():
        return t.text_frame.text.strip()
    return next((x.strip() for _, x in shape_texts(slide.shapes) if x.strip()), "")


def comment_count(slide):
    n = 0
    for rel in slide.part.rels.values():
        if rel.reltype.endswith(COMMENT_RELS) and not rel.is_external:
            blob = rel.target_part.blob.decode("utf-8", "replace")
            n += len(re.findall(r"<(?:\w+:)?cm[\s>/]", blob))
    return n


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("--numbers", help="text file with the numbers the deck may use")
    ap.add_argument("--slot", type=float, help="talk slot in minutes")
    ap.add_argument("--budget", type=int, help="most main slides allowed")
    ap.add_argument("--pattern", action="append", default=[], help="extra leftover regex (repeatable)")
    ap.add_argument("--exclude", action="append", default=[], help="drop leftover hits whose text matches")
    ap.add_argument("--backup-title", default=r"^(backup|appendix|annex)\b")
    a = ap.parse_args()

    prs = Presentation(a.deck)
    slides = list(prs.slides)
    backup = re.compile(a.backup_title, re.I)
    divider = next((i for i, s in enumerate(slides) if backup.search(title_of(s))), len(slides))
    main_slides = slides[:divider]

    with zipfile.ZipFile(a.deck) as z:
        fonts = [n for n in z.namelist() if n.startswith("ppt/fonts/") and not n.endswith("/")]
    hidden = [str(i) for i, s in enumerate(slides, 1) if s._element.get("show") == "0"]
    cp = prs.core_properties
    print(f"slides: {len(main_slides)} main, {len(slides) - len(main_slides)} backup")
    print(f"hidden: {', '.join(hidden) or 'none'}")
    print(f"embedded fonts: {', '.join(fonts) or 'none'}")
    print(f"title: {cp.title!r}  author: {cp.author!r}  last modified by: {cp.last_modified_by!r}")

    findings = []

    def add(n, check, detail):
        findings.append((n, check, detail))

    patterns = [re.compile(p, re.I) for p in LEFTOVER_PATTERNS + a.pattern]
    excludes = [re.compile(p, re.I) for p in a.exclude]
    allowed = None
    if a.numbers:
        with open(a.numbers, encoding="utf-8") as f:
            allowed = {v for _, v in numbers_in(f.read())}

    for n, s in enumerate(slides, 1):
        texts = list(shape_texts(s.shapes))
        notes = notes_of(s)
        for where, text in texts + [("notes", notes)]:
            for p in patterns:
                m = p.search(text)
                if m and not any(e.search(text) for e in excludes):
                    add(n, "leftover", f"{where}: {m.group(0)!r} in {LINEBREAKS.sub(' / ', text)[:60]!r}")
        c = comment_count(s)
        if c:
            add(n, "comment", f"{c} review comment{'s' if c != 1 else ''}")
        if allowed is not None:
            seen = set()
            for _, text in texts:
                for tok, v in numbers_in(text):
                    if v not in allowed and v not in seen:
                        seen.add(v)
                        add(n, "number", f"{tok} is not in {a.numbers}")

    # Timing: start times from the notes of main slides.
    starts = []
    for s in main_slides:
        m = START.match(notes_of(s))
        starts.append(int(m.group(1)) * 60 + int(m.group(2)) if m else None)
    end = a.slot * 60 if a.slot is not None else None
    prev = None
    for i, s in enumerate(main_slides):
        n = i + 1
        if starts[i] is None:
            add(n, "timing", "notes do not start with a start time (m:ss.)")
            continue
        if prev is not None and starts[i] <= prev:
            add(n, "timing", f"starts at {starts[i] // 60}:{starts[i] % 60:02d}, not after the previous slide")
        prev = starts[i]
        if i == len(main_slides) - 1:
            if end is not None and starts[i] > end - 60:
                add(n, "timing", f"last slide starts at {starts[i] // 60}:{starts[i] % 60:02d}, later than the slot minus a minute")
            nxt = end
        else:
            nxt = starts[i + 1]
        if nxt is not None and nxt > starts[i]:
            words = len(START.sub("", notes_of(s), count=1).split())
            minutes = (nxt - starts[i]) / 60
            if words / 150 > 1.3 * minutes:
                add(n, "timing", f"{words} words take {words / 150:.1f} min to say, the slide has {minutes:.1f}")

    if a.budget is not None and len(main_slides) > a.budget:
        add(a.budget + 1, "budget", f"{len(main_slides)} main slides, budget {a.budget}")

    for n, check, detail in sorted(findings, key=lambda f: f[0]):
        print(f"slide {n}  {check}  {detail}")
    sys.exit(1 if findings else 0)


if __name__ == "__main__":
    main()
