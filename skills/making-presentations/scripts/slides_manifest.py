# /// script
# requires-python = ">=3.10"
# dependencies = []
# ///
"""Read what the PowerPoint export of a Claude Slides deck drops (transitions, build-ins, videos) from the
deck's own files into a manifest, which Invoke-FinishManifest applies to the exported .pptx.

    uv run slides_manifest.py DECK_DIR              # manifest as JSON on stdout
    uv run slides_manifest.py DECK_DIR OUT.json     # manifest written to OUT.json

DECK_DIR holds the deck's files as fetched: project/deck.json and project/slides/<id>.html. The slides are
taken in deck.json's "order". A slide file that "order" leaves out is appended last (the deck shows it last)
and reported on stderr. Exit 1 with one line on stderr when the structure is not recognised: no deck.json, no
"order" list, a listed slide file missing, a slide file with no <section>. A transition or build name other
than fade|push|magic / rise|fade|pop is kept verbatim and warned about on stderr.

Manifest (flat; every key is always present):

    {
      "deck": "Q3 review",               # deck.json "title"
      "canvas_width": 1920,              # the section's width/height in px, else 1920 x 1080
      "canvas_height": 1080,
      "slides": [                        # deck order
        {
          "index": 1,                    # 1-based position in the exported deck
          "id": "cover",                 # slide id = file name = section id
          "title": "Q3 review",          # first non-empty heading (h1-h6), else the first block of text outside the
                                         # notes; whitespace collapsed, tags and entities resolved. "" if none
          "transition": "fade",          # fade | push | magic | null (other names kept verbatim)
          "builds": [                    # elements with data-build-in, in document order
            {"effect": "rise",           # rise | fade | pop (other names kept verbatim)
             "text": "Costs fell 12%"}   # the element's text, first 60 characters; "" for no text
          ],
          "videos": [                    # <video> elements, in document order
            {"src": "/_blob/ab12",       # the src attribute, else the first <source src>; "" if none
             "poster": null,             # the poster attribute or null
             "box": {"x": 128, "y": 200, "w": 960, "h": 540}}   # px on the canvas, or null
          ]
        }
      ]
    }

A video's box needs width and height in its inline style, and a position: left and top, or right and bottom
(then x = canvas_width - right - w). Without them the box is null: the video sits in the layout flow and the
file does not say where. Speaker notes (the <aside>) never reach titles or build texts.
"""
import json
import re
import sys
from html.parser import HTMLParser
from pathlib import Path

TRANSITIONS = {"fade", "push", "magic"}
BUILDS = {"rise", "fade", "pop"}
VOID = {"br", "img", "hr", "source", "input", "meta", "link", "col", "wbr"}
INLINE = {"b", "i", "u", "a", "span"}
HEADINGS = {"h1", "h2", "h3", "h4", "h5", "h6"}
SLIDE_ID = re.compile(r"[A-Za-z0-9_-]{1,64}")
TEXT_MAX = 60


class Fail(Exception):
    pass


def px(style, prop):
    m = re.search(r"(?:^|;)\s*" + prop + r"\s*:\s*(-?[\d.]+)px\s*(?:;|$)", style or "", re.I)
    return float(m.group(1)) if m else None


def num(v):
    return int(v) if v == int(v) else v


def collapse(parts):
    return " ".join("".join(parts).split())


class SlideParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.section = None       # attrs of the first <section>
        self.stack = []           # open elements: [tag, parts]
        self.in_section = False
        self.section_done = False
        self.aside_depth = 0
        self.builds = []          # [effect, text parts]
        self.videos = []          # {"src","poster","style"}
        self.headings = []
        self.first_text = None
        self.open_video = None

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "section" and not self.section_done and not self.in_section:
            self.section, self.in_section = a, True
            return
        if not self.in_section:
            return
        if tag == "aside":
            self.aside_depth += 1
        if tag == "br":
            self.add_text(" ")
        if tag == "source" and self.open_video is not None and not self.open_video["src"]:
            self.open_video["src"] = a.get("src") or ""
        parts = []
        if a.get("data-build-in") is not None:
            self.builds.append([a["data-build-in"], parts])  # shares the frame's text list
        if tag == "video":
            self.open_video = {"src": a.get("src") or "", "poster": a.get("poster") or None,
                               "style": a.get("style")}
            self.videos.append(self.open_video)
        if tag not in VOID:
            self.stack.append([tag, parts])

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        if tag not in VOID:
            self.handle_endtag(tag)

    def handle_endtag(self, tag):
        if not self.in_section:
            return
        if tag == "section":
            self.in_section, self.section_done = False, True
            return
        # Tolerant of unclosed tags: close back to the nearest matching open element.
        for i in range(len(self.stack) - 1, -1, -1):
            if self.stack[i][0] == tag:
                for t, parts in reversed(self.stack[i:]):
                    self.close_tag(t, parts)
                    if t == "aside":
                        self.aside_depth -= 1
                    if t == "video":
                        self.open_video = None
                del self.stack[i:]
                return

    def close_tag(self, tag, parts):
        if tag in HEADINGS:
            self.headings.append(collapse(parts))
        elif tag not in INLINE and self.first_text is None and collapse(parts):
            self.first_text = collapse(parts)  # first block with text, not the whole slide's

    def add_text(self, data):
        if self.aside_depth:
            return
        for _, parts in self.stack:
            parts.append(data)

    def handle_data(self, data):
        if self.in_section:
            self.add_text(data)


def slide_entry(index, sid, html, warnings):
    p = SlideParser()
    p.feed(html)
    p.close()
    for t, parts in reversed(p.stack):  # unclosed at end of file
        p.close_tag(t, parts)
    if p.section is None:
        raise Fail(f"slide {sid}: no <section> in its file")
    s = p.section
    transition = s.get("data-transition") or None
    if transition and transition not in TRANSITIONS:
        warnings.append(f"slide {sid}: unknown transition '{transition}' (kept)")
    builds = []
    for effect, parts in p.builds:
        if effect not in BUILDS:
            warnings.append(f"slide {sid}: unknown build-in '{effect}' (kept)")
        builds.append({"effect": effect, "text": collapse(parts)[:TEXT_MAX].rstrip()})
    title = next((h for h in p.headings if h), None)
    if title is None:
        title = p.first_text or ""
    videos = [{"src": v["src"], "poster": v["poster"], "box": box(v["style"], s)} for v in p.videos]
    return s, {"index": index, "id": sid, "title": title, "transition": transition,
               "builds": builds, "videos": videos}


def box(style, section_attrs):
    w, h = px(style, "width"), px(style, "height")
    if w is None or h is None:
        return None
    cw, ch = canvas(section_attrs)
    left, top = px(style, "left"), px(style, "top")
    if left is None:
        right = px(style, "right")
        left = None if right is None else cw - right - w
    if top is None:
        bottom = px(style, "bottom")
        top = None if bottom is None else ch - bottom - h
    if left is None or top is None:
        return None
    return {"x": num(left), "y": num(top), "w": num(w), "h": num(h)}


def canvas(section_attrs):
    style = section_attrs.get("style")
    return (num(px(style, "width") or 1920), num(px(style, "height") or 1080))


def build_manifest(deck_dir):
    root = Path(deck_dir)
    deck_json = root / "project" / "deck.json"
    if not deck_json.is_file():
        raise Fail(f"no project/deck.json in {root}")
    try:
        deck = json.loads(deck_json.read_text(encoding="utf-8-sig"))
    except ValueError as e:
        raise Fail(f"project/deck.json is not valid JSON: {e}")
    order = deck.get("order") if isinstance(deck, dict) else None
    if not isinstance(order, list) or not order or not all(isinstance(i, str) for i in order):
        raise Fail('project/deck.json has no "order" list of slide ids')
    slides_dir = root / "project" / "slides"
    warnings = []
    for sid in order:
        if not SLIDE_ID.fullmatch(sid):
            raise Fail(f"slide id '{sid}' in deck.json is not [A-Za-z0-9_-]{{1,64}}")
    extra = sorted(f.stem for f in slides_dir.glob("*.html") if f.stem not in order) if slides_dir.is_dir() else []
    if extra:
        warnings.append("slide files not in deck.json order, appended last: " + ", ".join(extra))
    entries, cw, ch = [], 1920, 1080
    for n, sid in enumerate(order + extra, 1):
        f = slides_dir / f"{sid}.html"
        if not f.is_file():
            raise Fail(f"slide '{sid}' is listed in deck.json but project/slides/{sid}.html is missing")
        s, entry = slide_entry(n, sid, f.read_text(encoding="utf-8-sig"), warnings)
        if n == 1:
            cw, ch = canvas(s)
        entries.append(entry)
    manifest = {"deck": deck.get("title") or "", "canvas_width": cw, "canvas_height": ch, "slides": entries}
    return manifest, warnings


def main(argv):
    if not 1 <= len(argv) <= 2 or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0 if argv and argv[0] in ("-h", "--help") else 1
    try:
        manifest, warnings = build_manifest(argv[0])
    except Fail as e:
        print(f"slides_manifest: {e}", file=sys.stderr)
        return 1
    for w in warnings:
        print(f"warning: {w}", file=sys.stderr)
    text = json.dumps(manifest, indent=2, ensure_ascii=False)
    if len(argv) == 2:
        Path(argv[1]).write_text(text + "\n", encoding="utf-8")
    else:
        sys.stdout.reconfigure(encoding="utf-8")
        print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
