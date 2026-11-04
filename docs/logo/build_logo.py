# /// script
# requires-python = ">=3.10"
# dependencies = ["fonttools"]
# ///
"""Draws the decks logo: three slides fanned from nearly one corner, the drafts v6 and final_v7
underneath and the last version on top with a tick.

    uv run docs/logo/build_logo.py            # writes the SVGs and PNGs next to this script

Writes the mark (light and dark), icons (512 to 16 px; 16 px keeps two slides), an avatar on the
paper background (also the plugin listing icon, .claude-plugin/icon.png), the README banners and the
1280 x 640 social preview. Text is converted to outlines with IBM Plex
Sans (SIL Open Font License), so the files look the same everywhere. PNGs are rendered with
Microsoft Edge or Google Chrome in headless mode.
"""
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from fontTools.pens.svgPathPen import SVGPathPen
from fontTools.pens.transformPen import TransformPen
from fontTools.ttLib import TTFont

HERE = Path(__file__).resolve().parent

INK, PAPER, CORAL = "#1B1F2B", "#F5F1EA", "#E2583E"
THEMES = {  # back slide (v6), its label, middle slide (final_v7)
    "light": {"v6": "#B3A897", "v6_label": "#5F574C", "v7": "#2F7D7A"},
    "dark": {"v6": "#5E6782", "v6_label": PAPER, "v7": "#3FA39E"},
}
# One slide box; each slide turns around its own bottom-left corner, spaced slightly apart.
X, Y, W, H = 26, 50, 80, 52
BACK = "translate(4 4) rotate(12 26 102)"
MID = "rotate(0 26 102)"
TOP = "translate(-5 -1) rotate(-15 26 102)"
VIEW = (3.5, 16.6, 119.5, 119.5)  # the whole fan, square, so icons are never cut


def find_font(style):
    """IBM Plex Sans from the user or system font folder, or IBM_PLEX_DIR."""
    name = f"IBMPlexSans-{style}.ttf"
    dirs = [os.environ.get("IBM_PLEX_DIR", ""),
            Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft/Windows/Fonts",
            Path("C:/Windows/Fonts"), Path.home() / "Library/Fonts", Path("/usr/share/fonts/truetype/ibm-plex")]
    for d in dirs:
        p = Path(d) / name
        if d and p.exists():
            return TTFont(p)
    raise SystemExit(f"{name} not found: install IBM Plex Sans or set IBM_PLEX_DIR")


FONTS = {}


def text_path(text, x, y, size, fill, style="Bold", anchor="start"):
    """Text as one SVG path (outlines), baseline at y."""
    font = FONTS.setdefault(style, find_font(style))
    gs, cmap = font.getGlyphSet(), font.getBestCmap()
    scale = size / font["head"].unitsPerEm
    width = sum(gs[cmap[ord(c)]].width for c in text) * scale
    x0 = x - width if anchor == "end" else x - width / 2 if anchor == "middle" else x
    pen, cx = SVGPathPen(gs), 0
    for c in text:
        g = gs[cmap[ord(c)]]
        g.draw(TransformPen(pen, (scale, 0, 0, -scale, x0 + cx * scale, y)))
        cx += g.width
    return f'<path d="{pen.getCommands()}" fill="{fill}"/>'


def label(text, fill, size):
    return text_path(text, X + W - 6, Y + H - 5, size, fill, anchor="end")


def mark(theme, labels=True, tick=5.5):
    """The three slides, in the 0..260 drawing coordinates."""
    t = THEMES[theme]
    return (f'<g transform="{BACK}"><rect x="{X}" y="{Y}" width="{W}" height="{H}" rx="5" fill="{t["v6"]}"/>'
            + (label("v6", t["v6_label"], 9) if labels else "") + "</g>"
            + f'<g transform="{MID}"><rect x="{X}" y="{Y}" width="{W}" height="{H}" rx="5" fill="{t["v7"]}"/>'
            + (label("final_v7", PAPER, 8.5) if labels else "") + "</g>"
            + f'<g transform="{TOP}"><rect x="{X}" y="{Y}" width="{W}" height="{H}" rx="5" fill="{CORAL}"/>'
            f'<path d="M{X+24} {Y+27} l9 9 l19 -20" stroke="{PAPER}" stroke-width="{tick}" fill="none" '
            'stroke-linecap="round" stroke-linejoin="round"/></g>')


def svg(body, size, view=VIEW, background=None):
    bg = f'<rect x="{view[0]}" y="{view[1]}" width="{view[2]}" height="{view[3]}" fill="{background}"/>' if background else ""
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{size[0]}" height="{size[1]}" '
            f'viewBox="{" ".join(f"{v:g}" for v in view)}">{bg}{body}</svg>\n')


def icon16(theme):
    """Two slides only: three merge into one blob at 16 px."""
    return svg(f'<g transform="rotate(10 8 50)"><rect x="10" y="22" width="44" height="30" rx="4" fill="{THEMES[theme]["v7"]}"/></g>'
               f'<g transform="rotate(-12 8 50)"><rect x="8" y="16" width="44" height="30" rx="4" fill="{CORAL}"/>'
               f'<path d="M20 31 l6 6 l12 -13" stroke="{PAPER}" stroke-width="5" fill="none" stroke-linecap="round" stroke-linejoin="round"/></g>',
               (16, 16), (2, 6, 60, 60))


def social():
    """1280 x 640: the mark, the name and the line, on paper."""
    k = 3.6  # mark scale
    body = (f'<rect width="1280" height="640" fill="{PAPER}"/>'
            f'<g transform="translate(150 95) scale({k}) translate({-VIEW[0]} {-VIEW[1]})">{mark("light")}</g>'
            + text_path("decks", 640, 330, 150, INK)
            + text_path("the last version", 646, 400, 44, BANNER["light"]["line"], "Regular")
            + text_path("From a short brief to a PowerPoint you can present.", 646, 466, 24, BANNER["light"]["body"], "Regular"))
    return svg(body, (1280, 640), (0, 0, 1280, 640))


BANNER = {  # panel, title, punchline, subtitle, and the faint outlines of past drafts in the background
    # line: the punchline in the logo's coral, deepened on light and lifted on dark (4.8:1 and 6.0:1);
    # body: the subtitle in a calm neutral (6.9:1 and 9.5:1).
    "light": {"panel": PAPER, "title": INK, "line": "#B9432C", "body": "#57524A", "ghost": "#D9D1C4"},
    "dark": {"panel": INK, "title": PAPER, "line": "#EE7A62", "body": "#C9C4B8", "ghost": "#2E3446"},
}


def banner(theme):
    """1280 x 360 README banner: the mark, the name and the line, with faint discarded drafts behind."""
    b = BANNER[theme]
    ghosts = ""
    for x, y, a, s in ((880, 36, -14, 3.0), (1010, 165, 9, 2.6), (1150, -40, -6, 2.3), (985, 300, 5, 2.0)):
        ghosts += (f'<g transform="translate({x} {y}) rotate({a}) scale({s})"><rect width="80" height="52" rx="6" fill="{b["panel"]}" '
                   f'stroke="{b["ghost"]}" stroke-width="{2.2 / s * 2:.2f}"/>'
                   f'<path d="M10 14h40 M10 22h52 M10 30h34 M10 38h46" stroke="{b["ghost"]}" stroke-width="{2.2 / s * 2:.2f}" stroke-linecap="round"/></g>')
    k = 2.45
    body = (f'<defs><clipPath id="panel"><rect width="1280" height="360" rx="28"/></clipPath></defs>'
            f'<g clip-path="url(#panel)"><rect width="1280" height="360" fill="{b["panel"]}"/>{ghosts}</g>'
            f'<g transform="translate(70 34) scale({k}) translate({-VIEW[0]} {-VIEW[1]})">{mark(theme)}</g>'
            + text_path("decks", 420, 190, 118, b["title"])
            + text_path("the last version", 425, 248, 36, b["line"], "Regular")
            + text_path("From a short brief to a PowerPoint you can present.", 425, 296, 22, b["body"], "Regular"))
    return svg(body, (1280, 360), (0, 0, 1280, 360))


def browser():
    for c in ("msedge", "chrome", "google-chrome", "chromium"):
        if shutil.which(c):
            return shutil.which(c)
    for p in (r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
              r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
              r"C:\Program Files\Google\Chrome\Application\chrome.exe"):
        if Path(p).exists():
            return p
    raise SystemExit("PNG rendering needs Microsoft Edge or Google Chrome")


def png(svg_text, out, w, h):
    """Renders an SVG to a PNG of exactly w x h with a transparent background."""
    with tempfile.TemporaryDirectory() as tmp:
        page = Path(tmp) / "page.html"
        page.write_text(f'<html><body style="margin:0;background:transparent">{svg_text}</body></html>', encoding="utf-8")
        subprocess.run([browser(), "--headless=new", "--disable-gpu", "--hide-scrollbars",
                        "--default-background-color=00000000", f"--window-size={w},{h}",
                        f"--screenshot={out}", page.as_uri()], check=True, capture_output=True)


def main():
    for theme in THEMES:
        (HERE / f"decks-mark-{theme}.svg").write_text(svg(mark(theme), (160, 160)), encoding="utf-8")
    small = svg(mark("light", labels=False, tick=7), (512, 512))
    (HERE / "decks-icon.svg").write_text(small, encoding="utf-8")
    for s in (512, 256, 64, 32):
        png(svg(mark("light", labels=s >= 256, tick=5.5 if s >= 256 else 7), (s, s)), HERE / f"decks-icon-{s}.png", s, s)
    png(icon16("light"), HERE / "decks-icon-16.png", 16, 16)
    pad = 14  # avatar: the mark on paper, with a margin
    view = (VIEW[0] - pad, VIEW[1] - pad, VIEW[2] + 2 * pad, VIEW[3] + 2 * pad)
    png(svg(mark("light"), (512, 512), view, PAPER), HERE / "decks-avatar-512.png", 512, 512)
    # The plugin directory's listing icon: a square PNG at .claude-plugin/icon.png (512-2048 px, no SVG).
    png(svg(mark("light"), (512, 512), view, PAPER), HERE.parents[1] / ".claude-plugin" / "icon.png", 512, 512)
    for theme in BANNER:
        b = banner(theme)
        (HERE / f"decks-banner-{theme}.svg").write_text(b, encoding="utf-8")
        png(b, HERE / f"decks-banner-{theme}.png", 1280, 360)
    s = social()
    (HERE / "decks-social-preview.svg").write_text(s, encoding="utf-8")
    png(s, HERE / "decks-social-preview.png", 1280, 640)
    print("logo files in", HERE)


if __name__ == "__main__":
    main()
