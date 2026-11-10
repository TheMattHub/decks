"""Tests for pptx_check.py, on decks built here from fictional content (python-pptx)."""
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Inches

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "skills" / "making-presentations" / "scripts" / "pptx_check.py"
CASES = ROOT / "tests" / "leftover_cases.txt"


def check(deck, *args):
    return subprocess.run([sys.executable, str(SCRIPT), str(deck), *map(str, args)],
                          capture_output=True, text=True)


def findings(r, kind=None):
    lines = [ln for ln in r.stdout.splitlines() if ln.startswith("slide ")]
    return [ln for ln in lines if kind is None or f"  {kind}  " in ln]


def build(tmp_path, slides, name="d.pptx"):
    """slides: list of dicts with optional title, texts (list), notes."""
    prs = Presentation()
    for i, spec in enumerate(slides):
        s = prs.slides.add_slide(prs.slide_layouts[5])  # title only
        s.shapes.title.text = spec.get("title", "A title")
        for i, t in enumerate(spec.get("texts", [])):
            tb = s.shapes.add_textbox(Inches(1), Inches(2 + i), Inches(5), Inches(0.5))
            tb.text_frame.text = t
        # a start time by default, so tests of other checks do not trip the timing check
        s.notes_slide.notes_text_frame.text = spec.get("notes", f"{i}:00. ok")
    path = tmp_path / name
    prs.save(path)
    return path


def timed(tmp_path, notes, name="t.pptx"):
    return build(tmp_path, [{"notes": n} for n in notes], name)


def add_comments(path, part_name, reltype, ctype, body, slide=1):
    """Rewrite the zip: add a comments part and relate it to slide `slide`."""
    out = path.with_name("c_" + path.name)
    rels = f"ppt/slides/_rels/slide{slide}.xml.rels"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == rels:
                data = data.replace(b"</Relationships>", (
                    f'<Relationship Id="rIdC1" Type="{reltype}" Target="../comments/{part_name}"/>'
                    "</Relationships>").encode())
            if item.filename == "[Content_Types].xml":
                data = data.replace(b"</Types>", (
                    f'<Override PartName="/ppt/comments/{part_name}" ContentType="{ctype}"/></Types>').encode())
            zout.writestr(item, data)
        zout.writestr(f"ppt/comments/{part_name}", body)
    return out


# --- info lines and exit code -------------------------------------------------------------------

def test_clean_deck_exits_0_with_info_lines(tmp_path):
    deck = build(tmp_path, [{"title": "Cover"}, {"title": "Plan"}, {"title": "Backup"}, {"title": "Detail"}])
    r = check(deck)
    assert r.returncode == 0, r.stdout + r.stderr
    assert "slides: 2 main, 2 backup" in r.stdout
    assert "hidden: none" in r.stdout
    assert "embedded fonts: none" in r.stdout
    assert "author:" in r.stdout and "last modified by:" in r.stdout
    assert findings(r) == []


def test_hidden_slides_and_embedded_fonts_are_info_not_findings(tmp_path):
    deck = build(tmp_path, [{}, {}])
    prs = Presentation(deck)
    list(prs.slides)[1]._element.set("show", "0")
    prs.save(deck)
    out = deck.with_name("f.pptx")
    with zipfile.ZipFile(deck) as zin, zipfile.ZipFile(out, "w") as zout:
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        zout.writestr("ppt/fonts/font1.fntdata", b"x")
    r = check(out)
    assert r.returncode == 0
    assert "hidden: 2" in r.stdout
    assert "embedded fonts: ppt/fonts/font1.fntdata" in r.stdout


def test_a_finding_gives_exit_1(tmp_path):
    r = check(build(tmp_path, [{"texts": ["TODO later"]}]))
    assert r.returncode == 1


# --- leftover -----------------------------------------------------------------------------------

def parse_cases():
    out = []
    for line in CASES.read_text(encoding="utf-8").splitlines():
        line = line.strip("\r\n")
        if line.strip():
            kind, text = line.split("\t", 1)
            out.append((kind, text))
    return out


def test_every_leftover_case_agrees(tmp_path):
    cases = parse_cases()
    assert len(cases) > 10
    deck = build(tmp_path, [{"texts": [text for _, text in cases]}])
    r = check(deck)
    out = findings(r, "leftover")
    for kind, text in cases:
        needle = text[:60]
        hit = any(repr(needle) in ln or needle in ln for ln in out)
        assert hit == (kind == "hit"), f"{kind}: {text!r}\n{r.stdout}"


def test_leftover_in_group_table_and_notes(tmp_path):
    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    g = s.shapes.add_group_shape()
    inner = g.shapes.add_group_shape()
    inner.shapes.add_textbox(0, 0, Inches(2), Inches(1)).text_frame.text = "Lorem in a group"
    tbl = s.shapes.add_table(1, 2, 0, Inches(3), Inches(4), Inches(1)).table
    tbl.cell(0, 1).text = "TBD in a cell"
    s.notes_slide.notes_text_frame.text = "Click to add notes"
    path = tmp_path / "g.pptx"
    prs.save(path)
    r = check(path)
    out = findings(r, "leftover")
    assert any("Lorem in a group" in ln for ln in out)
    assert any("TBD in a cell" in ln for ln in out)
    assert any("notes:" in ln for ln in out)


def test_pattern_adds_and_exclude_drops(tmp_path):
    deck = build(tmp_path, [{"texts": ["Acme internal draft", "Option A for the north"]}])
    assert findings(check(deck, "--pattern", "internal"), "leftover")
    assert not [ln for ln in findings(check(deck)) if "internal" in ln]
    r = check(deck, "--exclude", "north")
    assert r.returncode == 0 and not findings(r)


# --- comment ------------------------------------------------------------------------------------

LEGACY = ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
          "application/vnd.openxmlformats-officedocument.presentationml.comments+xml",
          "comment1.xml",
          '<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
          '<p:cm authorId="0" dt="2026-01-01T00:00:00" idx="1"><p:pos x="1" y="1"/><p:text>a</p:text></p:cm>'
          '<p:cm authorId="0" dt="2026-01-01T00:00:00" idx="2"><p:pos x="1" y="1"/><p:text>b</p:text></p:cm>'
          "</p:cmLst>")
MODERN = ("http://schemas.microsoft.com/office/2018/10/relationships/comments",
          "application/vnd.ms-powerpoint.comments+xml",
          "modernComment_100_1.xml",
          '<p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main">'
          '<p188:cm id="{00000000-0000-0000-0000-000000000001}" authorId="{00000000-0000-0000-0000-000000000002}"/>'
          "</p188:cmLst>")


@pytest.mark.parametrize("kind,count", [(LEGACY, 2), (MODERN, 1)])
def test_comment_kinds(tmp_path, kind, count):
    reltype, ctype, part, body = kind
    deck = add_comments(build(tmp_path, [{}, {}]), part, reltype, ctype, body, slide=2)
    r = check(deck)
    assert r.returncode == 1
    out = findings(r, "comment")
    assert len(out) == 1 and out[0].startswith("slide 2  comment  ") and f"{count} review comment" in out[0]


# --- number -------------------------------------------------------------------------------------

def numbers_file(tmp_path, text="17\n312\n38%\n1,200\n"):
    p = tmp_path / "numbers.txt"
    p.write_text(text, encoding="utf-8")
    return p


def test_numbers_that_match_give_no_finding(tmp_path):
    deck = build(tmp_path, [{"texts": ["17 trucks", "312 min", "38 %", "1200 parcels", "1 200 parcels",
                                       "Step 3", "Q3 and N01 plans"]}])
    r = check(deck, "--numbers", numbers_file(tmp_path))
    assert findings(r, "number") == [], r.stdout


def test_unknown_number_and_year_are_findings(tmp_path):
    deck = build(tmp_path, [{"texts": ["42 depots", "2027", "17 trucks"]}])
    out = findings(check(deck, "--numbers", numbers_file(tmp_path)), "number")
    assert len(out) == 2
    assert any("42" in ln for ln in out) and any("2027" in ln for ln in out)


def test_decimal_comma_equals_dot_and_notes_are_not_checked(tmp_path):
    deck = build(tmp_path, [{"texts": ["3,5 kg"], "notes": "0:00. 99 in the notes"}])
    r = check(deck, "--numbers", numbers_file(tmp_path, "3.5\n"))
    assert findings(r, "number") == []


def test_numbers_in_table_cells_are_checked(tmp_path):
    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_table(1, 1, 0, 0, Inches(2), Inches(1)).table.cell(0, 0).text = "42 depots"
    path = tmp_path / "tb.pptx"
    prs.save(path)
    assert findings(check(path, "--numbers", numbers_file(tmp_path)), "number")


def test_slide_number_placeholder_is_skipped(tmp_path):
    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    tb = s.shapes.add_textbox(0, 0, Inches(1), Inches(1))
    tb.text_frame.text = "88"
    ph = tb._element.nvSpPr.nvPr
    from lxml import etree
    etree.SubElement(ph, "{http://schemas.openxmlformats.org/presentationml/2006/main}ph", type="sldNum")
    path = tmp_path / "sn.pptx"
    prs.save(path)
    assert findings(check(path, "--numbers", numbers_file(tmp_path)), "number") == []


# --- timing -------------------------------------------------------------------------------------

def words(n):
    return " ".join(["word"] * n)


def test_timing_ok_when_starts_increase(tmp_path):
    r = check(timed(tmp_path, ["0:00. a", "1:00. b", "2:30. c"]))
    assert findings(r, "timing") == []


def test_timing_missing_and_not_increasing(tmp_path):
    out = findings(check(timed(tmp_path, ["0:00. a", "no time", "1:00. c", "1:00. d"])), "timing")
    assert len(out) == 2
    assert out[0].startswith("slide 2") and out[1].startswith("slide 4")


def test_timing_slot(tmp_path):
    ok = check(timed(tmp_path, ["0:00. a", "9:00. b"]), "--slot", 10)
    late = check(timed(tmp_path, ["0:00. a", "9:30. b"], "late.pptx"), "--slot", 10)
    assert findings(ok, "timing") == []
    out = findings(late, "timing")
    assert len(out) == 1 and out[0].startswith("slide 2")


def test_timing_words_per_minute(tmp_path):
    fine = timed(tmp_path, ["0:00. " + words(150), "1:00. b"], "fine.pptx")
    long = timed(tmp_path, ["0:00. " + words(300), "1:00. b"], "long.pptx")
    assert findings(check(fine), "timing") == []
    out = findings(check(long), "timing")
    assert len(out) == 1 and out[0].startswith("slide 1")


def test_timing_ignores_backup_slides(tmp_path):
    deck = build(tmp_path, [{"notes": "0:00. a"}, {"notes": "1:00. b"}, {"title": "Backup"}, {"title": "More"}])
    assert findings(check(deck), "timing") == []


# --- budget -------------------------------------------------------------------------------------

def test_budget(tmp_path):
    deck = build(tmp_path, [{} for _ in range(9)])
    out = findings(check(deck, "--budget", 8), "budget")
    assert len(out) == 1 and "9 main" in out[0]
    assert findings(check(deck, "--budget", 9), "budget") == []


def test_backup_slides_do_not_count_and_backup_title_overrides(tmp_path):
    deck = build(tmp_path, [{}, {}, {"title": "Reserve"}, {}, {}])
    assert findings(check(deck, "--backup-title", "^reserve", "--budget", 2), "budget") == []
    assert len(findings(check(deck, "--backup-title", "^reserve", "--budget", 1), "budget")) == 1
    r = check(deck, "--backup-title", "^reserve")
    assert "slides: 2 main, 3 backup" in r.stdout


def test_grouped_numbers_with_a_decimal_tail_are_compared_by_value(tmp_path):
    deck = build(tmp_path, [{"texts": ["1,234.56 EUR", "1.234,5 kg"]}])
    r = check(deck, "--numbers", numbers_file(tmp_path, "1234.56\n1234.5\n"))
    assert r.returncode == 0 and "Traceback" not in r.stderr, r.stderr
    assert findings(r, "number") == []
    assert len(findings(check(deck, "--numbers", numbers_file(tmp_path, "17\n")), "number")) == 2


def test_numbers_file_may_contain_a_grouped_decimal(tmp_path):
    deck = build(tmp_path, [{"texts": ["1234.56 EUR"]}])
    r = check(deck, "--numbers", numbers_file(tmp_path, "1,234.56\n"))
    assert r.returncode == 0 and findings(r, "number") == [], r.stderr


def test_space_grouped_number_passes_when_every_part_is_listed(tmp_path):
    deck = build(tmp_path, [{"texts": ["17 312 trucks"]}])
    assert findings(check(deck, "--numbers", numbers_file(tmp_path)), "number") == []
    deck = build(tmp_path, [{"texts": ["42 312 trucks"]}], "b.pptx")
    out = findings(check(deck, "--numbers", numbers_file(tmp_path)), "number")
    assert len(out) == 1 and "42 312" in out[0]
