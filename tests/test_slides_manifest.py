"""Tests for slides_manifest.py, on fictional decks written in the Claude Slides file format."""
import json
import subprocess
import sys
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[1] / "skills" / "making-presentations" / "scripts" / "slides_manifest.py"

SLIDES = {
    "cover": '<section id="cover" data-transition="fade" style="background:#111;width:1920px;height:1080px">'
             '<h1 style="font-size:96px">Fish &amp; <i>Chips</i> <span>Q3</span></h1>'
             '<p>Sub</p><aside>Notes: say <h1>hello</h1> data-build-in</aside></section>',
    "plan": '<section id="plan" data-transition="push" style="background:#fff">'
            '<p style="font-size:30px">Intro line before the heading</p>'
            '<h2>Plan</h2>'
            '<div data-build-in="rise" style="position:absolute"><div><b>Costs</b> fell\n   <i>12%</i></div></div>'
            '<img data-build-in="pop" src="/_blob/x" alt="">'
            '<p data-build-in="fade">' + "x" * 80 + '</p>'
            '<video src="/_blob/v1" poster="/_blob/p1" style="position:absolute;left:128px;top:200px;width:960px;height:540px"></video>'
            '<video style="width:100px"><source src="/_blob/v2"></video>'
            '<video src="/_blob/v3" style="position:absolute;right:100px;bottom:80px;width:400px;height:300px"></video>'
            '<aside>Notes</aside></section>',
    "end": '<section id="end" data-transition="wobble" style="background:#fff">'
           '<div data-build-in="spin">Thanks</div><p>Bye</p></section>',
}


def make_deck(tmp_path, order, slides=SLIDES):
    (tmp_path / "project" / "slides").mkdir(parents=True)
    (tmp_path / "project" / "deck.json").write_text(json.dumps({"v": 4, "title": "Q3", "order": order}))
    for sid, html in slides.items():
        (tmp_path / "project" / "slides" / f"{sid}.html").write_text(html, encoding="utf-8")
    return tmp_path


def run(*args):
    return subprocess.run([sys.executable, str(SCRIPT), *map(str, args)], capture_output=True, text=True, encoding="utf-8")


def manifest(tmp_path, order=("cover", "plan", "end")):
    r = run(make_deck(tmp_path, list(order)))
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout), r.stderr


def test_order_comes_from_deck_json_not_file_order(tmp_path):
    m, _ = manifest(tmp_path, ("end", "cover", "plan"))
    assert [(s["index"], s["id"]) for s in m["slides"]] == [(1, "end"), (2, "cover"), (3, "plan")]
    assert m["deck"] == "Q3" and (m["canvas_width"], m["canvas_height"]) == (1920, 1080)


def test_title_resolves_entities_and_nested_tags_and_skips_notes(tmp_path):
    m, _ = manifest(tmp_path)
    assert m["slides"][0]["title"] == "Fish & Chips Q3"
    # The first heading wins over earlier text; no heading falls back to the first text.
    assert m["slides"][1]["title"] == "Plan"
    assert m["slides"][2]["title"] == "Thanks"


def test_transitions_and_unknown_names_warned_but_kept(tmp_path):
    m, err = manifest(tmp_path)
    assert [s["transition"] for s in m["slides"]] == ["fade", "push", "wobble"]
    assert "unknown transition 'wobble'" in err and "unknown build-in 'spin'" in err
    assert m["slides"][2]["builds"] == [{"effect": "spin", "text": "Thanks"}]


def test_builds_nested_textless_and_trimmed_in_document_order(tmp_path):
    b = manifest(tmp_path)[0]["slides"][1]["builds"]
    assert [x["effect"] for x in b] == ["rise", "pop", "fade"]
    assert b[0]["text"] == "Costs fell 12%"
    assert b[1]["text"] == ""
    assert b[2]["text"] == "x" * 60


def test_notes_never_leak_into_builds_or_titles(tmp_path):
    m, _ = manifest(tmp_path)
    assert "Notes" not in json.dumps(m) and "hello" not in json.dumps(m)


def test_videos_with_and_without_box(tmp_path):
    v = manifest(tmp_path)[0]["slides"][1]["videos"]
    assert v[0] == {"src": "/_blob/v1", "poster": "/_blob/p1", "box": {"x": 128, "y": 200, "w": 960, "h": 540}}
    assert v[1] == {"src": "/_blob/v2", "poster": None, "box": None}
    assert v[2]["box"] == {"x": 1420, "y": 700, "w": 400, "h": 300}


def test_writes_out_file(tmp_path):
    deck = make_deck(tmp_path / "d", ["cover"], {"cover": SLIDES["cover"]})
    out = tmp_path / "m.json"
    r = run(deck, out)
    assert r.returncode == 0 and r.stdout == ""
    assert json.loads(out.read_text(encoding="utf-8"))["slides"][0]["id"] == "cover"


def fails(r, text):
    assert r.returncode == 1 and text in r.stderr and r.stdout == ""
    assert len(r.stderr.strip().splitlines()) == 1


def test_fails_without_deck_json(tmp_path):
    fails(run(tmp_path), "no project/deck.json")


def test_fails_on_missing_order_and_bad_json(tmp_path):
    (tmp_path / "project").mkdir()
    (tmp_path / "project" / "deck.json").write_text('{"title": "x"}')
    fails(run(tmp_path), '"order"')
    (tmp_path / "project" / "deck.json").write_text("{nope")
    fails(run(tmp_path), "not valid JSON")


def test_fails_on_listed_slide_file_missing(tmp_path):
    deck = make_deck(tmp_path, ["cover", "gone"], {"cover": SLIDES["cover"]})
    fails(run(deck), "'gone' is listed in deck.json but project/slides/gone.html is missing")


def test_fails_on_slide_without_section(tmp_path):
    deck = make_deck(tmp_path, ["cover"], {"cover": "<div><h1>Hi</h1></div>"})
    fails(run(deck), "slide cover: no <section>")


def test_slide_file_missing_from_order_is_left_out_with_warning(tmp_path):
    # The export has only the slides in "order", so the manifest must not count more.
    deck = make_deck(tmp_path, ["cover"], {"cover": SLIDES["cover"], "plan": SLIDES["plan"]})
    r = run(deck)
    assert r.returncode == 0 and "left out" in r.stderr and "plan" in r.stderr
    assert [s["id"] for s in json.loads(r.stdout)["slides"]] == ["cover"]


def test_nested_section_does_not_end_the_slide(tmp_path):
    html = '<section id=s><section>A</section><h2 data-build-in=rise>B</h2><video src="/v"></video></section>'
    deck = make_deck(tmp_path, ["s"], {"s": html})
    slide = json.loads(run(deck).stdout)["slides"][0]
    assert slide["builds"] == [{"effect": "rise", "text": "B"}]
    assert [v["src"] for v in slide["videos"]] == ["/v"]


def test_style_and_script_text_never_becomes_title_or_build_text(tmp_path):
    html = ('<section id=s><style>.a { color: red }</style><script>var x = 1;</script>'
            '<div data-build-in=fade><style>.b{}</style>Shown</div></section>')
    slide = json.loads(run(make_deck(tmp_path, ["s"], {"s": html})).stdout)["slides"][0]
    assert slide["title"] == "Shown"
    assert slide["builds"] == [{"effect": "fade", "text": "Shown"}]


def test_build_text_is_the_first_text_block(tmp_path):
    # The export puts each block in its own box, so the whole card's text would match none of them.
    html = ('<section id=s><h2>Cards</h2><div data-build-in=rise><h3>Card one</h3><p>What card one says.</p></div>'
            '<span data-build-in=pop>Inline <b>only</b></span></section>')
    slide = json.loads(run(make_deck(tmp_path, ["s"], {"s": html})).stdout)["slides"][0]
    assert [b["text"] for b in slide["builds"]] == ["Card one", "Inline only"]


def test_whitespace_in_title_and_build_text_is_collapsed(tmp_path):
    html = '<section id=s><h1>  Two\n   lines </h1><p data-build-in=pop>a \t b\n c</p></section>'
    slide = json.loads(run(make_deck(tmp_path, ["s"], {"s": html})).stdout)["slides"][0]
    assert slide["title"] == "Two lines" and slide["builds"][0]["text"] == "a b c"
