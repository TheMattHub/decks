"""Tests for the Python scripts, on decks built here from fictional content (python-pptx)."""
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt

SCRIPTS = Path(__file__).resolve().parents[1] / "skills" / "making-presentations" / "scripts"
XSL = Path(r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL")


def run(script, *args):
    return subprocess.run([sys.executable, str(SCRIPTS / script), *map(str, args)],
                          capture_output=True, text=True)


def deck_with_big_picture(tmp_path):
    img = tmp_path / "big.png"
    Image.effect_noise((4000, 2250), 40).convert("RGB").save(img)
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_picture(str(img), 0, 0, width=prs.slide_width)
    path = tmp_path / "in.pptx"
    prs.save(path)
    return path


def test_media_report_lists_media(tmp_path):
    deck = deck_with_big_picture(tmp_path)
    r = run("pptx_media.py", deck)
    assert r.returncode == 0, r.stderr
    assert "image1.png" in r.stdout


def test_media_copy_is_smaller_and_keeps_names(tmp_path):
    deck = deck_with_big_picture(tmp_path)
    out = tmp_path / "out.pptx"
    r = run("pptx_media.py", deck, out, "--skip-video")
    assert r.returncode == 0, r.stderr
    with zipfile.ZipFile(deck) as a, zipfile.ZipFile(out) as b:
        assert a.namelist() == b.namelist()
        assert b.getinfo("ppt/media/image1.png").file_size < a.getinfo("ppt/media/image1.png").file_size
    Presentation(out)  # still a valid deck


def test_media_goal_email_caps_pictures(tmp_path):
    deck = deck_with_big_picture(tmp_path)
    out = tmp_path / "out.pptx"
    r = run("pptx_media.py", deck, out, "--skip-video", "--goal", "email")
    assert r.returncode == 0, r.stderr
    with zipfile.ZipFile(out) as z, z.open("ppt/media/image1.png") as f:
        assert max(Image.open(f).size) == 1280


def test_media_refuses_to_overwrite_input(tmp_path):
    deck = deck_with_big_picture(tmp_path)
    assert run("pptx_media.py", deck, deck).returncode != 0


@pytest.mark.skipif(not XSL.exists(), reason="needs Microsoft Office's MML2OMML.XSL")
def test_equations_become_native(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
    box.text_frame.text = r"$$\sum_{i=1}^{n} y_{ik} = 1$$"
    box.text_frame.paragraphs[0].runs[0].font.size = Pt(28)
    deck, out = tmp_path / "eq.pptx", tmp_path / "eq-out.pptx"
    prs.save(deck)
    r = run("pptx_equations.py", deck, out)
    assert r.returncode == 0, r.stderr
    with zipfile.ZipFile(out) as z:
        xml = z.read("ppt/slides/slide1.xml").decode("utf-8")
    assert "a14:m" in xml and "oMath" in xml
