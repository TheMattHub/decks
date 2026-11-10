"""Tests for the Python scripts, on decks built here from fictional content (python-pptx)."""
import re
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest
import imageio_ffmpeg
from PIL import Image, ImageCms
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


def make_video(path, w, h):
    subprocess.run([imageio_ffmpeg.get_ffmpeg_exe(), "-y", "-loglevel", "error", "-f", "lavfi", "-i",
                    f"testsrc=size={w}x{h}:rate=10:duration=1", "-pix_fmt", "yuv420p", str(path)], check=True)


def video_size(path):
    info = subprocess.run([imageio_ffmpeg.get_ffmpeg_exe(), "-i", str(path)], capture_output=True, text=True).stderr
    return tuple(map(int, re.search(r"Video:.*?, (\d+)x(\d+)", info).groups()))


def test_media_video_is_capped_per_goal(tmp_path):
    video = tmp_path / "v.mp4"
    make_video(video, 1280, 720)
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_movie(str(video), 0, 0, prs.slide_width, prs.slide_height, mime_type="video/mp4")
    deck, out = tmp_path / "in.pptx", tmp_path / "out.pptx"
    prs.save(deck)
    r = run("pptx_media.py", deck, out, "--goal", "email")
    assert r.returncode == 0, r.stderr
    with zipfile.ZipFile(out) as z:
        names = [n for n in z.namelist() if n.endswith(".mp4")]
        z.extract(names[0], tmp_path)
    assert video_size(tmp_path / names[0]) == (960, 540)


def test_media_exit_code_one_when_copy_is_over_the_limit(tmp_path):
    deck = deck_with_big_picture(tmp_path)  # noise: the 4000 px PNG stays over 15 MB
    out = tmp_path / "out.pptx"
    r = run("pptx_media.py", deck, out, "--skip-video", "--goal", "email", "--max-px", 4000)
    assert out.exists() and r.returncode == 1 and "still over" in r.stdout


def test_media_keeps_icc_profile(tmp_path):
    icc = ImageCms.ImageCmsProfile(ImageCms.createProfile("sRGB")).tobytes()
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for name in ("p.png", "p.jpg"):
        Image.effect_noise((400, 300), 40).convert("RGB").save(tmp_path / name, icc_profile=icc)
        slide.shapes.add_picture(str(tmp_path / name), 0, 0, width=prs.slide_width // 4)
    deck, out = tmp_path / "in.pptx", tmp_path / "out.pptx"
    prs.save(deck)
    assert run("pptx_media.py", deck, out, "--skip-video", "--max-px", 200).returncode == 0
    with zipfile.ZipFile(out) as z:
        for part in ("ppt/media/image1.png", "ppt/media/image2.jpg"):
            with z.open(part) as f:
                assert Image.open(f).info.get("icc_profile") == icc, part


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
