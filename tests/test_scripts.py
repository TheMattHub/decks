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


# --- pptx_media footprint sizing: a 16:9 deck, goal share (picture cap 1920 px across the longer side) ---
QUARTER = 3048000  # EMU: a quarter of the slide width


def footprint_deck(tmp_path, size=(4000, 2250)):
    prs = Presentation()
    prs.slide_width, prs.slide_height = 12192000, 6858000
    img = tmp_path / "pic.png"
    Image.new("RGB", size, (200, 80, 40)).save(img)
    return prs, str(img)


def save(prs, tmp_path):
    deck = tmp_path / "in.pptx"
    prs.save(deck)
    return deck


def media_size(prs, tmp_path, *args):
    deck, out = save(prs, tmp_path), tmp_path / "out.pptx"
    r = run("pptx_media.py", deck, out, "--skip-video", "--goal", "share", *args)
    assert r.returncode == 0, r.stderr
    with zipfile.ZipFile(out) as z, z.open("ppt/media/image1.png") as f:
        return Image.open(f).size


def near(size, expected):
    return all(abs(a - b) <= 1 for a, b in zip(size, expected))


def test_footprint_full_width(tmp_path):
    prs, img = footprint_deck(tmp_path)
    prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_picture(img, 0, 0, width=prs.slide_width)
    assert near(media_size(prs, tmp_path), (1920, 1080))


def test_footprint_quarter_width(tmp_path):
    prs, img = footprint_deck(tmp_path)
    prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_picture(img, 0, 0, width=QUARTER)
    assert near(media_size(prs, tmp_path), (480, 270))


def test_footprint_crop_needs_more_pixels(tmp_path):
    prs, img = footprint_deck(tmp_path)
    pic = prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_picture(img, 0, 0, width=QUARTER)
    pic.crop_left = pic.crop_right = 0.25
    assert near(media_size(prs, tmp_path), (960, 540))


def test_footprint_shared_file_takes_the_largest_need(tmp_path):
    prs, img = footprint_deck(tmp_path)
    prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_picture(img, 0, 0, width=QUARTER)
    prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_picture(img, 0, 0, width=2 * QUARTER)
    assert near(media_size(prs, tmp_path), (960, 540))


def test_footprint_group_scale(tmp_path):
    prs, img = footprint_deck(tmp_path)
    group = prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_group_shape()
    group.shapes.add_picture(img, 0, 0, width=QUARTER)
    xfrm = group._element.grpSpPr.xfrm  # drawn twice as large as its children's own space
    xfrm.ext.cx, xfrm.ext.cy = 2 * xfrm.chExt.cx, 2 * xfrm.chExt.cy
    assert near(media_size(prs, tmp_path), (960, 540))


def test_footprint_small_picture_unchanged(tmp_path):
    prs, img = footprint_deck(tmp_path, (300, 169))
    prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_picture(img, 0, 0, width=QUARTER)
    assert media_size(prs, tmp_path) == (300, 169)


def test_footprint_off_keeps_the_plain_cap(tmp_path):
    prs, img = footprint_deck(tmp_path)
    prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_picture(img, 0, 0, width=QUARTER)
    assert near(media_size(prs, tmp_path, "--no-footprint"), (1920, 1080))


def test_footprint_video(tmp_path):
    video = tmp_path / "v.mp4"
    make_video(video, 1280, 720)
    prs, _ = footprint_deck(tmp_path)
    prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_movie(
        str(video), 0, 0, QUARTER, QUARTER * 9 // 16, mime_type="video/mp4")
    deck, out = save(prs, tmp_path), tmp_path / "out.pptx"
    assert run("pptx_media.py", deck, out, "--goal", "share").returncode == 0
    with zipfile.ZipFile(out) as z:
        name = next(n for n in z.namelist() if n.endswith(".mp4"))
        z.extract(name, tmp_path)
    assert near(video_size(tmp_path / name), (320, 180))


def test_footprint_media_also_on_a_layout_keeps_the_plain_cap(tmp_path):
    prs, img = footprint_deck(tmp_path)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    pic = slide.shapes.add_picture(img, 0, 0, width=QUARTER)
    # the same media on the layout too: its size there is unknown
    slide.slide_layout.part.relate_to(slide.part.related_part(pic._element.blip_rId),
                                      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
    assert near(media_size(prs, tmp_path), (1920, 1080))


def test_media_fit_steps_down_until_under_the_limit(tmp_path):
    deck = deck_with_big_picture(tmp_path)  # noise: 4000 px is ~21 MB, 3200 px ~10 MB
    out = tmp_path / "out.pptx"
    r = run("pptx_media.py", deck, out, "--skip-video", "--goal", "email", "--max-px", 4000, "--fit")
    assert r.returncode == 0, r.stdout + r.stderr
    assert "kept: 3200 px" in r.stdout
    with zipfile.ZipFile(out) as z, z.open("ppt/media/image1.png") as f:
        assert max(Image.open(f).size) == 3200


def test_media_fit_exits_1_when_the_last_step_is_still_over(tmp_path):
    deck = deck_with_big_picture(tmp_path)
    r = run("pptx_media.py", deck, tmp_path / "out.pptx", "--skip-video", "--goal", "email", "--max-px", 12000, "--fit")
    assert r.returncode == 1 and "--fit step 4" in r.stdout


def test_footprint_rotation_does_not_change_the_need(tmp_path):
    # a:ext is the unrotated box, so a quarter-turn leaves the need as it is: need 240 x 480 -> k 0.24
    prs, img = footprint_deck(tmp_path, (4000, 2000))
    pic = prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_picture(img, 0, 0, 1524000, 3048000)
    pic.rotation = 90
    assert near(media_size(prs, tmp_path), (960, 480))


def rewrite_zip(deck, change):
    """Replace parts of a deck: change(name, data) returns the new bytes."""
    out = deck.with_name("rewritten.pptx")
    with zipfile.ZipFile(deck) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, change(item.filename, zin.read(item)))
    return out


def run_media(deck, tmp_path, *args):
    out = tmp_path / "out.pptx"
    r = run("pptx_media.py", deck, out, "--skip-video", "--goal", "share", *args)
    assert r.returncode == 0, r.stdout + r.stderr
    with zipfile.ZipFile(out) as z, z.open("ppt/media/image1.png") as f:
        return Image.open(f).size


def test_footprint_photo_also_used_as_background_keeps_the_plain_cap(tmp_path):
    prs, img = footprint_deck(tmp_path)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    pic = slide.shapes.add_picture(img, 0, 0, width=QUARTER)
    rid = pic._element.blip_rId  # one rId per target: the background shares it with the small picture
    bg = ('<p:bg xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
          f'<p:bgPr><a:blipFill><a:blip r:embed="{rid}"/></a:blipFill><a:effectLst/></p:bgPr></p:bg>')
    from lxml import etree
    slide._element.cSld.insert(0, etree.fromstring(bg))
    assert near(media_size(prs, tmp_path), (1920, 1080))


def test_footprint_placeholder_picture_without_geometry_keeps_the_plain_cap(tmp_path):
    import copy
    prs, img = footprint_deck(tmp_path)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    pic = slide.shapes.add_picture(img, 0, 0, width=QUARTER)
    twin = copy.deepcopy(pic._element)
    twin.spPr.remove(twin.spPr.xfrm)
    twin.nvPicPr.cNvPr.set("id", "99")
    pic._element.addnext(twin)
    assert near(media_size(prs, tmp_path), (1920, 1080))


def test_footprint_group_without_chext_keeps_the_plain_cap_and_writes_the_copy(tmp_path):
    prs, img = footprint_deck(tmp_path)
    group = prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_group_shape()
    group.shapes.add_picture(img, 0, 0, width=QUARTER)
    xfrm = group._element.grpSpPr.xfrm
    xfrm.remove(xfrm.chExt)
    assert near(media_size(prs, tmp_path), (1920, 1080))


def test_footprint_slide_that_does_not_parse_keeps_the_plain_cap(tmp_path):
    prs, img = footprint_deck(tmp_path)
    prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_picture(img, 0, 0, width=QUARTER)
    deck = rewrite_zip(save(prs, tmp_path),
                       lambda n, d: b"<p:sld><unclosed" if n == "ppt/slides/slide1.xml" else d)
    assert near(run_media(deck, tmp_path), (1920, 1080))


def test_footprint_relationship_without_target_keeps_the_plain_cap(tmp_path):
    prs, img = footprint_deck(tmp_path)
    prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_picture(img, 0, 0, width=QUARTER)
    deck = rewrite_zip(save(prs, tmp_path), lambda n, d: d.replace(
        b"</Relationships>", b'<Relationship Id="rId77" Type="x"/></Relationships>')
        if n == "ppt/slides/_rels/slide1.xml.rels" else d)
    assert near(run_media(deck, tmp_path), (1920, 1080))


def test_media_cmyk_jpeg_is_kept_as_it_is(tmp_path):
    # converting to RGB while keeping the CMYK profile would be invalid colour: leave the file alone
    prs = Presentation()
    img = tmp_path / "c.jpg"
    Image.effect_noise((3000, 2000), 40).convert("CMYK").save(img, quality=95)
    prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_picture(str(img), 0, 0, width=prs.slide_width)
    deck, out = tmp_path / "in.pptx", tmp_path / "out.pptx"
    prs.save(deck)
    assert run("pptx_media.py", deck, out, "--skip-video", "--goal", "email").returncode == 0
    with zipfile.ZipFile(out) as z:
        assert z.read("ppt/media/image1.jpg") == img.read_bytes()


def test_media_greyscale_jpeg_stays_greyscale_with_its_profile(tmp_path):
    icc = ImageCms.ImageCmsProfile(ImageCms.createProfile("sRGB")).tobytes()  # stands in for a grey profile
    prs = Presentation()
    img = tmp_path / "g.jpg"
    Image.effect_noise((3000, 2000), 40).convert("L").save(img, quality=95, icc_profile=icc)
    prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_picture(str(img), 0, 0, width=prs.slide_width)
    deck, out = tmp_path / "in.pptx", tmp_path / "out.pptx"
    prs.save(deck)
    assert run("pptx_media.py", deck, out, "--skip-video", "--goal", "email").returncode == 0
    with zipfile.ZipFile(out) as z, z.open("ppt/media/image1.jpg") as f:
        got = Image.open(f)
        assert got.mode == "L" and got.info.get("icc_profile") == icc and got.width < 3000


def test_media_video_with_rotation_metadata_is_not_squashed(tmp_path):
    flat, turned = tmp_path / "flat.mp4", tmp_path / "turned.mp4"
    make_video(flat, 640, 360)
    subprocess.run([imageio_ffmpeg.get_ffmpeg_exe(), "-y", "-loglevel", "error", "-display_rotation", "90",
                    "-i", str(flat), "-c", "copy", str(turned)], check=True)
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_movie(
        str(turned), 0, 0, prs.slide_width, prs.slide_height, mime_type="video/mp4")
    deck, out = tmp_path / "in.pptx", tmp_path / "out.pptx"
    prs.save(deck)
    assert run("pptx_media.py", deck, out, "--max-video-px", 320, "--no-footprint").returncode == 0
    with zipfile.ZipFile(out) as z:
        name = next(n for n in z.namelist() if n.endswith(".mp4"))
        z.extract(name, tmp_path)
    assert near(video_size(tmp_path / name), (180, 320))  # the frames as ffmpeg turns them, aspect kept


def test_media_fit_without_a_limit_does_nothing(tmp_path):
    deck = deck_with_big_picture(tmp_path)
    r = run("pptx_media.py", deck, tmp_path / "out.pptx", "--skip-video", "--fit")
    assert r.returncode == 0 and "nothing to fit" in r.stdout and "step" not in r.stdout