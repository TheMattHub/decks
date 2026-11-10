# /// script
# requires-python = ">=3.10"
# dependencies = ["pillow", "imageio-ffmpeg"]
# ///
"""Report and shrink the media inside a .pptx, so a deck stays small enough to mail and share.

    uv run pptx_media.py IN.pptx                       # report only: size per media file, largest first
    uv run pptx_media.py IN.pptx OUT.pptx              # compressed copy for presenting; IN is never changed
    uv run pptx_media.py IN.pptx OUT.pptx --goal email # smaller copy to attach to a mail

--goal sets picture size, video size, JPEG quality and video quality for what the deck is for, and the
size the copy should stay under:
    present  2560 px, video 1920, JPEG 85, CRF 24, no limit   full-screen slides stay sharp on a projector or 4K
    share    1920 px, video 1280, JPEG 80, CRF 28, 25 MB      Teams, OneDrive, a download link; viewed on a laptop
    email    1280 px, video  960, JPEG 75, CRF 30, 15 MB      attachment; mail encoding adds about a third, so 15 MB
                                                              stays under the 20 MB limit of Outlook and most servers
--max-px, --max-video-px, --jpeg-quality and --crf override the goal's value. The pixel caps apply to the
longest side and never upscale; the colour profile of a re-saved picture is kept.

A picture or video needs only the pixels it is shown at: by default each file is also scaled to what its
slides display, measured as if the cap were the pixel count across the slide's longer side, with crops
(the stored file must still cover the uncropped image), groups and shared files taken into account, and
never upscaled. A file used where the size is unknown (layout, master, background, fill) keeps the plain
cap. --no-footprint turns this off.

The exit code is 1 when the copy was written but is still over the goal's size limit. --fit writes it again
from the original, up to 4 times, each step with both pixel caps x0.8, JPEG quality -5 (not under 50) and
CRF +2 (not over 34), until it is under the limit; the settings of the copy that was kept are printed.

Every part keeps its name and format, so slides, crops, trims and play settings are untouched. A file
is replaced only when the new version is smaller. Pictures wider or taller than the cap are scaled
down; JPEGs are re-saved at the quality; PNGs are re-saved optimised. Videos are re-encoded to H.264
with the same length, so PowerPoint's trims still hold. ffmpeg comes from imageio-ffmpeg, or from PATH.
"""
import argparse
import copy
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from posixpath import dirname, join, normpath
from pathlib import Path

from PIL import Image

PICTURES = {".png", ".jpg", ".jpeg"}
VIDEOS = {".mp4", ".m4v", ".mov"}
# goal: (max picture px, max video px, JPEG quality, video CRF, size limit in bytes or None); the docstring says why.
GOALS = {
    "present": (2560, 1920, 85, 24, None),
    "share": (1920, 1280, 80, 28, 25e6),
    "email": (1280, 960, 75, 30, 15e6),
}


def ffmpeg_exe():
    try:
        import imageio_ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except Exception:
        return shutil.which("ffmpeg")


NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
R = "{%s}" % NS["r"]


def media_targets(z, rels_name):
    """rId -> media part name, for the relationships of one part that point into ppt/media."""
    base = dirname(dirname(rels_name))  # the rels file sits in <part folder>/_rels/
    found = {}
    for rel in ET.fromstring(z.read(rels_name)).iter("{%s}Relationship" % NS["rel"]):
        if rel.get("TargetMode") == "External":
            continue
        t = rel.get("Target")
        t = t.lstrip("/") if t.startswith("/") else normpath(join(base, t))
        if t.startswith("ppt/media/"):
            found[rel.get("Id")] = t
    return found


def footprints(z):
    """media part -> (w, h): the most any slide shows it at, as fractions of the slide's longer side and
    of the uncropped image. A part missing from the result is used somewhere its size is unknown."""
    sld = ET.fromstring(z.read("ppt/presentation.xml")).find("p:sldSz", NS)
    longer = max(int(sld.get("cx")), int(sld.get("cy")))
    need, plain, accounted = {}, set(), set()

    def walk(el, sx, sy, slide, rids):
        for c in el:
            if c.tag == "{%s}pic" % NS["p"]:
                x, srcrect = c.find("p:spPr/a:xfrm", NS), c.find("p:blipFill/a:srcRect", NS)
                ext = x.find("a:ext", NS) if x is not None else None
                if ext is None or not int(ext.get("cx")) or not int(ext.get("cy")):
                    continue  # no own geometry (a placeholder): its media stay unaccounted, so plain
                w, h = int(ext.get("cx")) * sx, int(ext.get("cy")) * sy
                if 45 < int(x.get("rot", 0)) / 60000 % 180 < 135:
                    w, h = h, w  # turned a quarter: the image's width runs along the slide's height
                crop = {k: int(srcrect.get(k, 0)) / 100000 if srcrect is not None else 0 for k in "ltrb"}
                kw, kh = 1 - crop["l"] - crop["r"], 1 - crop["t"] - crop["b"]
                if kw <= 0 or kh <= 0:
                    continue
                refs = [e.get(R + "embed") for e in c.iterfind(".//a:blip", NS)]
                refs += [e.get(R + "link") for e in c.iterfind(".//a:videoFile", NS)]
                refs += [e.get(R + k) for e in c.iterfind(".//p14:media", NS) for k in ("embed", "link")]
                for rid in filter(None, refs):
                    if rid in rids:
                        old = need.get(rids[rid], (0, 0))
                        need[rids[rid]] = (max(old[0], w / longer / kw), max(old[1], h / longer / kh))
                        accounted.add((slide, rid))
                continue
            g = c.find("p:grpSpPr/a:xfrm", NS) if c.tag == "{%s}grpSp" % NS["p"] else None
            if g is not None and int(g.find("a:chExt", NS).get("cx")) and int(g.find("a:chExt", NS).get("cy")):
                e, ce = g.find("a:ext", NS), g.find("a:chExt", NS)
                walk(c, sx * int(e.get("cx")) / int(ce.get("cx")), sy * int(e.get("cy")) / int(ce.get("cy")),
                     slide, rids)
            else:
                walk(c, sx, sy, slide, rids)

    for name in z.namelist():
        if not name.endswith(".rels"):
            continue
        if re.fullmatch(r"ppt/slides/_rels/slide\d+\.xml\.rels", name):
            rids = media_targets(z, name)
            walk(ET.fromstring(z.read("ppt/slides/" + name.rsplit("/", 1)[1][:-5])), 1, 1, name, rids)
            plain.update(t for rid, t in rids.items() if (name, rid) not in accounted)
        else:
            plain.update(media_targets(z, name).values())
    return {k: v for k, v in need.items() if k not in plain}


def scale_factor(w, h, cap, foot):
    """Never up, never past the cap; with a footprint, only as large as the slide shows it."""
    k = min(1, cap / max(w, h))
    return min(k, max(foot[0] * cap / w, foot[1] * cap / h)) if foot else k


def shrink_picture(data, ext, max_px, quality, foot=None):
    img = Image.open(io.BytesIO(data))
    # Without the profile a wide-gamut photo shifts colour; resize() and convert() drop it, so take it first.
    icc = img.info.get("icc_profile")
    k = scale_factor(*img.size, max_px, foot)
    if k < 1:
        img = img.resize((max(1, round(img.width * k)), max(1, round(img.height * k))), Image.LANCZOS)
    out = io.BytesIO()
    if ext == ".png":
        img.save(out, "PNG", optimize=True, icc_profile=icc)
    else:
        img.convert("RGB").save(out, "JPEG", quality=quality, optimize=True, progressive=True, icc_profile=icc)
    return out.getvalue()


def video_size(ffmpeg, src):
    """Width and height from ffmpeg's stream listing (no ffprobe ships with imageio-ffmpeg)."""
    info = subprocess.run([ffmpeg, "-i", str(src)], capture_output=True, text=True).stderr
    m = re.search(r"Video:.*?, (\d{2,5})x(\d{2,5})", info)
    return (int(m[1]), int(m[2])) if m else None


def shrink_video(data, ext, crf, ffmpeg, max_px, foot=None):
    with tempfile.TemporaryDirectory() as tmp:
        src, dst = Path(tmp) / ("in" + ext), Path(tmp) / ("out" + ext)
        src.write_bytes(data)
        scale = []
        size = video_size(ffmpeg, src)
        k = scale_factor(*size, max_px, foot) if size else 1
        if k < 1:
            # libx264 with yuv420p needs even dimensions
            scale = ["-vf", "scale=" + ":".join(str(max(2, round(d * k / 2) * 2)) for d in size)]
        cmd = [ffmpeg, "-y", "-loglevel", "error", "-i", str(src), *scale, "-c:v", "libx264", "-crf", str(crf),
               "-preset", "slow", "-pix_fmt", "yuv420p", "-c:a", "aac", "-b:a", "128k",
               "-movflags", "+faststart", str(dst)]
        subprocess.run(cmd, check=True)
        return dst.read_bytes()


def mb(n):
    return f"{n / 1e6:7.2f} MB"


def write_copy(z, dst, ffmpeg, feet, max_px, video_px, quality, crf):
    """Write the shrunk copy of the open deck to dst; returns its size and the media bytes saved."""
    saved = 0
    try:
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as out:
            for item in z.infolist():
                data = z.read(item)
                ext = os.path.splitext(item.filename)[1].lower()
                new = None
                if item.filename.startswith("ppt/media/"):
                    try:
                        if ext in PICTURES:
                            new = shrink_picture(data, ext, max_px, quality, feet.get(item.filename))
                        elif ext in VIDEOS and ffmpeg:
                            new = shrink_video(data, ext, crf, ffmpeg, video_px, feet.get(item.filename))
                    except Exception as e:  # one odd file must not stop the rest
                        print(f"  kept {item.filename[10:]}: {e}")
                if new is not None and len(new) < len(data):
                    print(f"  {item.filename[10:]}: {mb(len(data))} -> {mb(len(new))}")
                    saved += len(data) - len(new)
                    data = new
                # Media are already compressed: store them, deflate the XML.
                kind = zipfile.ZIP_STORED if item.filename.startswith("ppt/media/") else zipfile.ZIP_DEFLATED
                out.writestr(copy.copy(item), data, compress_type=kind)  # writestr edits the ZipInfo; z is read again per step
    except BaseException:
        dst.unlink(missing_ok=True)
        raise
    return dst.stat().st_size, saved


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input")
    ap.add_argument("output", nargs="?")
    ap.add_argument("--goal", choices=GOALS, default="present")
    ap.add_argument("--max-px", type=int)
    ap.add_argument("--max-video-px", type=int)
    ap.add_argument("--jpeg-quality", type=int)
    ap.add_argument("--crf", type=int, help="video quality: higher is smaller, 28 for small cards")
    ap.add_argument("--no-footprint", action="store_true", help="cap by pixels only, not by displayed size")
    ap.add_argument("--fit", action="store_true", help="when over the goal's limit, retry with smaller settings")
    ap.add_argument("--skip-video", action="store_true")
    a = ap.parse_args()
    max_px, video_px, quality, crf, limit = GOALS[a.goal]
    max_px, video_px = a.max_px or max_px, a.max_video_px or video_px
    quality, crf = a.jpeg_quality or quality, a.crf or crf

    src = Path(a.input).resolve()
    with zipfile.ZipFile(src) as z:
        media = sorted((i for i in z.infolist() if i.filename.startswith("ppt/media/")),
                       key=lambda i: -i.file_size)
        print(f"{src.name}: {mb(src.stat().st_size)} on disk, media {mb(sum(i.file_size for i in media))}")
        for i in media[:15]:
            # Over 5 MB is worth a look: a raw recording, a huge photo, or a film that belongs in the backup.
            flag = "  <- large: --crf 28, a shorter or smaller clip, or a smaller picture" if i.file_size > 5e6 else ""
            print(f"  {mb(i.file_size)}  {i.filename[10:]}{flag}")
        if not a.output:
            return

        dst = Path(a.output).resolve()
        if dst == src:
            sys.exit("OUT must differ from IN: the input is never changed")
        ffmpeg = None if a.skip_video else ffmpeg_exe()
        if not a.skip_video and not ffmpeg:
            print("no ffmpeg found: videos are left as they are")
        feet = {} if a.no_footprint else footprints(z)
        settings = (max_px, video_px, quality, crf)
        size, saved = write_copy(z, dst, ffmpeg, feet, *settings)
        if a.fit and not limit:
            print(f"--fit: goal {a.goal} has no size limit, nothing to fit")
        # Each step starts again from the original media: re-compressing a compressed copy only adds loss.
        for step in range(1, 5 if a.fit and limit else 1):
            if size <= limit:
                break
            settings = (round(max_px * 0.8 ** step), round(video_px * 0.8 ** step),
                        min(quality, max(50, quality - 5 * step)), max(crf, min(34, crf + 2 * step)))
            print(f"over {limit / 1e6:.0f} MB, --fit step {step}: {settings[0]} px, video {settings[1]} px, "
                  f"JPEG {settings[2]}, CRF {settings[3]}")
            size, saved = write_copy(z, dst, ffmpeg, feet, *settings)
        if a.fit and limit:
            print("kept: {} px, video {} px, JPEG {}, CRF {}".format(*settings))
        print(f"{dst.name}: {mb(size)} (media {mb(saved)} smaller), goal {a.goal}")
        if limit and size > limit:
            print(f"still over {limit / 1e6:.0f} MB for {a.goal}: cut or shorten the largest files above, "
                  "move long films to a link, or try --goal email")
            sys.exit(1)


if __name__ == "__main__":
    main()
