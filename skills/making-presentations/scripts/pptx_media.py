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

Every part keeps its name and format, so slides, crops, trims and play settings are untouched. A file
is replaced only when the new version is smaller. Pictures wider or taller than the cap are scaled
down; JPEGs are re-saved at the quality; PNGs are re-saved optimised. Videos are re-encoded to H.264
with the same length, so PowerPoint's trims still hold. ffmpeg comes from imageio-ffmpeg, or from PATH.
"""
import argparse
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
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


def shrink_picture(data, ext, max_px, quality):
    img = Image.open(io.BytesIO(data))
    # Without the profile a wide-gamut photo shifts colour; resize() and convert() drop it, so take it first.
    icc = img.info.get("icc_profile")
    if max(img.size) > max_px:
        k = max_px / max(img.size)
        img = img.resize((round(img.width * k), round(img.height * k)), Image.LANCZOS)
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


def shrink_video(data, ext, crf, ffmpeg, max_px):
    with tempfile.TemporaryDirectory() as tmp:
        src, dst = Path(tmp) / ("in" + ext), Path(tmp) / ("out" + ext)
        src.write_bytes(data)
        scale = []
        size = video_size(ffmpeg, src)
        if size and max(size) > max_px:
            k = max_px / max(size)
            # libx264 with yuv420p needs even dimensions
            scale = ["-vf", "scale=" + ":".join(str(max(2, round(d * k / 2) * 2)) for d in size)]
        cmd = [ffmpeg, "-y", "-loglevel", "error", "-i", str(src), *scale, "-c:v", "libx264", "-crf", str(crf),
               "-preset", "slow", "-pix_fmt", "yuv420p", "-c:a", "aac", "-b:a", "128k",
               "-movflags", "+faststart", str(dst)]
        subprocess.run(cmd, check=True)
        return dst.read_bytes()


def mb(n):
    return f"{n / 1e6:7.2f} MB"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input")
    ap.add_argument("output", nargs="?")
    ap.add_argument("--goal", choices=GOALS, default="present")
    ap.add_argument("--max-px", type=int)
    ap.add_argument("--max-video-px", type=int)
    ap.add_argument("--jpeg-quality", type=int)
    ap.add_argument("--crf", type=int, help="video quality: higher is smaller, 28 for small cards")
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
                                new = shrink_picture(data, ext, max_px, quality)
                            elif ext in VIDEOS and ffmpeg:
                                new = shrink_video(data, ext, crf, ffmpeg, video_px)
                        except Exception as e:  # one odd file must not stop the rest
                            print(f"  kept {item.filename[10:]}: {e}")
                    if new is not None and len(new) < len(data):
                        print(f"  {item.filename[10:]}: {mb(len(data))} -> {mb(len(new))}")
                        saved += len(data) - len(new)
                        data = new
                    # Media are already compressed: store them, deflate the XML.
                    kind = zipfile.ZIP_STORED if item.filename.startswith("ppt/media/") else zipfile.ZIP_DEFLATED
                    out.writestr(item, data, compress_type=kind)
        except BaseException:
            dst.unlink(missing_ok=True)
            raise
        size = dst.stat().st_size
        print(f"{dst.name}: {mb(size)} (media {mb(saved)} smaller), goal {a.goal}")
        if limit and size > limit:
            print(f"still over {limit / 1e6:.0f} MB for {a.goal}: cut or shorten the largest files above, "
                  "move long films to a link, or try --goal email")
            sys.exit(1)


if __name__ == "__main__":
    main()
