# /// script
# requires-python = ">=3.10"
# dependencies = ["pillow", "imageio-ffmpeg"]
# ///
"""Report and shrink the media inside a .pptx, so a deck stays small enough to mail and share.

    uv run pptx_media.py IN.pptx                 # report only: size per media file, largest first
    uv run pptx_media.py IN.pptx OUT.pptx        # write a compressed copy; IN is never changed

Every part keeps its name and format, so slides, crops, trims and play settings are untouched. A file
is replaced only when the new version is smaller. Pictures wider or taller than --max-px are scaled
down (2560 px keeps full-screen slides sharp); JPEGs are re-saved at --jpeg-quality; PNGs are
re-saved optimised. Videos are re-encoded to H.264 at --crf with the same length, so PowerPoint's
trims still hold. ffmpeg comes from imageio-ffmpeg, or from PATH.
"""
import argparse
import io
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from PIL import Image

PICTURES = {".png", ".jpg", ".jpeg"}
VIDEOS = {".mp4", ".m4v", ".mov"}


def ffmpeg_exe():
    try:
        import imageio_ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except Exception:
        return shutil.which("ffmpeg")


def shrink_picture(data, ext, max_px, quality):
    img = Image.open(io.BytesIO(data))
    if max(img.size) > max_px:
        k = max_px / max(img.size)
        img = img.resize((round(img.width * k), round(img.height * k)), Image.LANCZOS)
    out = io.BytesIO()
    if ext == ".png":
        img.save(out, "PNG", optimize=True)
    else:
        img.convert("RGB").save(out, "JPEG", quality=quality, optimize=True, progressive=True)
    return out.getvalue()


def shrink_video(data, ext, crf, ffmpeg):
    with tempfile.TemporaryDirectory() as tmp:
        src, dst = Path(tmp) / ("in" + ext), Path(tmp) / ("out" + ext)
        src.write_bytes(data)
        cmd = [ffmpeg, "-y", "-loglevel", "error", "-i", str(src), "-c:v", "libx264", "-crf", str(crf),
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
    ap.add_argument("--max-px", type=int, default=2560)
    ap.add_argument("--jpeg-quality", type=int, default=85)
    ap.add_argument("--crf", type=int, default=24, help="video quality: 24 default, 28 for small cards")
    ap.add_argument("--skip-video", action="store_true")
    a = ap.parse_args()

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
                                new = shrink_picture(data, ext, a.max_px, a.jpeg_quality)
                            elif ext in VIDEOS and ffmpeg:
                                new = shrink_video(data, ext, a.crf, ffmpeg)
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
        print(f"{dst.name}: {mb(dst.stat().st_size)} (media {mb(saved)} smaller)")


if __name__ == "__main__":
    main()
