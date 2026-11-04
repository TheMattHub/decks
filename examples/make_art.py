# /// script
# requires-python = ">=3.10"
# dependencies = ["pillow"]
# ///
"""Generates the illustrations for the marketing demo deck: layered landscapes in four palettes.
Original artwork made here, so the README images carry no third-party licence.

    uv run examples/make_art.py OUT_DIR
"""
import math
import random
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

W, H = 1920, 1280

PALETTES = {  # sky top, sky bottom, sun, ridge colours from far to near
    "dawn": ("#F6C9A8", "#FBE9D7", "#FFF4E2", ["#E3A487", "#C97B66", "#8F4F4A", "#4E2E33"]),
    "forest": ("#BFD9D2", "#EEF2E6", "#FFFDF4", ["#8DB3A2", "#5E8C78", "#356353", "#1C3B33"]),
    "lake": ("#9EC5E0", "#E3EEF3", "#FFFFFF", ["#7FA4C0", "#567F9F", "#2F5878", "#16324A"]),
    "night": ("#141B33", "#3A3F6B", "#F3E6C4", ["#2E3460", "#232850", "#181C3C", "#0D1026"]),
}


def hex_rgb(h):
    return tuple(int(h[i:i + 2], 16) for i in (1, 3, 5))


def lerp(a, b, t):
    return tuple(round(x + (y - x) * t) for x, y in zip(a, b))


def ridge(rng, base, amp, rough):
    """Points of a mountain ridge: a few sine waves plus noise, across the width."""
    phases = [rng.uniform(0, math.tau) for _ in range(3)]
    pts = []
    for x in range(0, W + 20, 20):
        y = base
        for k, ph in enumerate(phases, 1):
            y -= amp / k * math.sin(x / W * math.tau * k * 0.9 + ph)
        y += rng.uniform(-rough, rough)
        pts.append((x, y))
    return pts + [(W, H), (0, H)]


def scene(name, seed):
    top, bottom, sun, ridges = PALETTES[name]
    rng = random.Random(seed)
    img = Image.new("RGB", (W, H))
    d = ImageDraw.Draw(img)
    a, b = hex_rgb(top), hex_rgb(bottom)
    for y in range(H):
        d.line([(0, y), (W, y)], fill=lerp(a, b, min(1, y / (H * 0.75))))
    if name == "night":  # stars
        for _ in range(260):
            x, y = rng.randrange(W), rng.randrange(int(H * 0.6))
            r = rng.choice([1, 1, 1, 2])
            d.ellipse([x, y, x + r, y + r], fill=(255, 250, 235))
    glow = Image.new("RGB", (W, H), (0, 0, 0))
    gx, gy = rng.randint(int(W * 0.55), int(W * 0.8)), rng.randint(int(H * 0.18), int(H * 0.32))
    ImageDraw.Draw(glow).ellipse([gx - 260, gy - 260, gx + 260, gy + 260], fill=hex_rgb(sun))
    strength = 0.12 if name == "night" else 0.35  # a strong halo turns grey on a dark sky
    img = Image.blend(img, Image.composite(glow, img, glow.convert("L")), strength).filter(ImageFilter.SMOOTH)
    d = ImageDraw.Draw(img)
    d.ellipse([gx - 90, gy - 90, gx + 90, gy + 90], fill=hex_rgb(sun))
    for i, c in enumerate(ridges):
        base = H * (0.48 + 0.13 * i)
        d.polygon(ridge(rng, base, 120 - 18 * i, 6 + 4 * i), fill=hex_rgb(c))
    if name == "lake":  # a calm water band with reflections
        wy = int(H * 0.8)
        d.rectangle([0, wy, W, H], fill=hex_rgb("#6E97B5"))
        for _ in range(70):
            x, y = rng.randrange(W), rng.randrange(wy + 8, H)
            d.line([(x, y), (x + rng.randint(40, 160), y)], fill=(208, 226, 236), width=2)
    return img


def main():
    out = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    out.mkdir(parents=True, exist_ok=True)
    for i, name in enumerate(PALETTES):
        scene(name, seed=7 + i).save(out / f"{name}.jpg", quality=88)
    print(f"4 illustrations in {out}")


if __name__ == "__main__":
    main()
