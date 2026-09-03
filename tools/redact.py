"""Blur sensitive regions of extracted frames before they go into a manual.

Regions are declared per source frame in a JSON spec so the same redaction
can be re-applied if frames are re-extracted:

    {"grid_00010.jpg": {"out": "02-payroll-notes.png",
                        "blur": [[596, 386, 1270, 660]],
                        "crop": [0, 0, 1113, 676]}}

Boxes are [left, top, right, bottom] in pixels against the extracted frame.
The optional "crop" trims the frame after blurring, which is how the webcam
column on a call recording is removed rather than blurred.
"""
import argparse
import json
from pathlib import Path

from PIL import Image, ImageFilter

BLUR_RADIUS = 14


def redact(src, boxes, crop=None):
    image = Image.open(src).convert("RGB")
    for left, top, right, bottom in boxes:
        left, top = max(0, left), max(0, top)
        right = min(image.width, right)
        bottom = min(image.height, bottom)
        if right <= left or bottom <= top:
            raise SystemExit(f"{src.name}: empty box {[left, top, right, bottom]}")
        patch = image.crop((left, top, right, bottom))
        # Scrub detail before blurring: a Gaussian blur alone can leave
        # large text faintly legible.
        patch = patch.resize((max(1, patch.width // 12), max(1, patch.height // 12)))
        patch = patch.resize((right - left, bottom - top), Image.NEAREST)
        image.paste(patch.filter(ImageFilter.GaussianBlur(BLUR_RADIUS)), (left, top))
    if crop:
        image = image.crop(tuple(crop))
    return image


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("--frames", required=True)
    ap.add_argument("--outdir", required=True)
    args = ap.parse_args()

    spec = json.loads(Path(args.spec).read_text())
    frames = Path(args.frames)
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    for name, entry in spec.items():
        src = frames / name
        if not src.exists():
            raise SystemExit(f"missing frame: {src}")
        image = redact(src, entry.get("blur", []), entry.get("crop"))
        dest = outdir / entry["out"]
        image.save(dest)
        print(f"{name} -> {dest.name}  ({len(entry.get('blur', []))} regions)")


if __name__ == "__main__":
    main()
