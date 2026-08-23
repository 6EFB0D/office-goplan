#!/usr/bin/env python3
"""Create gui-multiselect.png — paint multi-select on page strip of gui-drawings.png."""

from __future__ import annotations

from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont

SRC = Path(r"d:\Users\admin_mak\project\office-goplan\assets\pdfhandler\gui-drawings.png")
OUT = Path(r"d:\Users\admin_mak\project\office-goplan\assets\pdfhandler\gui-multiselect.png")
VERIFY = Path(r"d:\Users\admin_mak\project\pdf-handler_DEV\docs\development\mockups")

# Calibrated on gui-drawings.png (1024x596) — strip under right-pane toolbar
# Page1 selection wash/border observed around x≈575–655, y≈112–182
PAGE1 = (575, 112, 655, 182)
GAP = 8
THUMB_W = PAGE1[2] - PAGE1[0]
THUMB_H = PAGE1[3] - PAGE1[1]

BORDER = (85, 131, 191)
WASH = (190, 220, 246)


def thumb_boxes() -> list[tuple[int, int, int, int]]:
    x0, y0, _, y1 = PAGE1
    boxes = []
    x = x0
    for _ in range(3):
        boxes.append((x, y0, x + THUMB_W, y1))
        x += THUMB_W + GAP
    return boxes


def paint_selected(base: Image.Image, box: tuple[int, int, int, int]) -> None:
    """Paint selection wash + border matching LeafDesk strip style."""
    x0, y0, x1, y1 = box
    layer = Image.new("RGBA", base.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    # wash (match existing page1 look)
    d.rectangle([x0 + 2, y0 + 2, x1 - 2, y1 - 2], fill=(*WASH, 160))
    for t in range(3):
        d.rectangle([x0 + t, y0 + t, x1 - t, y1 - t], outline=(*BORDER, 255))
    composed = Image.alpha_composite(base.convert("RGBA"), layer)
    base.paste(composed)


def main() -> None:
    im = Image.open(SRC).convert("RGBA")
    boxes = thumb_boxes()
    print("boxes", boxes)

    # Clear existing single selection first by covering page1 with neutral white then repaint all
    # Sample background just left of page1
    bg = (255, 255, 255, 255)
    d0 = ImageDraw.Draw(im)
    # white out old selection area slightly expanded
    ox0, oy0, ox1, oy1 = PAGE1
    d0.rectangle([ox0 - 1, oy0 - 1, ox1 + 1, oy1 + 1], fill=bg)

    # Re-blit original content under selection from source (so we don't erase drawings)
    src_rgb = Image.open(SRC).convert("RGBA")
    for box in boxes:
        x0, y0, x1, y1 = box
        crop = src_rgb.crop((x0, y0, x1, y1))
        im.paste(crop, (x0, y0))

    for box in boxes:
        paint_selected(im, box)

    # Caption chip above strip
    d = ImageDraw.Draw(im)
    try:
        font = ImageFont.truetype(r"C:\Windows\Fonts\YuGothM.ttc", 16)
    except OSError:
        font = ImageFont.load_default()
    label = "Ctrl / Shift で複数ページを選択"
    # measure
    bbox = d.textbbox((0, 0), label, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    pad_x, pad_y = 12, 6
    lx = boxes[0][0]
    ly = boxes[0][1] - th - pad_y * 2 - 6
    if ly < 70:
        ly = boxes[0][3] + 6
    d.rounded_rectangle(
        [lx, ly, lx + tw + pad_x * 2, ly + th + pad_y * 2],
        radius=6,
        fill=(33, 37, 41, 220),
    )
    d.text((lx + pad_x, ly + pad_y), label, fill=(255, 255, 255, 255), font=font)

    out = im.convert("RGB")
    OUT.parent.mkdir(parents=True, exist_ok=True)
    out.save(OUT, optimize=True)
    VERIFY.mkdir(parents=True, exist_ok=True)
    out.crop((boxes[0][0] - 30, 70, min(out.width - 5, boxes[-1][2] + 30), boxes[0][3] + 40)).save(
        VERIFY / "_hp_multiselect_strip.png"
    )
    print(f"wrote {OUT} {out.size}")


if __name__ == "__main__":
    main()
