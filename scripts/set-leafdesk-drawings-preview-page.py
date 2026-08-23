#!/usr/bin/env python3
"""Set gui-drawings.png preview to a specific page of merged.pdf (default: page 3)."""

from __future__ import annotations

import argparse
from pathlib import Path

import fitz
import numpy as np
from PIL import Image, ImageDraw, ImageFont

ASSETS = Path(r"d:\Users\admin_mak\project\office-goplan\assets\pdfhandler")
BASE = ASSETS / "gui-main.png"
OUT = ASSETS / "gui-drawings.png"
VERIFY = Path(r"d:\Users\admin_mak\project\pdf-handler_DEV\docs\development\mockups")
DEMO_PDF = (
    Path.home() / "Documents" / "LeafDesk-HP-Demo" / "01_図面管理" / "merged.pdf"
)

# Calibrated on gui-drawings.png (1024x596)
STRIP_BOXES = [(575, 112, 655, 182), (663, 112, 743, 182), (751, 112, 831, 182)]
PREVIEW_BOX = (552, 196, 998, 558)
PAGE_LABEL_BOX = (768, 78, 812, 96)
GRID_PREVIEW_BOX = (362, 318, 466, 405)
GRID_PAGER_BAR_BOX = (360, 428, 470, 445)
GRID_PAGER_LABEL_BOX = (397, 432, 429, 437)

BORDER = (85, 131, 191)
WASH = (190, 220, 246)
GRAY_BORDER = (170, 170, 175)


def load_font(size: int) -> ImageFont.ImageFont:
    for name in ("YuGothM.ttc", "YuGothR.ttc", "meiryo.ttc", "segoeui.ttf"):
        path = Path(r"C:\Windows\Fonts") / name
        if path.exists():
            try:
                return ImageFont.truetype(str(path), size)
            except OSError:
                continue
    return ImageFont.load_default()


def render_page(pdf: Path, page_index: int, size: tuple[int, int]) -> Image.Image:
    doc = fitz.open(pdf)
    page = doc.load_page(page_index)
    tw, th = size
    zoom = min(tw / page.rect.width, th / page.rect.height) * 2.4
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    doc.close()

    scale = max(tw / im.width, th / im.height)
    nw, nh = max(tw, int(im.width * scale)), max(th, int(im.height * scale))
    im2 = im.resize((nw, nh), Image.Resampling.LANCZOS)
    left = max(0, (nw - tw) // 2)
    top = max(0, (nh - th) // 2)
    return im2.crop((left, top, left + tw, top + th))


def paste_in_box(canvas: Image.Image, thumb: Image.Image, box: tuple[int, int, int, int]) -> None:
    x0, y0, x1, y1 = box
    tw, th = x1 - x0, y1 - y0
    if thumb.size != (tw, th):
        thumb = thumb.resize((tw, th), Image.Resampling.LANCZOS)
    canvas.paste(thumb, (x0, y0))


def paint_strip_selection(canvas: Image.Image, box: tuple[int, int, int, int]) -> None:
    x0, y0, x1, y1 = box
    layer = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    d.rectangle([x0 + 2, y0 + 2, x1 - 2, y1 - 2], fill=(*WASH, 160))
    for t in range(3):
        d.rectangle([x0 + t, y0 + t, x1 - t, y1 - t], outline=(*BORDER, 255))
    composed = Image.alpha_composite(canvas.convert("RGBA"), layer)
    canvas.paste(composed.convert("RGB"))


def paint_strip_neutral(canvas: Image.Image, box: tuple[int, int, int, int]) -> None:
    x0, y0, x1, y1 = box
    d = ImageDraw.Draw(canvas)
    d.rectangle([x0, y0, x1 - 1, y1 - 1], outline=GRAY_BORDER)


def replace_page_label(
    canvas: Image.Image,
    box: tuple[int, int, int, int],
    text: str,
    *,
    font_size: int,
    fg: tuple[int, int, int] = (60, 60, 60),
    bg: tuple[int, int, int] | None = None,
) -> None:
    x0, y0, x1, y1 = box
    if bg is None:
        patch = canvas.crop(box)
        arr = np.asarray(patch)
        bg = tuple(int(v) for v in arr.mean(axis=(0, 1)))
        if sum(bg) > 700:
            bg = (245, 245, 245)
    d = ImageDraw.Draw(canvas)
    d.rectangle([x0, y0, x1, y1], fill=bg)
    font = load_font(font_size)
    bbox = d.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    d.text(
        (x0 + (x1 - x0 - tw) // 2, y0 + (y1 - y0 - th) // 2 - 1),
        text,
        fill=fg,
        font=font,
    )


def build_preview_page(page_no: int) -> Image.Image:
    if not DEMO_PDF.exists():
        raise SystemExit(f"missing {DEMO_PDF}")
    if page_no < 1 or page_no > 3:
        raise SystemExit("merged.pdf has 3 pages; use --page 1..3")

    base = Image.open(BASE).convert("RGB")
    pristine = base.copy()
    page_index = page_no - 1

    strip_size = (
        STRIP_BOXES[0][2] - STRIP_BOXES[0][0],
        STRIP_BOXES[0][3] - STRIP_BOXES[0][1],
    )
    grid_size = (
        GRID_PREVIEW_BOX[2] - GRID_PREVIEW_BOX[0],
        GRID_PREVIEW_BOX[3] - GRID_PREVIEW_BOX[1],
    )
    preview_size = (
        PREVIEW_BOX[2] - PREVIEW_BOX[0],
        PREVIEW_BOX[3] - PREVIEW_BOX[1],
    )

    strip_thumbs = [render_page(DEMO_PDF, i, strip_size) for i in range(3)]
    main_thumb = render_page(DEMO_PDF, page_index, preview_size)
    grid_thumb = render_page(DEMO_PDF, page_index, grid_size)

    d = ImageDraw.Draw(base)
    d.rectangle(PREVIEW_BOX, fill=(255, 255, 255))
    paste_in_box(base, main_thumb, PREVIEW_BOX)

    for i, box in enumerate(STRIP_BOXES):
        paste_in_box(base, strip_thumbs[i], box)
        if i + 1 == page_no:
            paint_strip_selection(base, box)
        else:
            paint_strip_neutral(base, box)

    label = f"{page_no}/3"
    replace_page_label(base, PAGE_LABEL_BOX, label, font_size=13)
    paste_in_box(base, grid_thumb, GRID_PREVIEW_BOX)
    base.paste(pristine.crop(GRID_PAGER_BAR_BOX), GRID_PAGER_BAR_BOX[:2])
    replace_page_label(
        base,
        GRID_PAGER_LABEL_BOX,
        label,
        font_size=10,
        fg=(255, 255, 255),
        bg=(45, 45, 48),
    )

    return base


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--page", type=int, default=3, help="1-based page number")
    args = parser.parse_args()

    out = build_preview_page(args.page)
    VERIFY.mkdir(parents=True, exist_ok=True)
    out.save(OUT, optimize=True)
    out.crop((540, 70, 1010, 565)).save(VERIFY / "_hp_strip_region2.png")
    print(f"wrote {OUT} page={args.page}/3 size={out.size}")


if __name__ == "__main__":
    main()
