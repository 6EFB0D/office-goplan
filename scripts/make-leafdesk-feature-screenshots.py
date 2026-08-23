#!/usr/bin/env python3
"""Rebuild LeafDesk HP screenshots:
  - gui-multiselect.png: strip with unselected pages on sides + middle multi-select
  - gui-rename-f2.png: F2 rename TextBox under selected thumbnail while preview visible
"""

from __future__ import annotations

from pathlib import Path

import fitz
import numpy as np
from PIL import Image, ImageDraw, ImageFont

ASSETS = Path(r"d:\Users\admin_mak\project\office-goplan\assets\pdfhandler")
BASE = ASSETS / "gui-drawings.png"
OUT_MULTI = ASSETS / "gui-multiselect.png"
OUT_RENAME = ASSETS / "gui-rename-f2.png"
VERIFY = Path(r"d:\Users\admin_mak\project\pdf-handler_DEV\docs\development\mockups")

DEMO_PDF = next(Path.home().joinpath("Documents", "LeafDesk-HP-Demo").glob("01_*")) / "strip-demo-6pages.pdf"

STRIP_Y0, STRIP_Y1 = 108, 188
STRIP_X0, STRIP_X1 = 555, 1000
THUMB_W, THUMB_H = 72, 68
GAP = 8
BORDER = (85, 131, 191)
WASH = (190, 220, 246)
GRAY_BORDER = (170, 170, 175)


def load_font(size: int) -> ImageFont.ImageFont:
    for name in ("YuGothM.ttc", "YuGothR.ttc", "meiryo.ttc"):
        path = Path(r"C:\Windows\Fonts") / name
        if path.exists():
            try:
                return ImageFont.truetype(str(path), size)
            except OSError:
                continue
    return ImageFont.load_default()


def render_page_thumbs(pdf: Path, count: int, size: tuple[int, int]) -> list[Image.Image]:
    doc = fitz.open(pdf)
    out: list[Image.Image] = []
    tw, th = size
    for i in range(min(count, doc.page_count)):
        page = doc.load_page(i)
        zoom = min(tw / page.rect.width, th / page.rect.height) * 2.2
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        scale = max(tw / im.width, th / im.height)
        nw, nh = max(tw, int(im.width * scale)), max(th, int(im.height * scale))
        im2 = im.resize((nw, nh), Image.Resampling.LANCZOS)
        left = max(0, (nw - tw) // 2)
        top = max(0, (nh - th) // 2)
        out.append(im2.crop((left, top, left + tw, top + th)))
    doc.close()
    return out


def paste_strip_card(
    canvas: Image.Image,
    thumb: Image.Image,
    xy: tuple[int, int],
    *,
    selected: bool,
    page_no: int,
) -> None:
    x, y = xy
    # white pad under thumb
    d = ImageDraw.Draw(canvas)
    d.rectangle([x, y, x + THUMB_W - 1, y + THUMB_H + 14], fill=(255, 255, 255))
    canvas.paste(thumb, (x, y))

    if selected:
        overlay = Image.new("RGBA", (THUMB_W, THUMB_H), (*WASH, 95))
        rgba = canvas.convert("RGBA")
        rgba.alpha_composite(overlay, dest=(x, y))
        canvas.paste(rgba.convert("RGB"))
        d = ImageDraw.Draw(canvas)
        for t in range(3):
            d.rectangle(
                [x + t, y + t, x + THUMB_W - 1 - t, y + THUMB_H - 1 - t],
                outline=BORDER,
            )
    else:
        d.rectangle([x, y, x + THUMB_W - 1, y + THUMB_H - 1], outline=GRAY_BORDER)

    font = load_font(11)
    label = str(page_no)
    bbox = d.textbbox((0, 0), label, font=font)
    tw = bbox[2] - bbox[0]
    d.text((x + (THUMB_W - tw) // 2, y + THUMB_H + 1), label, fill=(70, 70, 70), font=font)


def build_multiselect() -> Image.Image:
    base = Image.open(BASE).convert("RGB")
    thumbs = render_page_thumbs(DEMO_PDF, 6, (THUMB_W, THUMB_H))

    d = ImageDraw.Draw(base)
    d.rectangle([STRIP_X0, STRIP_Y0 - 4, STRIP_X1, STRIP_Y1 + 10], fill=(255, 255, 255))

    n = 6
    selected = {2, 3, 4}
    total_w = n * THUMB_W + (n - 1) * GAP
    start_x = STRIP_X0 + max(4, (STRIP_X1 - STRIP_X0 - total_w) // 2)
    y = STRIP_Y0 + 2

    for i, thumb in enumerate(thumbs, start=1):
        x = start_x + (i - 1) * (THUMB_W + GAP)
        paste_strip_card(base, thumb, (x, y), selected=(i in selected), page_no=i)

    font = load_font(14)
    label = "Ctrl / Shift で複数ページを選択"
    bbox = d.textbbox((0, 0), label, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    lx, ly = start_x, max(74, STRIP_Y0 - th - 12)
    pad_x, pad_y = 10, 5
    d.rounded_rectangle(
        [lx, ly, lx + tw + pad_x * 2, ly + th + pad_y * 2],
        radius=6,
        fill=(33, 37, 41),
    )
    d.text((lx + pad_x, ly + pad_y), label, fill=(255, 255, 255), font=font)
    return base


def find_selected_thumb_name_area(arr: np.ndarray) -> tuple[int, int, int, int]:
    h, w = arr.shape[:2]
    blue = (
        (arr[:, :, 2] > 160)
        & (arr[:, :, 0] < 140)
        & (arr[:, :, 1] < 200)
        & (np.arange(w)[None, :] > int(w * 0.22))
        & (np.arange(w)[None, :] < int(w * 0.52))
        & (np.arange(h)[:, None] > 80)
        & (np.arange(h)[:, None] < 420)
    )
    ys, xs = np.where(blue)
    if len(xs) < 100:
        return (330, 318, 460, 342)
    x0, x1 = int(xs.min()), int(xs.max())
    y0, y1 = int(ys.min()), int(ys.max())
    name_y0 = y0 + int((y1 - y0) * 0.78)
    return (x0 + 10, name_y0, x1 - 10, min(y1 - 6, name_y0 + 24))


def build_rename_f2() -> Image.Image:
    base = Image.open(BASE).convert("RGB")
    arr = np.asarray(base)
    x0, y0, x1, y1 = find_selected_thumb_name_area(arr)
    print("rename box", (x0, y0, x1, y1))
    d = ImageDraw.Draw(base)

    # ensure enough height for textbox
    if y1 - y0 < 22:
        y1 = y0 + 22

    d.rectangle([x0, y0, x1, y1], fill=(255, 255, 255), outline=(0, 120, 215), width=2)
    font = load_font(12)
    text = "merged_図面セット"
    tb = d.textbbox((x0 + 4, y0 + 3), text, font=font)
    sel_r = min(x1 - 3, tb[2] + 4)
    d.rectangle([x0 + 3, y0 + 3, sel_r, y1 - 3], fill=(0, 120, 215))
    d.text((x0 + 4, y0 + 3), text, fill=(255, 255, 255), font=font)

    font2 = load_font(14)
    label = "F2：プレビュー表示のままファイル名を変更"
    bbox = d.textbbox((0, 0), label, font=font2)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    lx, ly = max(210, x0 - 20), y1 + 10
    pad_x, pad_y = 10, 5
    d.rounded_rectangle(
        [lx, ly, lx + tw + pad_x * 2, ly + th + pad_y * 2],
        radius=6,
        fill=(33, 37, 41),
    )
    d.text((lx + pad_x, ly + pad_y), label, fill=(255, 255, 255), font=font2)
    return base


def main() -> None:
    if not DEMO_PDF.exists():
        raise SystemExit(f"missing {DEMO_PDF}")
    VERIFY.mkdir(parents=True, exist_ok=True)

    multi = build_multiselect()
    multi.save(OUT_MULTI, optimize=True)
    multi.crop((STRIP_X0 - 10, 70, STRIP_X1, STRIP_Y1 + 36)).save(VERIFY / "_hp_multiselect_strip.png")
    print(f"wrote {OUT_MULTI}")

    rename = build_rename_f2()
    rename.save(OUT_RENAME, optimize=True)
    rename.crop((210, 90, 520, 400)).save(VERIFY / "_hp_rename_f2.png")
    print(f"wrote {OUT_RENAME}")


if __name__ == "__main__":
    main()
