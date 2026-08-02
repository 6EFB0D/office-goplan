#!/usr/bin/env python3
"""Trim capture shadows and mask drives/system folders for LeafDesk HP screenshots."""

from __future__ import annotations

from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw

ASSETS = Path(
    r"C:\Users\admin_mak\.cursor\projects\d-Users-admin-mak-project-pdf-handler-DEV\assets"
)
OUT_DIR = Path(r"d:\Users\admin_mak\project\office-goplan\assets\pdfhandler")
DEMO_DIR = Path.home() / "Documents" / "LeafDesk-HP-Demo" / "screenshots"
VERIFY = Path(r"d:\Users\admin_mak\project\pdf-handler_DEV\docs\development\mockups")

SOURCES = {
    "gui-main.png": ASSETS
    / "c__Users_admin_mak_AppData_Roaming_Cursor_User_workspaceStorage_da820770282a5af9193dc7f909db0d08_images_image-baac1277-a57b-414e-9cc1-41194dc6d051.png",
    "gui-po.png": ASSETS
    / "c__Users_admin_mak_AppData_Roaming_Cursor_User_workspaceStorage_da820770282a5af9193dc7f909db0d08_images_image-02d135c2-b258-4bb8-abe7-08d2c62ef530.png",
    "gui-mix.png": ASSETS
    / "c__Users_admin_mak_AppData_Roaming_Cursor_User_workspaceStorage_da820770282a5af9193dc7f909db0d08_images_image-ba28e569-4164-4da1-abee-c1917ca2145c.png",
    "gui-rfq.png": ASSETS
    / "c__Users_admin_mak_AppData_Roaming_Cursor_User_workspaceStorage_da820770282a5af9193dc7f909db0d08_images_image-fd0d9efe-3f76-4060-a95e-07d1eba26b4e.png",
}


def trim_shadows(arr: np.ndarray) -> tuple[int, int, int, int]:
    h, w = arr.shape[:2]
    x0 = 0
    while x0 < w and float(arr[:, x0].mean()) < 45 and float(arr[:, x0].std()) < 12:
        x0 += 1
    x1 = w - 1
    while x1 > x0 and float(arr[:, x1].mean()) > 248 and float(arr[:, x1].std()) < 25:
        x1 -= 1
    x1 += 1
    y0 = 0
    while y0 < h and float(arr[y0].mean()) < 45 and float(arr[y0].std()) < 12:
        y0 += 1
    y1 = h - 1
    # bottom window drop-shadow / dark frame line
    while y1 > y0 and float(arr[y1].mean()) < 90 and float(arr[y1].std()) < 25:
        y1 -= 1
    y1 += 1
    return x0, y0, x1, y1


def find_splitter_x(arr: np.ndarray) -> int:
    h, w = arr.shape[:2]
    y0, y1 = int(h * 0.18), int(h * 0.65)
    best_x, best_score = int(w * 0.21), 1e9
    for x in range(int(w * 0.16), int(w * 0.30)):
        strip = arr[y0:y1, x].astype(np.float32)
        mean = float(strip.mean())
        std = float(strip.std())
        if 165 <= mean <= 225 and std < 22:
            score = std + abs(mean - 195) * 0.05
            if score < best_score:
                best_score, best_x = score, x
    return best_x


def mask_tree(im: Image.Image) -> None:
    """Hide Work..Downloads and C:/D:/X: — keep Computer + Favorites demo branch.

    Coordinates calibrated on 1024x~598 LeafDesk captures with demo tree expanded.
    """
    arr = np.asarray(im)
    h, w = arr.shape[:2]
    split = find_splitter_x(arr)
    # scale from reference height 598
    sy = h / 598.0
    sys_y0 = int(105 * sy)
    sys_y1 = int(170 * sy)
    # C:/D:/X: sit just under 04_見積依頼 (~y266 on 598px captures)
    drives_y0 = int(266 * sy)
    # stop above status bar
    drives_y1 = int(h * 0.92)

    draw = ImageDraw.Draw(im)
    tx0, tx1 = 1, max(2, split - 1)
    draw.rectangle([tx0, sys_y0, tx1, sys_y1], fill=(255, 255, 255))
    draw.rectangle([tx0, drives_y0, tx1, drives_y1], fill=(255, 255, 255))


def process(src: Path, dest_name: str) -> Path:
    im = Image.open(src).convert("RGB")
    arr = np.asarray(im)
    box = trim_shadows(arr)
    cropped = im.crop(box)
    mask_tree(cropped)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    DEMO_DIR.mkdir(parents=True, exist_ok=True)
    VERIFY.mkdir(parents=True, exist_ok=True)

    out = OUT_DIR / dest_name
    cropped.save(out, optimize=True)
    cropped.save(DEMO_DIR / dest_name, optimize=True)

    # verify crops
    cropped.crop((0, 0, min(230, cropped.width), min(200, cropped.height))).save(
        VERIFY / f"_hp_top_{dest_name}"
    )
    cropped.crop((0, 240, min(230, cropped.width), min(370, cropped.height))).save(
        VERIFY / f"_hp_bot_{dest_name}"
    )

    print(f"{dest_name}: {im.size} trim={box} -> {cropped.size}")
    return out


def main() -> None:
    for dest, src in SOURCES.items():
        if not src.exists():
            raise SystemExit(f"missing: {src}")
        process(src, dest)


if __name__ == "__main__":
    main()
