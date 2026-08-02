#!/usr/bin/env python3
"""Generate fake PDFs for LeafDesk HP screenshots (no real customer data).

Output default: %USERPROFILE%\\Documents\\LeafDesk-HP-Demo\\
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

from reportlab.lib.colors import Color, HexColor, black, white
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

FONT = "YuGothic"
FONT_B = "YuGothicBold"


def register_fonts() -> None:
    yu_r = Path(r"C:\Windows\Fonts\YuGothR.ttc")
    yu_m = Path(r"C:\Windows\Fonts\YuGothM.ttc")
    yu_b = Path(r"C:\Windows\Fonts\YuGothB.ttc")
    if not yu_r.exists():
        raise SystemExit("Yu Gothic font not found under C:\\Windows\\Fonts")
    pdfmetrics.registerFont(TTFont(FONT, str(yu_r), subfontIndex=0))
    pdfmetrics.registerFont(TTFont(FONT_B, str(yu_b if yu_b.exists() else yu_m), subfontIndex=0))


def out_root() -> Path:
    if len(sys.argv) > 1:
        return Path(sys.argv[1]).expanduser().resolve()
    return Path.home() / "Documents" / "LeafDesk-HP-Demo"


def draw_title_block(
    c: canvas.Canvas,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    drawing_no: str,
    title: str,
    rev: str,
    scale: str,
    accent: Color,
) -> None:
    c.setStrokeColor(black)
    c.setLineWidth(1.2)
    c.setFillColor(white)
    c.rect(x, y, w, h, fill=1, stroke=1)
    c.setFillColor(accent)
    c.rect(x, y + h - 8 * mm, w, 8 * mm, fill=1, stroke=1)
    c.setFillColor(white)
    c.setFont(FONT_B, 9)
    c.drawString(x + 2 * mm, y + h - 5.5 * mm, "TITLE BLOCK (DEMO)")
    c.setFillColor(black)
    c.setFont(FONT, 8)
    rows = [
        ("図番", drawing_no),
        ("名称", title),
        ("改訂", rev),
        ("縮尺", scale),
        ("作成", "デモ用・架空データ"),
    ]
    row_h = (h - 8 * mm) / len(rows)
    for i, (k, v) in enumerate(rows):
        ry = y + h - 8 * mm - (i + 1) * row_h
        c.line(x, ry, x + w, ry)
        c.line(x + 18 * mm, ry, x + 18 * mm, ry + row_h)
        c.setFont(FONT, 7)
        c.drawString(x + 1.5 * mm, ry + 2 * mm, k)
        c.setFont(FONT_B, 8)
        c.drawString(x + 19 * mm, ry + 2 * mm, v)


def pdf_drawing_bracket(path: Path) -> None:
    """L-bracket style — distinctive silhouette."""
    c = canvas.Canvas(str(path), pagesize=landscape(A4))
    w, h = landscape(A4)
    margin = 12 * mm
    c.setStrokeColor(HexColor("#1a365d"))
    c.setLineWidth(2)
    c.rect(margin, margin, w - 2 * margin, h - 2 * margin)

    # Outer frame line (drawing border look)
    c.setStrokeColor(HexColor("#2b6cb0"))
    c.setLineWidth(0.8)
    c.rect(margin + 4 * mm, margin + 4 * mm, w - 2 * margin - 8 * mm, h - 2 * margin - 8 * mm)

    # L shape
    ox, oy = 55 * mm, 40 * mm
    c.setFillColor(HexColor("#ebf8ff"))
    c.setStrokeColor(HexColor("#2c5282"))
    c.setLineWidth(1.5)
    path_l = c.beginPath()
    path_l.moveTo(ox, oy)
    path_l.lineTo(ox + 90 * mm, oy)
    path_l.lineTo(ox + 90 * mm, oy + 25 * mm)
    path_l.lineTo(ox + 30 * mm, oy + 25 * mm)
    path_l.lineTo(ox + 30 * mm, oy + 70 * mm)
    path_l.lineTo(ox, oy + 70 * mm)
    path_l.close()
    c.drawPath(path_l, fill=1, stroke=1)

    # holes
    c.setFillColor(HexColor("#90cdf4"))
    for hx, hy in [(ox + 12 * mm, oy + 12 * mm), (ox + 70 * mm, oy + 12 * mm), (ox + 12 * mm, oy + 55 * mm)]:
        c.circle(hx, hy, 4 * mm, fill=1, stroke=1)

    # dimension-ish labels
    c.setFillColor(HexColor("#2d3748"))
    c.setFont(FONT, 9)
    c.drawCentredString(ox + 45 * mm, oy - 8 * mm, "90")
    c.drawString(ox - 12 * mm, oy + 30 * mm, "70")

    draw_title_block(
        c,
        w - margin - 70 * mm,
        margin + 6 * mm,
        65 * mm,
        42 * mm,
        drawing_no="DWG-A100",
        title="ブラケット（L形）",
        rev="A",
        scale="1:2",
        accent=HexColor("#2b6cb0"),
    )
    c.setFont(FONT_B, 14)
    c.setFillColor(HexColor("#1a365d"))
    c.drawString(margin + 8 * mm, h - margin - 10 * mm, "図面デモ — ブラケット")
    c.setFont(FONT, 8)
    c.setFillColor(HexColor("#718096"))
    c.drawString(margin + 8 * mm, h - margin - 16 * mm, "HP撮影用ダミー。実在の図面ではありません。")
    c.save()


def pdf_drawing_shaft(path: Path) -> None:
    """Round shaft — different visual from bracket."""
    c = canvas.Canvas(str(path), pagesize=landscape(A4))
    w, h = landscape(A4)
    margin = 12 * mm
    c.setStrokeColor(HexColor("#22543d"))
    c.setLineWidth(2)
    c.rect(margin, margin, w - 2 * margin, h - 2 * margin)

    cx, cy = 95 * mm, 105 * mm
    c.setFillColor(HexColor("#c6f6d5"))
    c.setStrokeColor(HexColor("#276749"))
    c.setLineWidth(1.5)
    # side view bar
    c.roundRect(50 * mm, 70 * mm, 140 * mm, 28 * mm, 4 * mm, fill=1, stroke=1)
    # end circles (different diameters look)
    c.setFillColor(HexColor("#9ae6b4"))
    c.circle(50 * mm, 84 * mm, 14 * mm, fill=1, stroke=1)
    c.circle(190 * mm, 84 * mm, 10 * mm, fill=1, stroke=1)
    # centerline
    c.setDash(3, 3)
    c.setStrokeColor(HexColor("#38a169"))
    c.line(40 * mm, 84 * mm, 200 * mm, 84 * mm)
    c.setDash()

    c.setFillColor(HexColor("#22543d"))
    c.setFont(FONT, 9)
    c.drawCentredString(120 * mm, 58 * mm, "φ28 / φ20  （デモ寸法）")

    # top view hint — concentric circles elsewhere
    c.setFillColor(HexColor("#f0fff4"))
    c.circle(230 * mm, 130 * mm, 28 * mm, fill=1, stroke=1)
    c.circle(230 * mm, 130 * mm, 18 * mm, fill=0, stroke=1)
    c.circle(230 * mm, 130 * mm, 8 * mm, fill=0, stroke=1)

    draw_title_block(
        c,
        w - margin - 70 * mm,
        margin + 6 * mm,
        65 * mm,
        42 * mm,
        drawing_no="DWG-B220",
        title="シャフト",
        rev="B",
        scale="1:1",
        accent=HexColor("#276749"),
    )
    c.setFont(FONT_B, 14)
    c.setFillColor(HexColor("#22543d"))
    c.drawString(margin + 8 * mm, h - margin - 10 * mm, "図面デモ — シャフト")
    c.setFont(FONT, 8)
    c.setFillColor(HexColor("#718096"))
    c.drawString(margin + 8 * mm, h - margin - 16 * mm, "HP撮影用ダミー。実在の図面ではありません。")
    c.save()


def pdf_drawing_housing(path: Path) -> None:
    """Rectangular housing with hole pattern — third silhouette."""
    c = canvas.Canvas(str(path), pagesize=landscape(A4))
    w, h = landscape(A4)
    margin = 12 * mm
    c.setStrokeColor(HexColor("#744210"))
    c.setLineWidth(2)
    c.rect(margin, margin, w - 2 * margin, h - 2 * margin)

    bx, by, bw, bh = 45 * mm, 45 * mm, 130 * mm, 95 * mm
    c.setFillColor(HexColor("#fefcbf"))
    c.setStrokeColor(HexColor("#975a16"))
    c.setLineWidth(1.5)
    c.rect(bx, by, bw, bh, fill=1, stroke=1)
    c.setFillColor(HexColor("#faf089"))
    c.rect(bx + 15 * mm, by + 15 * mm, bw - 30 * mm, bh - 30 * mm, fill=1, stroke=1)

    c.setFillColor(HexColor("#d69e2e"))
    for i in range(3):
        for j in range(2):
            c.circle(bx + 35 * mm + i * 30 * mm, by + 35 * mm + j * 30 * mm, 5 * mm, fill=1, stroke=1)

    c.setFillColor(HexColor("#744210"))
    c.setFont(FONT_B, 11)
    c.drawCentredString(bx + bw / 2, by + bh + 6 * mm, "HOUSING — DEMO PATTERN")

    draw_title_block(
        c,
        w - margin - 70 * mm,
        margin + 6 * mm,
        65 * mm,
        42 * mm,
        drawing_no="DWG-C315",
        title="ハウジング",
        rev="C",
        scale="1:5",
        accent=HexColor("#b7791f"),
    )
    c.setFont(FONT_B, 14)
    c.setFillColor(HexColor("#744210"))
    c.drawString(margin + 8 * mm, h - margin - 10 * mm, "図面デモ — ハウジング")
    c.setFont(FONT, 8)
    c.setFillColor(HexColor("#718096"))
    c.drawString(margin + 8 * mm, h - margin - 16 * mm, "HP撮影用ダミー。実在の図面ではありません。")
    c.save()


def pdf_po_sample_shosha(path: Path) -> None:
    """Customer A — red header, dense table (portrait)."""
    c = canvas.Canvas(str(path), pagesize=A4)
    w, h = A4
    c.setFillColor(HexColor("#c53030"))
    c.rect(0, h - 28 * mm, w, 28 * mm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(FONT_B, 18)
    c.drawCentredString(w / 2, h - 18 * mm, "注文書")
    c.setFont(FONT, 9)
    c.drawString(15 * mm, h - 25 * mm, "株式会社サンプル商事（架空）")

    c.setFillColor(black)
    c.setFont(FONT, 10)
    y = h - 40 * mm
    lines = [
        "注文番号: PO-DEMO-2026-081",
        "発注日: 2026-07-15",
        "納入先: サンプル商事 本社倉庫（架空住所）",
        "支払条件: 月末締め翌月末払い（デモ）",
    ]
    for line in lines:
        c.drawString(18 * mm, y, line)
        y -= 7 * mm

    # table
    y -= 4 * mm
    cols = [18 * mm, 55 * mm, 95 * mm, 130 * mm, 165 * mm]
    headers = ["No", "品名", "図番", "数量", "備考"]
    c.setFillColor(HexColor("#fed7d7"))
    c.rect(18 * mm, y - 2 * mm, w - 36 * mm, 9 * mm, fill=1, stroke=1)
    c.setFillColor(black)
    c.setFont(FONT_B, 9)
    for i, hd in enumerate(headers):
        c.drawString(cols[i], y, hd)
    y -= 10 * mm
    c.setFont(FONT, 9)
    rows = [
        ("1", "ブラケット", "DWG-A100", "20", "塗装あり"),
        ("2", "シャフト", "DWG-B220", "10", "—"),
        ("3", "ハウジング", "DWG-C315", "5", "検査成績同梱"),
    ]
    for row in rows:
        for i, cell in enumerate(row):
            c.drawString(cols[i], y, cell)
        c.setStrokeColor(HexColor("#e2e8f0"))
        c.line(18 * mm, y - 2 * mm, w - 18 * mm, y - 2 * mm)
        y -= 9 * mm

    c.setFillColor(HexColor("#718096"))
    c.setFont(FONT, 8)
    c.drawString(18 * mm, 18 * mm, "※ HP撮影用ダミー注文書。実在の取引先・取引ではありません。")
    # big stamp-like mark for thumbnail uniqueness
    c.setStrokeColor(HexColor("#c53030"))
    c.setLineWidth(2)
    c.circle(w - 40 * mm, 55 * mm, 18 * mm, fill=0, stroke=1)
    c.setFillColor(HexColor("#c53030"))
    c.setFont(FONT_B, 11)
    c.drawCentredString(w - 40 * mm, 53 * mm, "DEMO")
    c.save()


def pdf_po_kaku_kogyo(path: Path) -> None:
    """Customer B — blue two-column layout (different look from A)."""
    c = canvas.Canvas(str(path), pagesize=A4)
    w, h = A4
    c.setStrokeColor(HexColor("#2b6cb0"))
    c.setLineWidth(3)
    c.rect(12 * mm, 12 * mm, w - 24 * mm, h - 24 * mm)
    c.setFillColor(HexColor("#2b6cb0"))
    c.setFont(FONT_B, 16)
    c.drawString(20 * mm, h - 28 * mm, "PURCHASE ORDER")
    c.setFont(FONT, 11)
    c.drawString(20 * mm, h - 36 * mm, "架空工業株式会社（デモ顧客B）")

    # left info box
    c.setFillColor(HexColor("#ebf8ff"))
    c.rect(20 * mm, h - 95 * mm, 85 * mm, 50 * mm, fill=1, stroke=1)
    c.setFillColor(black)
    c.setFont(FONT, 9)
    info = [
        "PO No. KAKU-7781",
        "Date 2026-07-20",
        "Buyer: 購買部（架空）",
        "Ship to: 第2工場",
        "Incoterms: デモ用",
    ]
    iy = h - 52 * mm
    for t in info:
        c.drawString(24 * mm, iy, t)
        iy -= 7 * mm

    # right logo block
    c.setFillColor(HexColor("#bee3f8"))
    c.rect(115 * mm, h - 95 * mm, 70 * mm, 50 * mm, fill=1, stroke=1)
    c.setFillColor(HexColor("#2c5282"))
    c.setFont(FONT_B, 20)
    c.drawCentredString(150 * mm, h - 72 * mm, "KAKU")
    c.setFont(FONT, 8)
    c.drawCentredString(150 * mm, h - 82 * mm, "DEMO CUSTOMER B")

    c.setFillColor(black)
    c.setFont(FONT_B, 10)
    c.drawString(20 * mm, h - 110 * mm, "Line items")
    y = h - 120 * mm
    c.setFont(FONT, 9)
    for i, (name, qty) in enumerate(
        [("検査治具セット", "2式"), ("要求仕様書参照部品", "1式"), ("緩衝材（指定）", "1式")],
        start=1,
    ):
        c.setFillColor(HexColor("#f7fafc") if i % 2 else white)
        c.rect(20 * mm, y - 3 * mm, w - 40 * mm, 10 * mm, fill=1, stroke=0)
        c.setFillColor(black)
        c.drawString(24 * mm, y, f"{i}. {name}")
        c.drawRightString(w - 24 * mm, y, qty)
        y -= 12 * mm

    c.setFillColor(HexColor("#718096"))
    c.setFont(FONT, 8)
    c.drawString(20 * mm, 18 * mm, "※ HP撮影用ダミー。レイアウト差の訴求用（顧客B）。")
    c.save()


def pdf_spec_inspection(path: Path) -> None:
    """Inspection criteria — text/table heavy, still distinct header."""
    c = canvas.Canvas(str(path), pagesize=A4)
    w, h = A4
    c.setFillColor(HexColor("#553c9a"))
    c.rect(0, h - 22 * mm, w, 22 * mm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(FONT_B, 14)
    c.drawString(15 * mm, h - 14 * mm, "検査基準書（デモ）")
    c.setFont(FONT, 8)
    c.drawRightString(w - 15 * mm, h - 14 * mm, "DOC-INSP-001")

    c.setFillColor(black)
    c.setFont(FONT, 10)
    y = h - 35 * mm
    paras = [
        "適用範囲: デモ製品シリーズ A/B/C（架空）",
        "判定: 外観・寸法・機能の3項目。不合格はロット隔離。",
        "サンプリング: 入荷ロットごと n=5（デモ値）",
    ]
    for p in paras:
        c.drawString(18 * mm, y, p)
        y -= 8 * mm

    y -= 4 * mm
    c.setFont(FONT_B, 10)
    c.drawString(18 * mm, y, "検査項目")
    y -= 10 * mm
    c.setFont(FONT, 9)
    for item, criteria in [
        ("外観", "傷・打痕なきこと（目視）"),
        ("寸法", "図面公差内（ノギス）"),
        ("機能", "指定トルクで動作すること"),
        ("表示", "図番・改訂が読み取れること"),
    ]:
        c.setStrokeColor(HexColor("#e2e8f0"))
        c.rect(18 * mm, y - 3 * mm, w - 36 * mm, 10 * mm, fill=0, stroke=1)
        c.setFillColor(HexColor("#553c9a"))
        c.setFont(FONT_B, 9)
        c.drawString(22 * mm, y, item)
        c.setFillColor(black)
        c.setFont(FONT, 9)
        c.drawString(50 * mm, y, criteria)
        y -= 12 * mm

    c.setFillColor(HexColor("#718096"))
    c.setFont(FONT, 8)
    c.drawString(18 * mm, 16 * mm, "※ テキスト中心帳票の例。サムネ識別力は図面より弱い、という対比用。")
    c.save()


def pdf_spec_rfq(path: Path) -> None:
    """RFQ / required spec — yellow banner, attachment list."""
    c = canvas.Canvas(str(path), pagesize=A4)
    w, h = A4
    c.setFillColor(HexColor("#d69e2e"))
    c.rect(0, h - 30 * mm, w, 30 * mm, fill=1, stroke=0)
    c.setFillColor(HexColor("#1a202c"))
    c.setFont(FONT_B, 16)
    c.drawCentredString(w / 2, h - 14 * mm, "見積依頼書・要求仕様書")
    c.setFont(FONT, 9)
    c.drawCentredString(w / 2, h - 22 * mm, "RFQ-DEMO-2026-044（架空）")

    c.setFillColor(black)
    c.setFont(FONT, 10)
    y = h - 45 * mm
    for line in [
        "依頼元: 株式会社サンプル商事 技術部（架空）",
        "件名: ブラケット一式の試作見積",
        "希望納期: 2026-09-30（デモ）",
        "添付資料:",
    ]:
        c.drawString(18 * mm, y, line)
        y -= 8 * mm

    c.setFont(FONT_B, 10)
    for att in ["・DWG-A100 ブラケット図面", "・DWG-B220 シャフト図面", "・検査基準書 DOC-INSP-001"]:
        c.setFillColor(HexColor("#744210"))
        c.drawString(24 * mm, y, att)
        y -= 8 * mm

    y -= 6 * mm
    c.setFillColor(HexColor("#fffff0"))
    c.setStrokeColor(HexColor("#d69e2e"))
    c.rect(18 * mm, y - 45 * mm, w - 36 * mm, 50 * mm, fill=1, stroke=1)
    c.setFillColor(black)
    c.setFont(FONT, 9)
    c.drawString(22 * mm, y - 5 * mm, "要求概要（デモ文言）")
    c.setFont(FONT, 8)
    for t in [
        "1. 材質は指定図面どおり。代替案は見積時に明示すること。",
        "2. 表面処理は顧客A仕様を優先。",
        "3. 初回ロットは検査成績書を添付すること。",
    ]:
        y -= 8 * mm
        c.drawString(22 * mm, y - 5 * mm, t)

    c.setFillColor(HexColor("#718096"))
    c.setFont(FONT, 8)
    c.drawString(18 * mm, 16 * mm, "※ HP撮影用ダミー。見積依頼＋添付図面シナリオ用。")
    c.save()


def write_readme(root: Path) -> None:
    text = """LeafDesk HP 撮影用ダミー PDF
================================

実在の顧客・図面・注文書ではありません。office-goplan 商品ページ用スクリーンショット専用です。

フォルダ構成
------------
01_図面管理/     … 見た目の違う図面 3 点（◎図面管理）
02_顧客注文書/   … レイアウト差のある注文書 2 点（◎顧客注文書）
03_添付混在/     … 図面＋検査基準＋見積依頼＋注文書（◎添付の見分け）
04_見積依頼/     … 見積依頼・要求仕様＋関連図面（◎RFQ）

撮影のコツ
----------
1. このフォルダ（または 01〜04 のいずれか）を LeafDesk の「お気に入り」に追加する
2. ツリーは「お気に入り」だけ展開し、D: など実ドライブは畳む／映さない
3. タイトルバーが LeafDesk、ステータスが Standard 版であることを確認
4. サムネイル表示で図面の枠・形状差が分かる状態で撮影

再生成
------
office-goplan/scripts/generate-leafdesk-hp-demo-pdfs.py
"""
    (root / "README.txt").write_text(text, encoding="utf-8")


def main() -> None:
    register_fonts()
    root = out_root()
    if root.exists():
        shutil.rmtree(root)
    dirs = {
        "draw": root / "01_図面管理",
        "po": root / "02_顧客注文書",
        "mix": root / "03_添付混在",
        "rfq": root / "04_見積依頼",
    }
    for d in dirs.values():
        d.mkdir(parents=True)

    files = {
        "bracket": dirs["draw"] / "DWG-A100_ブラケット.pdf",
        "shaft": dirs["draw"] / "DWG-B220_シャフト.pdf",
        "housing": dirs["draw"] / "DWG-C315_ハウジング.pdf",
        "po_a": dirs["po"] / "サンプル商事_注文書.pdf",
        "po_b": dirs["po"] / "架空工業_注文書.pdf",
        "insp": dirs["mix"] / "DOC-INSP-001_検査基準書.pdf",
        "rfq": dirs["rfq"] / "RFQ-DEMO_見積依頼_要求仕様.pdf",
    }

    pdf_drawing_bracket(files["bracket"])
    pdf_drawing_shaft(files["shaft"])
    pdf_drawing_housing(files["housing"])
    pdf_po_sample_shosha(files["po_a"])
    pdf_po_kaku_kogyo(files["po_b"])
    pdf_spec_inspection(files["insp"])
    pdf_spec_rfq(files["rfq"])

    # mix folder: copies for "attachments without opening"
    for src_key, name in [
        ("bracket", "DWG-A100_ブラケット.pdf"),
        ("shaft", "DWG-B220_シャフト.pdf"),
        ("po_a", "サンプル商事_注文書.pdf"),
        ("rfq", "RFQ-DEMO_見積依頼_要求仕様.pdf"),
    ]:
        shutil.copy2(files[src_key], dirs["mix"] / name)

    # rfq folder companions
    shutil.copy2(files["bracket"], dirs["rfq"] / "添付_DWG-A100_ブラケット.pdf")
    shutil.copy2(files["shaft"], dirs["rfq"] / "添付_DWG-B220_シャフト.pdf")
    shutil.copy2(files["insp"], dirs["rfq"] / "添付_DOC-INSP-001_検査基準書.pdf")

    write_readme(root)
    print(f"Generated under: {root}")
    for p in sorted(root.rglob("*.pdf")):
        print(f"  {p.relative_to(root)}")


if __name__ == "__main__":
    main()
