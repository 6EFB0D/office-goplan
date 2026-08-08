# -*- coding: utf-8 -*-
"""Generate site-wide legal HTML from docs/legal/*.txt masters."""
from html import escape
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
LEGAL = ROOT / "docs" / "legal"

SHELL = """<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <link rel="icon" type="image/jpeg" href="assets/logo/logo-tab.jpg">
  <link rel="apple-touch-icon" sizes="180x180" href="assets/logo/apple-touch-icon-large.png">
  <link rel="manifest" href="site.webmanifest">
  <meta name="theme-color" content="#ffffff">
  <meta name="description" content="{desc}">
  <title>{title} | Office Go Plan</title>
  <link rel="canonical" href="{canonical}">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;600;700&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="styles.css">
  <script src="assets/js/legacy-redirect.js"></script>
  <script defer src="assets/js/ga4.js"></script>
</head>
<body>
  <header class="header">
    <div class="container">
      <a href="/" class="logo">
        <img src="assets/logo/logo-a.jpg" alt="Office Go Plan" class="logo-img">
      </a>
      <nav class="nav">
        <a href="/">ホーム</a>
        <a href="/#products">製品</a>
        <a href="/privacy-policy">プライバシーポリシー</a>
        <a href="/terms-of-service">利用規約</a>
        <a href="/specified-commercial-transactions">特定商取引法に基づく表記</a>
      </nav>
    </div>
  </header>

  <main>
    <article class="privacy-content">
      <div class="container">
        <p style="font-size:0.95rem;color:var(--color-text-muted);margin-bottom:1rem;">
          本{page_kind}は、当サイトおよび当社が提供するすべてのソフトウェア製品に共通して適用されます。
          アプリ内にも同内容を表示する場合があります。必要に応じて、ブラウザの印刷機能で保存できます。
        </p>
        <h1>{h1}</h1>
{body}
        <p style="margin-top: 32px; font-size: 0.9rem; color: var(--color-text-muted);">最終更新日：{date}</p>
      </div>
    </article>
  </main>

  <footer class="footer">
    <div class="container">
      <nav class="footer-nav">
        <a href="/">ホーム</a>
        <a href="/#products">製品</a>
        <a href="/privacy-policy">プライバシーポリシー</a>
        <a href="/terms-of-service">利用規約</a>
        <a href="/specified-commercial-transactions">特定商取引法に基づく表記</a>
      </nav>
      <p class="copyright">&copy; Office Go Plan. All rights reserved.</p>
    </div>
  </footer>
</body>
</html>
"""

REDIRECT = """<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="refresh" content="0; url={canonical}">
  <link rel="canonical" href="{canonical}">
  <script src="assets/js/legacy-redirect.js"></script>
  <title>移動しました | Office Go Plan</title>
</head>
<body>
  <p>利用規約・プライバシーポリシーの掲載場所を整理しました。<a href="{canonical}">最新の内容はこちら</a>をご覧ください。</p>
</body>
</html>
"""


def is_heading(line: str) -> bool:
    t = line.strip()
    if t.startswith("第") and "条" in t:
        return True
    if t in ("附則", "お問い合わせ", "要約", "制定・改訂"):
        return True
    return False


def txt_to_body(path: Path):
    lines = path.read_text(encoding="utf-8").splitlines()
    parts = []
    date = ""
    i = 0
    while i < len(lines) and not lines[i].strip():
        i += 1
    i += 1  # title
    while i < len(lines) and not lines[i].strip():
        i += 1
    if i < len(lines) and "最終更新日" in lines[i]:
        date = (
            lines[i]
            .replace("最終更新日", "")
            .replace(":", "")
            .replace("：", "")
            .strip()
        )
        i += 1

    buf = []

    def flush_para():
        nonlocal buf
        if not buf:
            return
        text = "<br>\n".join(escape(x.strip()) for x in buf if x.strip())
        parts.append(f"<p>{text}</p>")
        buf = []

    while i < len(lines):
        t = lines[i].strip()
        i += 1
        if not t or set(t) <= {"━", "─"}:
            flush_para()
            continue
        if is_heading(t):
            flush_para()
            parts.append(f"<h2>{escape(t)}</h2>")
            continue
        if t.startswith("•"):
            flush_para()
            items = [t.lstrip("•").strip()]
            while i < len(lines) and lines[i].strip().startswith("•"):
                items.append(lines[i].strip().lstrip("•").strip())
                i += 1
            parts.append(
                "<ul>\n"
                + "\n".join(f"<li>{escape(x)}</li>" for x in items)
                + "\n</ul>"
            )
            continue
        buf.append(t)
    flush_para()
    return "\n".join(parts), date or "2026年7月12日"


def main():
    jobs = [
        (
            LEGAL / "TERMS_OF_SERVICE.txt",
            "terms-of-service.html",
            "利用規約",
            "利用規約",
            "Office Go Plan 利用規約（全製品共通）",
            "利用規約",
            "https://office-goplan.com/terms-of-service",
        ),
        (
            LEGAL / "PRIVACY_POLICY.txt",
            "privacy-policy.html",
            "プライバシーポリシー",
            "プライバシーポリシー",
            "Office Go Plan プライバシーポリシー（全製品共通）",
            "プライバシーポリシー",
            "https://office-goplan.com/privacy-policy",
        ),
    ]
    for src, out_name, title, h1, desc, page_kind, canonical in jobs:
        body, date = txt_to_body(src)
        html = SHELL.format(
            desc=desc,
            title=title,
            h1=h1,
            body=body,
            date=date,
            page_kind=page_kind,
            canonical=canonical,
        )
        (ROOT / out_name).write_text(html, encoding="utf-8")
        print("wrote", out_name)

    # Product-specific pages → redirect to site-wide
    base = "https://office-goplan.com/"
    (ROOT / "pdfhandler-terms.html").write_text(
        REDIRECT.format(canonical=base + "terms-of-service"),
        encoding="utf-8",
    )
    (ROOT / "pdfhandler-privacy.html").write_text(
        REDIRECT.format(canonical=base + "privacy-policy"),
        encoding="utf-8",
    )
    print("wrote redirects pdfhandler-terms/privacy.html")


if __name__ == "__main__":
    main()
