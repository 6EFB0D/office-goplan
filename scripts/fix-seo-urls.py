"""Normalize internal links and canonicals to Cloudflare Pages extensionless URLs."""
from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

CANONICALS = {
    "index.html": "https://office-goplan.com/",
    "zipsearch.html": "https://office-goplan.com/zipsearch",
    "leafdesk.html": "https://office-goplan.com/leafdesk",
    "pdfhandler.html": "https://office-goplan.com/leafdesk",
    "pictcomp.html": "https://office-goplan.com/pictcomp",
    "terms-of-service.html": "https://office-goplan.com/terms-of-service",
    "privacy-policy.html": "https://office-goplan.com/privacy-policy",
    "specified-commercial-transactions.html": "https://office-goplan.com/specified-commercial-transactions",
    "success.html": "https://office-goplan.com/success",
    "cancel.html": "https://office-goplan.com/cancel",
    "pdfhandler-terms.html": "https://office-goplan.com/terms-of-service",
    "pdfhandler-privacy.html": "https://office-goplan.com/privacy-policy",
}

HREF_MAP = [
    ("index.html#products", "/#products"),
    ("index.html", "/"),
    ("specified-commercial-transactions.html", "/specified-commercial-transactions"),
    ("terms-of-service.html", "/terms-of-service"),
    ("privacy-policy.html", "/privacy-policy"),
    ("leafdesk.html", "/leafdesk"),
    ("pdfhandler.html", "/leafdesk"),
    ("zipsearch.html", "/zipsearch"),
    ("pictcomp.html", "/pictcomp"),
]


def main() -> None:
    for path in sorted(ROOT.glob("*.html")):
        text = path.read_text(encoding="utf-8")
        orig = text

        for old, new in HREF_MAP:
            text = text.replace(f'href="{old}"', f'href="{new}"')

        text = re.sub(
            r"(https://office-goplan\.com/)([a-z0-9\-]+)\.html",
            r"\1\2",
            text,
        )

        name = path.name
        if name in CANONICALS:
            canon = CANONICALS[name]
            if 'rel="canonical"' in text:
                text = re.sub(
                    r'<link rel="canonical" href="[^"]+">',
                    f'<link rel="canonical" href="{canon}">',
                    text,
                    count=1,
                )
            else:
                text = re.sub(
                    r"(<title>[^<]*</title>\s*)",
                    rf'\1  <link rel="canonical" href="{canon}">\n',
                    text,
                    count=1,
                )

        if text != orig:
            path.write_text(text, encoding="utf-8", newline="\n")
            print(f"updated {path.name}")
        else:
            print(f"unchanged {path.name}")


if __name__ == "__main__":
    main()
