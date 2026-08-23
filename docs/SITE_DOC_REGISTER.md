# サイト文書台帳（SITE_DOC_REGISTER）

| 項目 | 内容 |
|---|---|
| **文書番号** | SITE_DOC_REGISTER |
| **種別** | 動的台帳 |
| **上位手順** | QP-DOC-001 §7 / `terms-embedded-vs-website.md` v3 |
| **サイト** | https://office-goplan.com/ |
| **最終更新** | 2026-08-23 |

---

## 登録一覧

| WEB ID | 名称 | 主ファイル | 版／日付 | 備考 |
|---|---|---|---|---|
| WEB-INDEX | トップ | `index.html` | — | |
| WEB-PDFH | LeafDesk（旧 pdfHandler） | `leafdesk.html` | **2026-08-23** | 正規 URL `/leafdesk`。FAQ・`/guides` 導線 |
| WEB-GUIDES | LeafDesk 用途ガイド一覧 | `guides/index.html` | **2026-08-23** | `/guides` |
| WEB-PDFH-G1 | 図面を開かずに見分ける | `guides/pdf-drawings-without-opening.html` | **2026-08-23** | 旧 `/pdf-drawings-without-opening` は 301 |
| WEB-PDFH-G2 | F2 リネーム | `guides/pdf-rename-while-preview.html` | **2026-08-23** | |
| WEB-PDFH-G3 | ページ挿入で差し替え | `guides/pdf-page-replace-by-insert.html` | **2026-08-23** | |
| WEB-PDFH-G4 | 複数ページ選択 | `guides/pdf-multiselect-pages.html` | **2026-08-23** | |
| WEB-PDFH-G5 | 結合・分割 | `guides/pdf-merge-split-on-server.html` | **2026-08-23** | |
| WEB-PDFH-G6 | ヘッダ・フッター | `guides/pdf-header-footer.html` | **2026-08-23** | |
| WEB-PDFH-FLYER | 紹介チラシ A4両面 | `leafdesk-flyer.html` | **2026-08-23** | `/leafdesk-flyer` 印刷用 |
| WEB-ZIP | ZipSearch | `zipsearch.html` | — | |
| WEB-PICT | PictComp | `pictcomp.html` | — | |
| **WEB-TERMS** | **利用規約（全製品）** | `terms-of-service.html` | **2026-08-01** | 正本 `docs/legal/TERMS_OF_SERVICE.txt`（LeafDesk 表記） |
| **WEB-PP** | **プライバシー（全製品）** | `privacy-policy.html` | **2026-08-01** | 同上 |
| WEB-PDFH-TERMS | （廃止・リダイレクト） | `pdfhandler-terms.html` | → WEB-TERMS | |
| WEB-PDFH-PP | （廃止・リダイレクト） | `pdfhandler-privacy.html` | → WEB-PP | |
| WEB-TOKUTEI | 特商法 | `specified-commercial-transactions.html` | — | |
| WEB-CHECKOUT | success/cancel | — | — | |

再生成: `python scripts/sync-site-legal.py`

---

## 台帳改訂

| 日付 | 内容 |
|---|---|
| 2026-08-23 | **WEB-PDFH-FLYER** — A4両面紹介チラシ `/leafdesk-flyer` |
| 2026-08-23 | **WEB-GUIDES / G1〜G6** — `/guides/` 配下に用途ガイド一括。旧フラット URL は 301 |
| 2026-08-23 | **WEB-PDFH-G1** — 用途ガイド「図面PDFを開かずに見分けるには」追加（P1） |
| 2026-08-23 | **WEB-PDFH** — F2 リネーム／複数選択スクショ追加・機能文言更新（v1.3.10） |
| 2026-08-03 | **WEB-PDFH** — ユースケース（図面・注文書等）＋ダミー PDF スクショ差し替え |
| 2026-07-17 | **WEB-PDFH** / トップ — 固定版番号・主な更新を廃止（`/releases/latest` 誘導） |
| 2026-07-12 | **WEB-PDFH** / トップ — 最新版 **v1.2.7**（ツールバー見切れ修正） |
| 2026-07-12 | 法務をサイト全体2本に一本化（処理対象データ）。製品別はリダイレクト |
