# サイト文書台帳（SITE_DOC_REGISTER）

| 項目 | 内容 |
|---|---|
| **文書番号** | SITE_DOC_REGISTER |
| **種別** | 動的台帳 |
| **上位手順** | QP-DOC-001 §7 / `terms-embedded-vs-website.md` v3 |
| **サイト** | https://office-goplan.com/ |
| **最終更新** | 2026-11-14 |

---

## 登録一覧

| WEB ID | 名称 | 主ファイル | 版／日付 | 備考 |
|---|---|---|---|---|
| WEB-INDEX | トップ | `index.html` | — | |
| WEB-PDFH | LeafDesk（旧 pdfHandler） | `leafdesk.html` | **2026-11-14** | 正規 URL `/leafdesk`。ヒーロー対比は紹介チラシと揃え。「画面で見る」に YT-01／YT-02（YouTube）と注文書・N-up のローカル動画 |
| WEB-GUIDES | LeafDesk 用途ガイド一覧 | `leafdesk/guides/index.html` | **2026-08-23** | `/leafdesk/guides` |
| WEB-PDFH-G1 | 図面を開かずに見分ける | `leafdesk/guides/pdf-drawings-without-opening.html` | **2026-09-13** | 旧 `/guides/…`・`/pdf-drawings-without-opening` は 301。お気に入り／フィルタ追記 |
| WEB-PDFH-G2 | F2 リネーム | `leafdesk/guides/pdf-rename-while-preview.html` | **2026-09-13** | フィルタ用途例を追記 |
| WEB-PDFH-G3 | ページ挿入で差し替え | `leafdesk/guides/pdf-page-replace-by-insert.html` | **2026-08-23** | |
| WEB-PDFH-G4 | 複数ページ選択 | `leafdesk/guides/pdf-multiselect-pages.html` | **2026-08-23** | |
| WEB-PDFH-G5 | 結合・分割 | `leafdesk/guides/pdf-merge-split-on-server.html` | **2026-08-23** | |
| WEB-PDFH-G6 | ヘッダ・フッター | `leafdesk/guides/pdf-header-footer.html` | **2026-08-23** | |
| WEB-PDFH-FLYER | 紹介チラシ A4両面 | `leafdesk-flyer.html` | **2026-11-13** | `/leafdesk-flyer` 印刷用。裏面の見出し・締めを更新 |
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
| 2026-11-14 | **WEB-PDFH** — 「画面で見る」に YT-02（回転・拡大縮小表示）を追加 |
| 2026-11-13 | **WEB-PDFH** — ヒーロー対比の見出し・締めを紹介チラシと揃え |
| 2026-11-13 | **WEB-PDFH-FLYER** — 裏面の見出し・締めを更新。「5 工程／5 機能」を削除 |
| 2026-11-08 | **WEB-PDFH** — 「画面で見る」に N-up 面付けの解説動画 |
| 2026-11-08 | **WEB-PDFH** — 15秒動画を外し左右対比を掲載。試用ボタンを3ボタンの下へ。顧客注文書に解説動画 |
| 2026-09-13 | **WEB-PDFH / G1 / G2** — お気に入り短縮・ファイル名フィルタの用途例（IMP-guide-fav-filter-1） |
| 2026-08-23 | **WEB-PDFH-FLYER** — A4両面紹介チラシ `/leafdesk-flyer` |
| 2026-08-24 | **WEB-GUIDES / G1〜G6** — `/leafdesk/guides/` へ移行。旧 `/guides/` は 301 |
| 2026-08-23 | **WEB-GUIDES / G1〜G6** — `/guides/` 配下に用途ガイド一括。旧フラット URL は 301 |
| 2026-08-23 | **WEB-PDFH-G1** — 用途ガイド「図面PDFを開かずに見分けるには」追加（P1） |
| 2026-08-23 | **WEB-PDFH** — F2 リネーム／複数選択スクショ追加・機能文言更新（v1.3.10） |
| 2026-08-03 | **WEB-PDFH** — ユースケース（図面・注文書等）＋ダミー PDF スクショ差し替え |
| 2026-07-17 | **WEB-PDFH** / トップ — 固定版番号・主な更新を廃止（`/releases/latest` 誘導） |
| 2026-07-12 | **WEB-PDFH** / トップ — 最新版 **v1.2.7**（ツールバー見切れ修正） |
| 2026-07-12 | 法務をサイト全体2本に一本化（処理対象データ）。製品別はリダイレクト |
