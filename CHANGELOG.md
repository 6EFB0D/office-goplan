# CHANGELOG — Office Go Plan サイト（office-goplan）

## 2026-07-12（製品順・リンク色統一）

- 利用規約・PP の製品例示順を **pdfHandler → ZipSearch → PictComp** に変更
- 青系テキスト色を `--color-accent` / `--color-link` でページ通し統一。ボタン背景は `--color-cta` に分離

## 2026-07-12（対外表現・リンク色）

- 法務ページ先頭の社内向け文言（`docs/legal/`・編集正本など）を削除し、お客様向けの説明に変更
- リンク色を明るいシアン系（`--color-link`）に変更。`a:visited` の紫潰れを防止（本文・pro-note・フッター等）
- 製品ページ注記・アプリ埋め込みの「正本／実質同一」表現を平易化

## 2026-07-12（追記）

- **WEB-TERMS / WEB-PP** — 全製品共通の利用規約・プライバシーに一本化。用語 **処理対象データ**（アプリが扱うファイル等）を定義
- 編集正本を `docs/legal/TERMS_OF_SERVICE.txt` / `PRIVACY_POLICY.txt` に設置。生成: `scripts/sync-site-legal.py`
- **WEB-PDFH-TERMS / WEB-PDFH-PP** — サイト全体ページへリダイレクトに変更
- **WEB-PDFH** — フッター・注記を共通法務 URL に戻す

## 2026-07-12（午前）

- 製品別法務ページ新設・文書管理台帳開設（その後、一本化方針で上記に置換）

## 以前の変更

Git 履歴を技術的な正とする。
