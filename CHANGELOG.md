# CHANGELOG — Office Go Plan サイト（office-goplan）

## 2026-08-08（LeafDesk URL `/leafdesk`）

- **WEB-PDFH** — 正規 URL を `/pdfhandler` から **`/leafdesk`** へ変更（`leafdesk.html`）
- 旧 `/pdfhandler` は Cloudflare `_redirects`（301）と `pdfhandler.html` スタブで転送
- `sitemap.xml` / トップカード / canonical を同期

## 2026-08-03（LeafDesk 商品ページ — ユースケース訴求）

- **WEB-PDFH** — ヒーロー／概要を「開かずに見分ける」訴求に更新。「こんな場面で」を追加（図面・注文書・添付・見積）
- ◎〇△の効果判例表示をやめ、見出し＋説明のみに整理
- スクリーンショットをダミー図面／注文書／添付混在／見積依頼の 4 枚に差し替え（ツリーの実ドライブはマスク済み）
- 図面スクショを `gui-drawings.png` に改名して参照（同名 `gui-main.png` のブラウザ／CDN キャッシュ回避）
- **WEB-INDEX** — LeafDesk カード説明文を同趣旨に同期

## 2026-08-02（LeafDesk HP 本番反映・v1.3.9）

- **WEB-PDFH** — 未 push だった LeafDesk 表記（H1／副題／法務）を `main` へ push → Cloudflare Pages 本番デプロイ
- ZIP 案内を `LeafDesk-*-prod-setup.zip` に更新（旧 `PdfHandler-*` も注記）
- スクリーンショット alt を LeafDesk 表記に更新

## 2026-08-01（法務 URL 一本化・LeafDesk 表記）

- **法務正本** — `docs/legal/` 最終更新日を 2026-08-01 に更新。附則に「office-goplan.com 一本化／旧 GitHub Pages 移行／LeafDesk 表記」を記録
- **HP** — `terms-of-service.html` / `privacy-policy.html` を sync 再生成。トップ・製品ページの表示名を **LeafDesk**（副: pdf Handler for Windows）に更新
- **公開 URL 正本**: https://office-goplan.com/terms-of-service ／ https://office-goplan.com/privacy-policy

## 2026-07-17（pdfHandler DL 案内 — 版番号非記載）

- **WEB-PDFH** / トップ — 製品ページから固定版番号・「主な更新」を削除。ダウンロードは常に `/releases/latest` へ誘導（リリースごとの HP メンテを不要化）

## 2026-07-12（pdfHandler v1.2.7）

- **WEB-PDFH** / トップ — 最新版表示を **v1.2.7** に更新（プレビューツールバー見切れ修正）

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
