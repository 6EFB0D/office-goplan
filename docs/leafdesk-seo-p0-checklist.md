# LeafDesk SEO / 索引チェックリスト（P0〜P1）

**対象**: https://office-goplan.com/leafdesk  
**更新**: 2026-08-24

## 実装済み（サイト側・P0）

- [x] `canonical` / meta description（旧称 pdfHandler 併記）
- [x] Open Graph / Twitter Card
- [x] JSON-LD `SoftwareApplication`（alternateName に旧称）
- [x] ヒーロー CTA（試用 DL / 価格 / 法人まとめ / 紹介チラシ）
- [x] `#enterprise` 法人・まとめ購入セクション
- [x] sitemap.xml に LeafDesk 関連 URL 掲載（`/leafdesk` `/leafdesk/guides/*` `/leafdesk-flyer`）
- [x] 旧 `/guides` → `/leafdesk/guides` へ 301（`_redirects`）
- [x] TOP 製品カードからチラシ・ガイドへの導線

## Search Console — 移行後手順（2026-08-24）

プロパティ: **https://office-goplan.com/**（ドメインプロパティ推奨）

### 1. サイトマップ（再送信）

1. Search Console → **サイトマップ**
2. 既存 `sitemap.xml` があっても **もう一度送信**（中身が `/leafdesk/guides` に更新済みのため）
   - 送信 URL: `https://office-goplan.com/sitemap.xml`
3. ステータスが「成功」になることを確認（数時間〜1日）

### 2. URL 検査 → インデックス登録をリクエスト

**URL 検査**で次を **上から順に** 実行（「インデックス登録をリクエスト」）。

| 優先 | URL |
|------|-----|
| 1 | `https://office-goplan.com/leafdesk` |
| 2 | `https://office-goplan.com/leafdesk/guides` |
| 3 | `https://office-goplan.com/leafdesk-flyer` |
| 4 | `https://office-goplan.com/leafdesk/guides/pdf-drawings-without-opening` |
| 5 | `https://office-goplan.com/leafdesk/guides/pdf-rename-while-preview` |
| 6 | `https://office-goplan.com/leafdesk/guides/pdf-page-replace-by-insert` |
| 7 | `https://office-goplan.com/leafdesk/guides/pdf-multiselect-pages` |
| 8 | `https://office-goplan.com/leafdesk/guides/pdf-merge-split-on-server` |
| 9 | `https://office-goplan.com/leafdesk/guides/pdf-header-footer` |

1 日あたりのリクエスト数に上限があるため、**4〜9 は翌日以降でも可**。

### 3. 旧 URL の扱い（心配しないもの）

| 現象 | 対応 |
|------|------|
| 旧 `/guides` や `/guides/…` が「**ページにリダイレクトがあります**」 | **正常**。301 で新 URL へ集約中。削除リクエストは不要 |
| 旧 `/pdf-drawings-without-opening` も同様 | 同上 |
| `.html` 付き URL（例: `/leafdesk.html`）がリダイレクト扱い | Cloudflare の 308。**正規は拡張子なし**（sitemap・canonical と一致） |
| 旧 GitHub Pages URL（`*.github.io/office-goplan/…`） | 別プロパティなら **アドレス変更** を検討。本番ドメインを正とする |

### 4. 1〜2 週間後の確認

- **検索パフォーマンス**: 「LeafDesk」「pdfHandler」「PDF 図面 サムネイル」など
- **ページのインデックス登録**: 新 URL が「インデックス登録済み」に増えているか
- **サイトマップ**: 送信済み URL の「検出」数が 10 前後（LeafDesk 関連 + 既存ページ）

### 5. やらなくてよいこと

- 旧 `/guides` の **削除リクエスト**（301 があるため不要）
- サイト全体の **再クロールの大量リクエスト**（上表の優先 URL のみで十分）

## 人手で実施（Search Console）— P0 完了

1. [x] Search Console でプロパティ確認
2. [x] sitemap.xml 送信／確認（**2026-08-24 再送信推奨** — `/leafdesk/guides` 移行後）
3. [x] `/leafdesk` インデックス登録リクエスト（登録済みでも再リクエスト可）
4. [ ] 1〜2 週間後、「LeafDesk」「pdfHandler」「PDF 図面 サムネイル」の表示回数を確認

## P1（進行中）

- [x] 製品ページ FAQ + `FAQPage` / `featureList`
- [x] 用途ガイドを `/leafdesk/guides/` 配下に統一（一覧 + 6 本）。旧 `/guides/`・`/pdf-drawings-without-opening` は 301
- [ ] Search Console: 上記 **移行後手順** の 2〜4 を実施
- [ ] 比較記事
- [ ] 導入事例

### 正規ガイド URL

| パス | タイトル |
|------|----------|
| `/leafdesk/guides` | 一覧 |
| `/leafdesk/guides/pdf-drawings-without-opening` | 図面PDFを開かずに見分ける |
| `/leafdesk/guides/pdf-rename-while-preview` | F2 リネーム |
| `/leafdesk/guides/pdf-page-replace-by-insert` | ページ挿入で差し替え |
| `/leafdesk/guides/pdf-multiselect-pages` | 複数ページ選択 |
| `/leafdesk/guides/pdf-merge-split-on-server` | 結合・分割 |
| `/leafdesk/guides/pdf-header-footer` | ヘッダ・フッター |

再生成: `python scripts/generate-leafdesk-guides.py`

## 確認コマンド（任意）

```powershell
(Invoke-WebRequest https://office-goplan.com/leafdesk -UseBasicParsing).Content |
  Select-String -Pattern 'og:title|SoftwareApplication|enterprise'
```
