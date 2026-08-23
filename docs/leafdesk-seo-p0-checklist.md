# LeafDesk SEO / 索引チェックリスト（P0〜P1）

**対象**: https://office-goplan.com/leafdesk  
**更新**: 2026-08-23

## 実装済み（サイト側・P0）

- [x] `canonical` / meta description（旧称 pdfHandler 併記）
- [x] Open Graph / Twitter Card
- [x] JSON-LD `SoftwareApplication`（alternateName に旧称）
- [x] ヒーロー CTA（試用 DL / 価格 / 法人まとめ）
- [x] `#enterprise` 法人・まとめ購入セクション
- [x] sitemap.xml に `/leafdesk` 掲載済み

## 人手で実施（Search Console）— P0 完了

1. [x] Search Console でプロパティ確認
2. [x] sitemap.xml 送信／確認
3. [x] `/leafdesk` インデックス登録リクエスト（登録済みでも再リクエスト済・2026-08-23）
4. [ ] 1〜2 週間後、「LeafDesk」「pdfHandler」「PDF 図面 サムネイル」の表示回数を確認

## P1（進行中）

- [x] 製品ページ FAQ + `FAQPage` / `featureList`
- [x] 用途ガイドを `/guides/` 配下に統一（一覧 + 6 本）。旧 `/pdf-drawings-without-opening` は 301
- [ ] Search Console: `/guides` および各ガイド URL の索引（**URL 確定後・任意のタイミング**）
- [ ] 比較記事
- [ ] 導入事例

### 正規ガイド URL

| パス | タイトル |
|------|----------|
| `/guides` | 一覧 |
| `/guides/pdf-drawings-without-opening` | 図面PDFを開かずに見分ける |
| `/guides/pdf-rename-while-preview` | F2 リネーム |
| `/guides/pdf-page-replace-by-insert` | ページ挿入で差し替え |
| `/guides/pdf-multiselect-pages` | 複数ページ選択 |
| `/guides/pdf-merge-split-on-server` | 結合・分割 |
| `/guides/pdf-header-footer` | ヘッダ・フッター |

再生成: `python scripts/generate-leafdesk-guides.py`

## 確認コマンド（任意）

```powershell
(Invoke-WebRequest https://office-goplan.com/leafdesk -UseBasicParsing).Content |
  Select-String -Pattern 'og:title|SoftwareApplication|enterprise'
```
