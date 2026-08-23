# LeafDesk SEO / 索引チェックリスト（P0）

**対象**: https://office-goplan.com/leafdesk  
**更新**: 2026-08-23

## 実装済み（サイト側）

- [x] `canonical` / meta description（旧称 pdfHandler 併記）
- [x] Open Graph / Twitter Card
- [x] JSON-LD `SoftwareApplication`（alternateName に旧称）
- [x] ヒーロー CTA（試用 DL / 価格 / 法人まとめ）
- [x] `#enterprise` 法人・まとめ購入セクション
- [x] sitemap.xml に `/leafdesk` 掲載済み

## 人手で実施（Search Console）

1. [Google Search Console](https://search.google.com/search-console) でプロパティ `https://office-goplan.com/` を開く
2. **sitemap.xml** を送信（未送信なら）または再読み込み
3. URL 検査で `https://office-goplan.com/leafdesk` を「インデックス登録をリクエスト」
4. 1〜2 週間後、「LeafDesk」「pdfHandler」「PDF 図面 サムネイル」の表示回数を確認

## 確認コマンド（任意）

```powershell
# OGP / JSON-LD が配信されているか
(Invoke-WebRequest https://office-goplan.com/leafdesk -UseBasicParsing).Content |
  Select-String -Pattern 'og:title|SoftwareApplication|enterprise'
```
