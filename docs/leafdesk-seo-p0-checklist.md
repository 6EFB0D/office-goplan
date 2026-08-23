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
- [x] 用途記事: https://office-goplan.com/pdf-drawings-without-opening （sitemap 掲載）
- [ ] 比較記事
- [ ] 導入事例

## 確認コマンド（任意）

```powershell
# OGP / JSON-LD が配信されているか
(Invoke-WebRequest https://office-goplan.com/leafdesk -UseBasicParsing).Content |
  Select-String -Pattern 'og:title|SoftwareApplication|enterprise'
```
