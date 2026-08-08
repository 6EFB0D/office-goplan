# Office Go Plan

Office Go Plan の公式ウェブサイトです。

**文書・公開管理（社内）**: Quality Management System の **QP-DOC-001 §7**（Web／HP）。  
ページ ID・版は [`docs/SITE_DOC_REGISTER.md`](./docs/SITE_DOC_REGISTER.md)、公開要約は [`CHANGELOG.md`](./CHANGELOG.md)。

## プロジェクト構成

```
office-goplan/
├── index.html          # ホームページ
├── leafdesk.html       # LeafDesk 製品ページ（旧 pdfhandler.html）
├── pdfhandler.html     # /pdfhandler → /leafdesk リダイレクト
├── zipsearch.html      # ZipSearch 製品ページ
├── pictcomp.html      # PictComp 製品ページ
├── privacy-policy.html # プライバシーポリシー
├── terms-of-service.html # 利用規約
├── specified-commercial-transactions.html # 特定商取引法に基づく表記
├── styles.css          # 共通スタイル
├── .nojekyll           # Jekyll 無効化（GitHub Pages 用）
├── .gitignore
├── README.md
└── assets/
    ├── js/
    │   └── ga4.js      # Google アナリティクス GA4（計測 ID を設定）
    ├── logo/           # ブランドロゴ
    │   ├── logo-a.jpg  # ヘッダー用（暗色背景向けフィルター適用）
    │   ├── logo-b.jpg
    │   └── logo-c.jpg
    ├── zipsearch/      # ZipSearch 製品用アセット
    │   ├── zipsearch_blue.png
    │   ├── zipsearch_gray.png
    │   ├── gui-main.png
    │   ├── gui-results.png
    │   ├── web-main.png
    │   └── README.md
    ├── pictcomp/       # PictComp 製品用アセット
    │   ├── pictcomp_bright.jpg  # 製品ロゴ
    │   ├── pictcomp_trimmed.png # Web表示用トリム済みロゴ
    │   ├── gui-main.svg         # スクリーンショット（PNG に差し替え）
    │   ├── gui-compressed.svg
    │   ├── gui-viewer.svg
    │   ├── point-1.svg～point-8.svg  # 推しポイント（PNG に差し替え）
    │   └── README.md
    └── pdfhandler/     # pdfHandler 製品用アセット
        ├── PdfHandler.ico
        ├── PdfHandler.png
        └── README.md
```

## Google アナリティクス（GA4）

1. [Google アナリティクス](https://analytics.google.com/) でプロパティと Web データストリームを作成し、**計測 ID**（`G-` で始まる値）を取得します。
2. `assets/js/ga4.js` 内の `MEASUREMENT_ID` をその ID に置き換えて保存します（`XXXX` が残っているとタグは無効のままです）。
3. 変更をデプロイ後、[管理画面のレポート](https://analytics.google.com/)でリアルタイムなどにヒットが出るか確認します。

## 公開方法

本番は Cloudflare Pages（プロジェクト `office-goplan`）で、`main` への push で自動デプロイされます。

- 本番: `https://office-goplan.com/`
- 確認用: `https://office-goplan.pages.dev/`（Cloudflare Access でログイン必須）

ビルドは不要で、リポジトリのルートをそのまま配信します（Framework preset: None / Build output directory: `.`）。

### 旧 GitHub Pages

旧URL `https://<ユーザー名>.github.io/office-goplan/` は、`assets/js/legacy-redirect.js` により本番ドメインへ転送します。**転送を生かすため GitHub Pages の公開は止めないこと**。

### Search Console（インデックス）

Cloudflare Pages は `*.html` を拡張子なし URL へ **308** します。サイト内リンクと `sitemap.xml`・`rel=canonical` は拡張子なし（例: `/zipsearch`）に揃えています。

Search Console の「ページにリダイレクトがあります」は、旧 `.html` URL や GitHub Pages URL が **転送先を正とする除外**として出ることが多いです。次を確認してください。

1. プロパティが **https://office-goplan.com/**（ドメインプロパティ推奨）であること
2. [sitemap.xml](https://office-goplan.com/sitemap.xml) を送信済みであること
3. 除外理由の URL が `.html` や `github.io` なら、正規 URL（拡張子なし）が「インデックス登録済み」になっているか確認
4. 旧 GitHub Pages プロパティがある場合は、Search Console の「アドレス変更」で `office-goplan.com` へ移行

## サポート

お問い合わせは support@office-goplan.com までご連絡ください。
