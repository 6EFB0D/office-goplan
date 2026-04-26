# Office Go Plan

Office Go Plan の公式ウェブサイトです。

## プロジェクト構成

```
office-goplan/
├── index.html          # ホームページ
├── pdfhandler.html     # pdfHandler 製品ページ
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

## GitHub Pages での公開方法

1. このリポジトリを GitHub にプッシュします
2. リポジトリの **Settings** → **Pages** を開きます
3. **Source** で「Deploy from a branch」を選択します
4. **Branch** で `main`（または `master`）を選択し、`/ (root)` を指定します
5. **Save** をクリックします

数分後、以下のURLでサイトが公開されます：
- `https://<ユーザー名>.github.io/office-goplan/`

## サポート

お問い合わせは support@office-goplan.com までご連絡ください。
