# Office Go Plan

Office Go Plan の公式ウェブサイトです。

## プロジェクト構成

```
office-goplan/
├── index.html          # ホームページ
├── zipsearch.html      # ZipSearch 製品ページ
├── pictcomp.html      # PictComp 製品ページ
├── privacy-policy.html # プライバシーポリシー
├── styles.css          # 共通スタイル
├── .nojekyll           # Jekyll 無効化（GitHub Pages 用）
├── .gitignore
├── README.md
└── assets/
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
    └── pictcomp/       # PictComp 製品用アセット
        ├── pictcomp_bright.jpg  # 製品ロゴ
        ├── gui-main.svg         # スクリーンショット（PNG に差し替え）
        ├── gui-compressed.svg
        ├── gui-viewer.svg
        ├── point-1.svg～point-8.svg  # 推しポイント（PNG に差し替え）
        └── README.md
```

## GitHub Pages での公開方法

1. このリポジトリを GitHub にプッシュします
2. リポジトリの **Settings** → **Pages** を開きます
3. **Source** で「Deploy from a branch」を選択します
4. **Branch** で `main`（または `master`）を選択し、`/ (root)` を指定します
5. **Save** をクリックします

数分後、以下のURLでサイトが公開されます：
- `https://<ユーザー名>.github.io/office-goplan/`

## サポート

お問い合わせは support@office-gioplan.com までご連絡ください。
