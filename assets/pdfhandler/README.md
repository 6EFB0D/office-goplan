# pdfHandler / LeafDesk アセット

このフォルダには LeafDesk（旧 pdfHandler）製品ページ用の画像アセットを配置します。

## アイコン

- `PdfHandler.ico`: `pdf-handler/src/PdfHandler.UI/Assets/PdfHandler.ico` からコピーしたアプリ本体のアイコン。
- `PdfHandler.png`: Web表示用。`.ico` の外周につながる白背景を透明化したアイコン。

## スクリーンショット（製品ページ・ダミー PDF）

撮影元は `%USERPROFILE%\Documents\LeafDesk-HP-Demo\`。ツリーの実ドライブ等は `scripts/fix-leafdesk-hp-screenshots.py` でマスク済み。

- `gui-drawings.png`: 図面管理（サムネ＋ページストリップ）※旧 `gui-main.png` はキャッシュ残りのため参照しない
- `gui-po.png`: 顧客注文書（レイアウト差）
- `gui-mix.png`: 添付混在（図面／検査基準など）
- `gui-rfq.png`: 見積依頼・要求仕様
- `gui-multiselect.png`: ページストリップ複数選択（v1.3.10）。`scripts/make-leafdesk-multiselect-screenshot.py` で `gui-drawings.png` から生成

旧ダイアログ単体画像（`merge-dialog.png` 等）は差し替え前の控えとして残置可。
