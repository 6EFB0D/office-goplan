# Office Go Plan - 書類管理アプリ

見積書・請求書・領収書の発行、採番、取引履歴管理を行うローカル GUI アプリです。

## 機能

- 顧客情報・品目・数量・単価を入力して取引を登録
- 見積書 / 請求書 / 領収書を Word + PDF で自動生成・採番
- 取引ステータス（新規 → 見積発行済 → 請求発行済 → 完了）の一覧管理
- Excel (`transactions.xlsx`) に全取引・書類番号を自動記録

## 起動方法

```powershell
cd D:\Users\admin_mak\project\office-goplan\docs\invoice-app
python app.py
```

## 初回セットアップ

```powershell
pip install -r requirements.txt
```

## ファイル構成

```
invoice-app/
├── app.py               # メイン GUI
├── config.py            # 事業者情報・パス定数
├── db.py                # Excel DB 操作
├── docgen.py            # Word / PDF 生成
├── requirements.txt     # 依存パッケージ
├── transactions.xlsx    # 取引管理 Excel（自動生成、Git 管理外）
└── output/              # 生成 PDF/Word（Git 管理外）
    ├── quotes/
    ├── invoices/
    └── receipts/
```

## Excel の構成

| シート | 内容 |
|--------|------|
| 取引管理 | 全取引・書類番号・ステータス |
| 採番    | Q / I / R の年度別最終連番 |

## 書類番号の体系

| 種別 | 接頭辞 | 例 |
|------|--------|-----|
| 見積書 | Q- | Q-20260425-0001 |
| 請求書 | I- | I-20260425-0001 |
| 領収書 | R- | R-20260425-0001 |

## 見積変更が生じた場合

1. 該当取引を選択し「見積書を再発行」ボタンを押す
2. 新しい見積番号が採番され、旧番号は Excel の備考欄に手動メモ

## PDF 変換について

Word (Microsoft Word) がインストールされている環境では、.docx が自動的に .pdf に変換されます。
Word がない場合は .docx のみ出力されます。
