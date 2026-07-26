# サイト文書台帳（SITE_DOC_REGISTER）

| 項目 | 内容 |
|---|---|
| **文書番号** | SITE_DOC_REGISTER |
| **種別** | 動的台帳 |
| **上位手順** | QP-DOC-001 §7 / `terms-embedded-vs-website.md` v3 |
| **サイト** | https://office-goplan.com/ |
| **最終更新** | 2026-07-17 |

---

## 登録一覧

| WEB ID | 名称 | 主ファイル | 版／日付 | 備考 |
|---|---|---|---|---|
| WEB-INDEX | トップ | `index.html` | — | |
| WEB-PDFH | pdfHandler | `pdfhandler.html` | 2026-07-17 | 版番号非記載。DL は `/releases/latest`。Assets 案内。法務はサイト全体へリンク |
| WEB-ZIP | ZipSearch | `zipsearch.html` | — | |
| WEB-PICT | PictComp | `pictcomp.html` | — | |
| **WEB-TERMS** | **利用規約（全製品）** | `terms-of-service.html` | **2026-07-12** | 正本 `docs/legal/TERMS_OF_SERVICE.txt`。**Pages 反映** commit `d9da018` |
| **WEB-PP** | **プライバシー（全製品）** | `privacy-policy.html` | **2026-07-12** | 同上 |
| WEB-PDFH-TERMS | （廃止・リダイレクト） | `pdfhandler-terms.html` | → WEB-TERMS | |
| WEB-PDFH-PP | （廃止・リダイレクト） | `pdfhandler-privacy.html` | → WEB-PP | |
| WEB-TOKUTEI | 特商法 | `specified-commercial-transactions.html` | — | |
| WEB-CHECKOUT | success/cancel | — | — | |

再生成: `python scripts/sync-site-legal.py`

---

## 台帳改訂

| 日付 | 内容 |
|---|---|
| 2026-07-17 | **WEB-PDFH** / トップ — 固定版番号・主な更新を廃止（`/releases/latest` 誘導） |
| 2026-07-12 | **WEB-PDFH** / トップ — 最新版 **v1.2.7**（ツールバー見切れ修正） |
| 2026-07-12 | 法務をサイト全体2本に一本化（処理対象データ）。製品別はリダイレクト |
