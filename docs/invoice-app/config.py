# -*- coding: utf-8 -*-
"""アプリ全体で使用する定数。事業者情報を変更する場合はこのファイルを編集する。"""
import pathlib

_APP_DIR = pathlib.Path(__file__).parent
_PROJECT_ROOT = _APP_DIR.parent.parent

# ---- 事業者情報 ----
SELLER_NAME    = "Office Go Plan"
SELLER_POSTAL  = "〒106-0032"
SELLER_ADDRESS = "東京都港区六本木2-1-19 S-Building 3F"
SELLER_EMAIL   = "support@office-goplan.com"
SELLER_WEB     = "https://office-goplan.com"

# ---- 銀行口座 ----
BANK_NAME           = "GMOあおぞらネット銀行"
BANK_BRANCH         = "にじ支店（支店番号 302）"
BANK_ACCOUNT_TYPE   = "普通"
BANK_ACCOUNT_NUMBER = "3189329"
BANK_ACCOUNT_HOLDER = "カワギシ　マコト"

# ---- パス ----
LOGO_PATH     = str(_PROJECT_ROOT / "assets" / "logo" / "logo-a.jpg")
LOGO_RENDERED = str(_APP_DIR / "_logo_rendered.png")
EXCEL_PATH    = str(_APP_DIR / "transactions.xlsx")
OUTPUT_DIR    = str(_APP_DIR / "output")

# ---- デフォルト商品 ----
PRODUCT_DEFAULT    = "PDF Handler 買い切り版ライセンス"
UNIT_PRICE_DEFAULT = 5000

# ---- 定型文 ----
NOTE_TAX = (
    "当サービスの運営者は免税事業者であり、適格請求書発行事業者の登録番号（T番号）を有していません。"
    "表示金額は「不課税（免税事業者）」として税込一本価格で表示し、消費税額は別建てで請求しておりません。"
    "仕入税額控除の経過措置（〜2026年9月30日は80%、2026年10月1日〜2029年9月30日は50%）をご参照ください。"
)
NOTE_LICENSE = (
    "ライセンスキーは、決済承認後または入金確認後に、"
    "ご注文時のメールアドレス宛にお送りします。"
)
NOTE_TERMS_B2B = (
    "法人ユーザーの後払い取引の条件（催告・一時停止・失効・遅延損害金 年14.6% 等）は、"
    "利用規約 第5条の2 に定めます（https://office-goplan.com/terms-of-service.html）。"
)
