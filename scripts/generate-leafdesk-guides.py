# -*- coding: utf-8 -*-
"""Generate LeafDesk guide pages under /leafdesk/guides/ (run once from repo root)."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
GUIDE_URL = "/leafdesk/guides"
GUIDES = ROOT / "leafdesk" / "guides"

HEADER = """<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <link rel="icon" type="image/jpeg" href="../../assets/logo/logo-tab.jpg">
  <link rel="apple-touch-icon" sizes="180x180" href="../../assets/logo/apple-touch-icon-large.png">
  <link rel="manifest" href="../../site.webmanifest">
  <meta name="theme-color" content="#ffffff">
  <meta name="description" content="{description}">
  <title>{title}｜LeafDesk</title>
  <link rel="canonical" href="https://office-goplan.com/leafdesk/guides/{slug}">

  <meta property="og:type" content="article">
  <meta property="og:locale" content="ja_JP">
  <meta property="og:site_name" content="Office Go Plan">
  <meta property="og:url" content="https://office-goplan.com/leafdesk/guides/{slug}">
  <meta property="og:title" content="{title}">
  <meta property="og:description" content="{og_description}">
  <meta property="og:image" content="https://office-goplan.com/assets/pdfhandler/{og_image}">

  <meta name="twitter:card" content="summary_large_image">
  <meta name="twitter:title" content="{title}">
  <meta name="twitter:description" content="{og_description}">
  <meta name="twitter:image" content="https://office-goplan.com/assets/pdfhandler/{og_image}">

  <script type="application/ld+json">
  {{
    "@context": "https://schema.org",
    "@type": "Article",
    "headline": "{title}",
    "description": "{json_description}",
    "image": "https://office-goplan.com/assets/pdfhandler/{og_image}",
    "datePublished": "2026-08-23",
    "dateModified": "2026-08-23",
    "author": {{
      "@type": "Organization",
      "name": "Office Go Plan",
      "url": "https://office-goplan.com/"
    }},
    "publisher": {{
      "@type": "Organization",
      "name": "Office Go Plan",
      "url": "https://office-goplan.com/"
    }},
    "mainEntityOfPage": "https://office-goplan.com/leafdesk/guides/{slug}",
    "about": {{
      "@type": "SoftwareApplication",
      "name": "LeafDesk",
      "url": "https://office-goplan.com/leafdesk"
    }}
  }}
  </script>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;600;700&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="../../styles.css">
  <script src="../../assets/js/legacy-redirect.js"></script>
  <script defer src="../../assets/js/ga4.js"></script>
</head>
<body>
  <header class="header">
    <div class="container">
      <a href="/" class="logo">
        <img src="../../assets/logo/logo-a.jpg" alt="Office Go Plan" class="logo-img">
      </a>
      <nav class="nav">
        <a href="/">ホーム</a>
        <a href="/leafdesk">LeafDesk</a>
        <a href="/leafdesk/guides">ガイド</a>
        <a href="/#products">製品</a>
      </nav>
    </div>
  </header>

  <main>
    <article class="guide-article">
      <div class="container">
        <p class="guide-kicker"><a href="/leafdesk">LeafDesk</a> · <a href="/leafdesk/guides">用途ガイド</a></p>
        <h1>{title}</h1>
        <p class="guide-lead">{lead}</p>
"""

FOOTER = """
        <h2>試してみる</h2>
        <p>LeafDesk は初回起動から <strong>14日間</strong>全機能を試用できます。Standard 版は <strong>4,980円（税込・買い切り）</strong>です。</p>
        <p class="guide-cta-row">
          <a href="https://github.com/6EFB0D/pdf-handler/releases/latest" class="cta-button cta-button-primary">14日試用をはじめる（無料DL）</a>
          <a href="/leafdesk" class="cta-button cta-button-secondary">製品ページへ</a>
        </p>
        <p class="pro-note">ほかの用途: <a href="/leafdesk/guides">ガイド一覧</a> · <a href="/leafdesk#faq">FAQ</a> · <a href="/leafdesk#enterprise">法人・まとめ購入</a></p>
      </div>
    </article>
  </main>

  <footer class="footer">
    <div class="container">
      <nav class="footer-nav">
        <a href="/">ホーム</a>
        <a href="/leafdesk">LeafDesk</a>
        <a href="/leafdesk/guides">ガイド</a>
        <a href="/privacy-policy">プライバシーポリシー</a>
        <a href="/terms-of-service">利用規約</a>
        <a href="/specified-commercial-transactions">特定商取引法に基づく表記</a>
      </nav>
      <p class="copyright">&copy; Office Go Plan. All rights reserved.</p>
    </div>
  </footer>
</body>
</html>
"""

GUIDES_META = [
    {
        "slug": "pdf-drawings-without-opening",
        "title": "図面PDFを開かずに見分けるには",
        "description": "ファイルサーバ上の図面・注文書 PDF を、いちいち開かずにサムネイルで見分ける方法。Windows の限界と LeafDesk（旧 pdfHandler）での整理の考え方。",
        "og_description": "ファイルサーバ上の図面・注文書を、開かずにサムネイルで見分けて整理する考え方。",
        "json_description": "ファイルサーバ上の図面・注文書 PDF を開かずにサムネイルで見分ける方法と、Windows アプリ LeafDesk の使いどころ。",
        "og_image": "gui-drawings.png",
        "lead": "ファイルサーバに並んだ図面・注文書・検査資料。目的の1枚を探すたびに PDF を開いて閉じる作業は、思った以上に時間を食います。ここでは「開かずに見分ける」考え方と、Windows アプリ <strong>LeafDesk</strong>（旧称 pdfHandler）での整理の仕方をまとめます。",
        "body": """
        <figure class="guide-figure">
          <img src="../../assets/pdfhandler/gui-drawings.png" alt="LeafDesk で図面 PDF をサムネイル一覧とプレビューで確認している画面" width="1024" height="596">
          <figcaption>図面フォルダを一覧し、右側プレビューとページストリップで内容を確認（説明用のダミー PDF）</figcaption>
        </figure>

        <h2>よくある困りごと</h2>
        <ul class="guide-list">
          <li>ファイル名だけでは図番・改訂・添付種別が分からない</li>
          <li>エクスプローラーで1つずつ開くと、ビューアの起動待ちが続く</li>
          <li>ネットワークドライブだと、開くたびに待ち時間が積み重なる</li>
          <li>見つけたあと、そのままリネームや結合まで行きたい</li>
        </ul>

        <h2>Windows だけでもできること・限界</h2>
        <p>エクスプローラーの「大きいアイコン」やプレビューウィンドウは第一歩です。ただしフォルダ横断での「見ながら名前を変える」、複数ページの抜き出し・並べ替えまでは別ツールになりがちです。</p>

        <h2>開かずに見分ける、基本の流れ</h2>
        <ol class="guide-steps">
          <li><strong>フォルダを開く</strong> — 共有フォルダを左ペインで選ぶ</li>
          <li><strong>サムネイルで見比べる</strong> — 枠・図番・レイアウトの違いを把握する</li>
          <li><strong>必要ならプレビュー</strong> — 右側でページをめくり、ストリップで複数枚を確認する</li>
          <li><strong>そのまま整理</strong> — F2 リネーム、結合・分割、ページ操作へ</li>
        </ol>

        <h2>向いている場面</h2>
        <ul class="guide-list">
          <li>見た目で種別が分かる図面・注文書が多い</li>
          <li>ファイルサーバ上で探すことが多い</li>
          <li>見分けたあと、同じ画面で整理まで済ませたい</li>
        </ul>
""",
    },
    {
        "slug": "pdf-rename-while-preview",
        "title": "プレビューを見ながら F2 でリネームする",
        "description": "PDF を開き直さず、プレビュー表示のまま F2 でファイル名を変える方法。LeafDesk が元ファイルをロックしない仕組みと実務での使い方。",
        "og_description": "プレビューを見ながら F2 でリネーム。図面・注文書の名前付けをその場で。",
        "json_description": "LeafDesk で PDF プレビュー中に F2 リネームする方法。ファイルロックを避け、見分けと名付けを同時に行う。",
        "og_image": "gui-rename-f2.png",
        "lead": "図面や注文書は「見てから正しいファイル名にする」ことが多いです。いったん閉じてエクスプローラーで改名、また開く、という往復を減らすのがこのガイドの目的です。",
        "body": """
        <figure class="guide-figure">
          <img src="../../assets/pdfhandler/gui-rename-f2.png" alt="LeafDesk でプレビュー表示中に F2 でファイル名を編集している画面" width="1024" height="596">
          <figcaption>右側プレビューを確認しながら、一覧側で F2 編集（説明用のダミー PDF）</figcaption>
        </figure>

        <h2>なぜプレビュー中に改名しにくいか</h2>
        <p>多くのビューアは表示中の PDF をファイルロックします。その結果、「見た → 閉じた → エクスプローラーで改名 → また開く」が発生します。ネットワーク上だと待ち時間が特に長く感じられます。</p>

        <h2>LeafDesk での手順</h2>
        <ol class="guide-steps">
          <li>フォルダ内の PDF を選び、右側プレビューで内容を確認する</li>
          <li>中央一覧で対象を選択したまま <strong>F2</strong>（またはコンテキストメニューの名前変更）</li>
          <li>図番・客先・日付など、見た内容に合わせて名前を確定する</li>
          <li>続けて次のファイルへ。開き直しは不要</li>
        </ol>
        <p>LeafDesk は表示用にメモリへ読み込むため、<strong>元ファイルをロックしません</strong>。プレビュー中でもリネームできます（取説 FAQ Q1）。</p>

        <h2>実務でのコツ</h2>
        <ul class="guide-list">
          <li>命名規則（例: 図番_改訂_客先）を先に決めておく</li>
          <li>お気に入りフォルダに「今日の仕掛かり」を登録しておく</li>
          <li>改名後に結合・ページ整理へ進む場合も、同じ画面のまま続けられる</li>
        </ul>
""",
    },
    {
        "slug": "pdf-page-replace-by-insert",
        "title": "ページ挿入で PDF の差し替えをする",
        "description": "LeafDesk のページ挿入を使って、図面や添付の「差し替え」を行う方法。プレビューへのドロップや削除・挿入ダイアログの使い分け。",
        "og_description": "ページ挿入で差し替え。改訂図面や添付の入れ替えを LeafDesk で。",
        "json_description": "LeafDesk のページ挿入機能を、PDF ページの差し替え用途で使う手順。ドロップ挿入と削除・挿入ダイアログ。",
        "og_image": "gui-multiselect.png",
        "lead": "改訂図面や差し替え添付は、「古いページを除いて新しいページを入れる」作業です。LeafDesk ではページ<strong>挿入</strong>を軸に、実務では<strong>差し替え</strong>として使うケースが多くあります。",
        "body": """
        <h2>差し替え＝削除＋挿入、という考え方</h2>
        <p>PDF 全体を作り直さず、対象ページだけ入れ替える流れです。</p>
        <ol class="guide-steps">
          <li>差し替えたいページをプレビューまたはページストリップで特定する</li>
          <li>不要なページを<strong>削除</strong>する（または後ろに残して後で整理）</li>
          <li>新しい PDF／ページを<strong>挿入</strong>する</li>
          <li>順序を確認し、必要ならストリップで並べ替えて保存する</li>
        </ol>

        <h2>挿入の主な方法</h2>
        <ul class="guide-list">
          <li><strong>プレビューへ PDF をドロップ</strong> — 表示中ページの<strong>前</strong>に、ドロップした PDF のページが入ります（有償のページ編集）</li>
          <li><strong>削除・挿入ダイアログ</strong> — 全ページ／1ページ／ページ範囲を選んで挿入できます</li>
          <li><strong>ストリップでの複製・移動</strong> — 同じ PDF 内の並び替えや複製と組み合わせると、差し替え後の体裁を整えやすいです</li>
        </ul>

        <h2>向いている例</h2>
        <ul class="guide-list">
          <li>図面セットのうち、改訂された1枚だけを入れ替える</li>
          <li>注文書の添付ページを、最新版スキャンに差し替える</li>
          <li>表紙や仕様書ページだけ差し替え、残りはそのまま残す</li>
        </ul>
        <p class="pro-note">操作の詳細はアプリ内取扱説明書のページ編集・FAQ Q8（ドラッグ&ドロップ）も参照してください。</p>
""",
    },
    {
        "slug": "pdf-multiselect-pages",
        "title": "複数ページを選んで抜き出し・並べ替える",
        "description": "LeafDesk のページストリップで Ctrl／Shift 選択し、複製・並べ替え・別ファイルへのコピー／切り出しを行う方法。",
        "og_description": "ページストリップで複数選択。抜き出し・並べ替え・複製まで。",
        "json_description": "LeafDesk のページストリップで複数ページを選択し、並べ替え・複製・切り出しを行う使い方。",
        "og_image": "gui-multiselect.png",
        "lead": "長い PDF から必要なページだけ抜き出す、順序を組み替える、といった作業はストリップ上の<strong>複数選択</strong>が中心になります。",
        "body": """
        <figure class="guide-figure">
          <img src="../../assets/pdfhandler/gui-multiselect.png" alt="LeafDesk のページストリップで複数ページを選択した画面" width="1024" height="596">
          <figcaption>ページストリップで複数ページを選択（説明用のダミー PDF）</figcaption>
        </figure>

        <h2>基本操作</h2>
        <ol class="guide-steps">
          <li>PDF を開き、下部（またはプレビュー連動）の<strong>ページストリップ</strong>を表示する</li>
          <li><strong>Ctrl＋クリック</strong>で飛び飛び選択、<strong>Shift＋クリック</strong>で範囲選択</li>
          <li>複製・並べ替え・別ファイルへのコピー／切り出しなど、目的の操作を実行する</li>
          <li>並べ替え後は、必要に応じて保存手順に従う（ストリップの移動は「並べ替えを保存」が必要な場合があります）</li>
        </ol>

        <h2>こんなときに使う</h2>
        <ul class="guide-list">
          <li>検査成績書から該当ページだけを客先提出用に切り出す</li>
          <li>図面セットの順番を、現場の作業順に並べ替える</li>
          <li>同じページ構成を複製して、別案件用のたたき台を作る</li>
        </ul>
        <p>単ページの回転や、プレビュー上の削除・挿入と組み合わせると、一連のページ編集が同じ画面で完結しやすくなります。</p>
""",
    },
    {
        "slug": "pdf-merge-split-on-server",
        "title": "ファイルサーバ上の PDF を結合・分割する",
        "description": "共有フォルダ上の図面・注文書 PDF を、LeafDesk で結合・分割して整理する方法。保存先の指定とネットワーク利用時の注意。",
        "og_description": "ファイルサーバ上の PDF を結合・分割。図面セットの整理に。",
        "json_description": "ネットワークドライブを含むファイルサーバ上で、LeafDesk により PDF を結合・分割する手順と注意点。",
        "og_image": "merge-dialog.png",
        "lead": "客先提出用に複数図面を1つにまとめる、逆に巨大なセットを用途別に分ける——いずれもファイルサーバ上で完結できると手間が減ります。",
        "body": """
        <figure class="guide-figure">
          <img src="../../assets/pdfhandler/merge-dialog.png" alt="LeafDesk の PDF 結合ダイアログ" width="1024" height="596">
          <figcaption>結合ダイアログの例（画面は製品バージョンにより異なる場合があります）</figcaption>
        </figure>

        <h2>結合の流れ</h2>
        <ol class="guide-steps">
          <li>左ペインで対象フォルダを開く（ネットワークドライブも可）</li>
          <li>まとめたい PDF を選ぶ（順番は結合ダイアログ側で調整できる場合があります）</li>
          <li>結合を実行し、保存先を指定する（既定は元フォルダ付近が多いです）</li>
          <li>結果をプレビューで確認し、必要ならリネームする</li>
        </ol>

        <h2>分割の流れ</h2>
        <ol class="guide-steps">
          <li>分割したい PDF を開く</li>
          <li>分割ダイアログでページ範囲や方式を指定する</li>
          <li>保存先を確認して実行する（既定は元 PDF と同じフォルダになりやすいです）</li>
        </ol>

        <h2>ネットワーク利用時の注意</h2>
        <ul class="guide-list">
          <li>未接続のネットワークドライブは、エクスプローラーで一度開いてからアプリ側を「更新」すると安定しやすいです</li>
          <li>大きなファイルやページ数が多い PDF は時間がかかることがあります。品質・メモリ設定の調整も有効です</li>
          <li>処理中のキャンセルはできない場合があるため、範囲を確認してから実行してください</li>
        </ul>
""",
    },
    {
        "slug": "pdf-header-footer",
        "title": "図面・書類 PDF にヘッダ・フッターを付ける",
        "description": "LeafDesk で文書タイトルやページ番号などのヘッダ・フッターを付ける手順。オフセットやページ編集後の再適用の考え方。",
        "og_description": "図面・書類 PDF にヘッダ・フッター。ページ番号や文書タイトルを。",
        "json_description": "LeafDesk のヘッダ・フッター機能で、文書タイトルやページ番号を付与する手順と注意点。",
        "og_image": "header-footer.png",
        "lead": "提出用にページ番号や文書タイトルを揃えたいとき、PDF ごとに別ツールを開かず、整理と同じ画面でヘッダ・フッターを付けられます。",
        "body": """
        <figure class="guide-figure">
          <img src="../../assets/pdfhandler/header-footer.png" alt="LeafDesk のヘッダ・フッター設定画面の例" width="1024" height="596">
          <figcaption>ヘッダ・フッター設定の例（画面は製品バージョンにより異なる場合があります）</figcaption>
        </figure>

        <h2>手順の概要</h2>
        <ol class="guide-steps">
          <li>対象 PDF を選択する</li>
          <li>「ツール」→「ヘッダ・フッター」、またはツールバーから開く</li>
          <li>ヘッダー（文書タイトル・配置・上オフセット）を設定する</li>
          <li>フッター（任意文字列・ページ番号形式・下オフセット）を設定する</li>
          <li>フォントを必要に応じて調整し、適用する</li>
        </ol>

        <h2>うまく載せるコツ</h2>
        <ul class="guide-list">
          <li>本文と重なるときは、上端／下端からの<strong>オフセット (pt)</strong> を大きくする（最大 200）</li>
          <li>ページ番号は <code>1</code> / <code>1/10</code> / <code>-1-</code> など形式を選べます</li>
          <li>ページの削除・挿入のあと、設定によりヘッダ・フッターが<strong>自動再適用</strong>される場合があります</li>
        </ul>
        <p class="pro-note">適用エラー時は、最新版か・他アプリで同じ PDF を開いていないかを確認し、必要ならログ（%LOCALAPPDATA%\\LeafDesk\\logs）をサポートへ共有してください（取説 FAQ Q13）。</p>
""",
    },
]


def esc_json(s: str) -> str:
    return s.replace("\\", "\\\\").replace('"', '\\"')


def write_guide(g: dict) -> None:
    head = HEADER.format(
        slug=g["slug"],
        title=g["title"],
        description=g["description"],
        og_description=g["og_description"],
        json_description=esc_json(g["json_description"]),
        og_image=g["og_image"],
        lead=g["lead"],
    )
    # Fix double braces already in HEADER for JSON - we used {{ }} so format works
    path = GUIDES / f"{g['slug']}.html"
    path.write_text(head + g["body"] + FOOTER, encoding="utf-8", newline="\n")
    print("wrote", path.relative_to(ROOT))


INDEX = """<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <link rel="icon" type="image/jpeg" href="../../assets/logo/logo-tab.jpg">
  <link rel="apple-touch-icon" sizes="180x180" href="../../assets/logo/apple-touch-icon-large.png">
  <link rel="manifest" href="../../site.webmanifest">
  <meta name="theme-color" content="#ffffff">
  <meta name="description" content="LeafDesk（旧 pdfHandler）の用途ガイド一覧。図面 PDF の見分け、リネーム、ページ差し替え、結合・分割など。">
  <title>LeafDesk 用途ガイド｜Office Go Plan</title>
  <link rel="canonical" href="https://office-goplan.com/leafdesk/guides">
  <meta property="og:type" content="website">
  <meta property="og:locale" content="ja_JP">
  <meta property="og:site_name" content="Office Go Plan">
  <meta property="og:url" content="https://office-goplan.com/leafdesk/guides">
  <meta property="og:title" content="LeafDesk 用途ガイド">
  <meta property="og:description" content="図面・注文書 PDF の見分け、リネーム、差し替え、結合など用途別ガイド。">
  <meta property="og:image" content="https://office-goplan.com/assets/pdfhandler/gui-drawings.png">
  <script type="application/ld+json">
  {{
    "@context": "https://schema.org",
    "@type": "CollectionPage",
    "name": "LeafDesk 用途ガイド",
    "url": "https://office-goplan.com/leafdesk/guides",
    "isPartOf": {{
      "@type": "WebSite",
      "name": "Office Go Plan",
      "url": "https://office-goplan.com/"
    }},
    "about": {{
      "@type": "SoftwareApplication",
      "name": "LeafDesk",
      "url": "https://office-goplan.com/leafdesk"
    }}
  }}
  </script>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;600;700&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="../../styles.css">
  <script src="../../assets/js/legacy-redirect.js"></script>
  <script defer src="../../assets/js/ga4.js"></script>
</head>
<body>
  <header class="header">
    <div class="container">
      <a href="/" class="logo">
        <img src="../../assets/logo/logo-a.jpg" alt="Office Go Plan" class="logo-img">
      </a>
      <nav class="nav">
        <a href="/">ホーム</a>
        <a href="/leafdesk">LeafDesk</a>
        <a href="/leafdesk/guides">ガイド</a>
        <a href="/#products">製品</a>
      </nav>
    </div>
  </header>

  <main>
    <section class="guide-article">
      <div class="container">
        <p class="guide-kicker"><a href="/leafdesk">LeafDesk</a> · 用途ガイド</p>
        <h1>LeafDesk 用途ガイド</h1>
        <p class="guide-lead">ファイルサーバ上の図面・注文書 PDF を「開かずに見分け、そのまま整理する」ための短いガイドです。正規 URL はすべて <code>/leafdesk/guides/…</code> 配下に統一しています。</p>
        <ul class="guide-index-list">
{items}
        </ul>
        <p class="guide-cta-row">
          <a href="/leafdesk" class="cta-button cta-button-primary">製品ページへ</a>
          <a href="/leafdesk-flyer" class="cta-button cta-button-secondary">紹介チラシ（A4両面）</a>
          <a href="https://github.com/6EFB0D/pdf-handler/releases/latest" class="cta-button cta-button-secondary">14日試用をはじめる</a>
        </p>
      </div>
    </section>
  </main>

  <footer class="footer">
    <div class="container">
      <nav class="footer-nav">
        <a href="/">ホーム</a>
        <a href="/leafdesk">LeafDesk</a>
        <a href="/leafdesk/guides">ガイド</a>
        <a href="/privacy-policy">プライバシーポリシー</a>
        <a href="/terms-of-service">利用規約</a>
      </nav>
      <p class="copyright">&copy; Office Go Plan. All rights reserved.</p>
    </div>
  </footer>
</body>
</html>
"""


def main() -> None:
    GUIDES.mkdir(parents=True, exist_ok=True)
    for g in GUIDES_META:
        write_guide(g)
    items = "\n".join(
        f'          <li><a href="/leafdesk/guides/{g["slug"]}">{g["title"]}</a>'
        f'<span class="guide-index-tags">{" / ".join(tags_for(g["slug"]))}</span></li>'
        for g in GUIDES_META
    )
    (GUIDES / "index.html").write_text(
        INDEX.format(items=items), encoding="utf-8", newline="\n"
    )
    print("wrote leafdesk/guides/index.html")


def tags_for(slug: str) -> list[str]:
    return {
        "pdf-drawings-without-opening": ["サムネイル判別", "図面・注文書", "ファイルサーバ"],
        "pdf-rename-while-preview": ["F2リネーム", "プレビュー", "ロックしない"],
        "pdf-page-replace-by-insert": ["ページ挿入", "差し替え", "ドロップ"],
        "pdf-multiselect-pages": ["ページストリップ", "複数ページ選択", "並べ替え"],
        "pdf-merge-split-on-server": ["結合", "分割", "ネットワークドライブ"],
        "pdf-header-footer": ["ヘッダ・フッター", "ページ番号"],
    }.get(slug, [])


if __name__ == "__main__":
    main()
