/**
 * Google Analytics 4 (gtag.js)
 *
 * 手順:
 * 1. https://analytics.google.com/ でプロパティを作成し、データストリームで Web を追加
 * 2. 「計測 ID」（G- で始まる文字列）をコピー
 * 3. 下の MEASUREMENT_ID だけ実際の ID に置き換えて保存
 *
 * プレースホルダ（XXXX を含む）のままだとスクリプトは実行されません。
 */
(function () {
  var MEASUREMENT_ID = 'G-PE1TCKLF9V';

  if (!MEASUREMENT_ID || /XXXX/i.test(MEASUREMENT_ID)) {
    return;
  }

  window.dataLayer = window.dataLayer || [];
  function gtag() {
    window.dataLayer.push(arguments);
  }
  window.gtag = gtag;
  gtag('js', new Date());

  var script = document.createElement('script');
  script.async = true;
  script.src = 'https://www.googletagmanager.com/gtag/js?id=' + MEASUREMENT_ID;
  document.head.appendChild(script);

  gtag('config', MEASUREMENT_ID);
})();
