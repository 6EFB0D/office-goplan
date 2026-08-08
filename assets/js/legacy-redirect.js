/**
 * GitHub Pages (旧URL) から本番ドメインへの転送。
 * Cloudflare Pages / office-goplan.com / pages.dev では何もしない。
 */
(function () {
  var host = location.hostname;
  if (host !== '6efb0d.github.io') return;

  // 旧ホストのインデックスを避け、正規 URL へ寄せる
  var robots = document.createElement('meta');
  robots.name = 'robots';
  robots.content = 'noindex, follow';
  if (document.head) {
    document.head.appendChild(robots);
  }

  var prefix = '/office-goplan';
  var path = location.pathname;
  if (path.indexOf(prefix) === 0) {
    path = path.slice(prefix.length) || '/';
  }
  // Cloudflare Pages は *.html を拡張子なしへ 308 するため揃える
  if (path.length > 1 && path.slice(-5) === '.html') {
    path = path.slice(0, -5);
  }
  if (path === '/index') {
    path = '/';
  }

  location.replace(
    'https://office-goplan.com' + path + location.search + location.hash
  );
})();
