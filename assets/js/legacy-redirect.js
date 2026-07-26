/**
 * GitHub Pages (旧URL) から本番ドメインへの転送。
 * Cloudflare Pages / office-goplan.com / pages.dev では何もしない。
 */
(function () {
  var host = location.hostname;
  if (host !== '6efb0d.github.io') return;

  var prefix = '/office-goplan';
  var path = location.pathname;
  if (path.indexOf(prefix) === 0) {
    path = path.slice(prefix.length) || '/';
  }

  location.replace(
    'https://office-goplan.com' + path + location.search + location.hash
  );
})();
