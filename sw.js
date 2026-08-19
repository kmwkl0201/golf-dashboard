// 골프 라운딩 대시보드 서비스 워커
// - 정적 앱 셸(HTML/manifest/아이콘)만 캐시해서 오프라인에서도 앱이 열리게 한다.
// - 실제 데이터는 Google Sheets(Apps Script)에서 오므로 그 요청은 절대 캐시하지 않고
//   항상 네트워크로 직접 보낸다.

var CACHE_NAME = "golf-dashboard-v1";
var APP_SHELL = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png"
];

self.addEventListener("install", function (event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) {
      return cache.addAll(APP_SHELL);
    })
  );
  self.skipWaiting();
});

self.addEventListener("activate", function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(
        keys.filter(function (k) { return k !== CACHE_NAME; })
            .map(function (k) { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener("fetch", function (event) {
  var req = event.request;

  // GET이 아니거나(POST 등) Apps Script/Drive로 가는 요청은 그대로 네트워크로 통과시킨다.
  if (req.method !== "GET" || req.url.indexOf("script.google") !== -1 || req.url.indexOf("googleusercontent.com") !== -1) {
    return;
  }

  // 앱 셸: 캐시를 먼저 보여주고, 백그라운드에서 최신 버전을 받아 캐시를 갱신한다.
  event.respondWith(
    caches.match(req).then(function (cached) {
      var network = fetch(req).then(function (res) {
        if (res && res.status === 200) {
          var copy = res.clone();
          caches.open(CACHE_NAME).then(function (cache) { cache.put(req, copy); });
        }
        return res;
      }).catch(function () { return cached; });
      return cached || network;
    })
  );
});
