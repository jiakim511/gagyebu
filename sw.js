/**
 * 가계부 앱 서비스 워커
 *
 * 목적은 "오프라인에서도 최소한 뜨게" + "업데이트를 최대한 빨리 반영"입니다.
 * 캐시를 먼저 믿는 방식(cache-first)이 아니라 반대로, 네트워크가 되면 항상 최신
 * 파일을 새로 받아오고, 네트워크가 안 될 때만 마지막으로 받아둔 캐시를 보여주는
 * "네트워크 우선(network-first)" 전략을 씁니다. 그래야 캐시가 "업데이트를 막는
 * 원인"이 되는 흔한 함정을 피할 수 있습니다.
 *
 * index.html/manifest.json 내용을 바꿀 때마다, 아래 CACHE_VERSION 숫자만 하나씩
 * 올려주세요. 그러면 이전 캐시가 자동으로 정리됩니다 (안 올려도 network-first라
 * 최신 내용은 온라인 상태에서 바로 반영되지만, 오프라인 폴백용 캐시를 깨끗하게
 * 정리하는 습관으로 올려두는 걸 추천합니다).
 */
var CACHE_VERSION = 'gagyebu-v1';
var ASSETS = ['./', './index.html', './manifest.json'];

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(CACHE_VERSION).then(function (cache) {
      return cache.addAll(ASSETS);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(
        keys.filter(function (k) { return k !== CACHE_VERSION; })
            .map(function (k) { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function (event) {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    fetch(event.request)
      .then(function (res) {
        var copy = res.clone();
        caches.open(CACHE_VERSION).then(function (cache) { cache.put(event.request, copy); });
        return res;
      })
      .catch(function () {
        return caches.match(event.request);
      })
  );
});
