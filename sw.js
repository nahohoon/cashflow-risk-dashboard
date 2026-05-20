/* ════════════════════════════════════════════════════════
 *  현금흐름 위험관리판 — Service Worker
 *
 *  전략:
 *   - HTML navigation: network-first (온라인에서 최신 shell)
 *   - 그 외 동일 출처 자산: cache-first
 *   - API (script.google.com): 네트워크만
 *   - 기타 정적 자원 (폰트 등): stale-while-revalidate
 * ════════════════════════════════════════════════════════ */

const VERSION = 'cashflow-risk-v20260520-5';
const CACHE_SHELL = `cashflow-risk-shell-${VERSION}`;
const CACHE_RUNTIME = `cashflow-risk-runtime-${VERSION}`;

const SHELL_ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon.svg'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_SHELL).then(cache => cache.addAll(SHELL_ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(
      // 현재 버전이 아닌 캐시는 전부 삭제 (버전명 무관하게)
      keys
        .filter(k => k !== CACHE_SHELL && k !== CACHE_RUNTIME)
        .map(k => {
          console.log('[SW] deleting old cache:', k);
          return caches.delete(k);
        })
    )).then(() => {
      console.log('[SW] activated version:', VERSION);
      return self.clients.claim();
    })
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);

  // ─ API 호출: Apps Script 도메인은 항상 네트워크, 캐시 금지
  if (url.hostname.endsWith('script.google.com') ||
      url.hostname.endsWith('googleusercontent.com')) {
    event.respondWith(
      fetch(req, { cache: 'no-store' }).catch(() => new Response(
        JSON.stringify({ error: 'offline' }),
        { status: 503, headers: { 'Content-Type': 'application/json' } }
      ))
    );
    return;
  }

  // ─ 동일 출처: HTML 문서(navigation) / index.html → network-first
  if (url.origin === self.location.origin) {
    const isNavigation =
      req.mode === 'navigate' ||
      req.destination === 'document' ||
      url.pathname.endsWith('/') ||
      url.pathname.endsWith('index.html');

    if (isNavigation) {
      event.respondWith(
        fetch(req, { cache: 'no-store' }).then(res => {
          // 성공하면 최신 버전으로 캐시 갱신
          if (res.ok && res.status === 200) {
            const clone = res.clone();
            caches.open(CACHE_SHELL).then(c => c.put(req, clone));
          }
          return res;
        }).catch(() => {
          // 오프라인 fallback
          return caches.match('./index.html');
        })
      );
      return;
    }

    // 그 외 동일 출처 자산: cache-first
    event.respondWith(
      caches.match(req).then(cached => {
        if (cached) return cached;
        return fetch(req).then(res => {
          if (res.ok && res.status === 200) {
            const clone = res.clone();
            caches.open(CACHE_SHELL).then(c => c.put(req, clone));
          }
          return res;
        }).catch(() => caches.match('./index.html'));
      })
    );
    return;
  }

  // ─ 외부 정적 자원 (Google Fonts 등): stale-while-revalidate
  event.respondWith(
    caches.match(req).then(cached => {
      const fetchPromise = fetch(req).then(res => {
        if (res && res.status === 200 && res.type !== 'opaque') {
          const clone = res.clone();
          caches.open(CACHE_RUNTIME).then(c => c.put(req, clone));
        }
        return res;
      }).catch(() => cached);
      return cached || fetchPromise;
    })
  );
});
