// Importa o motor do OneSignal (Deve ser a primeira linha do arquivo)
importScripts("https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js");

const CACHE_NAME = 'missrose-v9';
const urlsToCache = [

    './',
    './index.html',
    './styles.css',
    './app.js',
    './logo_missrose.png',
    './manifest.json',
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css',
    'https://cdn.jsdelivr.net/npm/sweetalert2@11',
    'https://unpkg.com/imask',
    'https://cdn.jsdelivr.net/npm/chart.js',
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
];

// Instalação do Service Worker
self.addEventListener('install', (e) => {
    console.log('[Service Worker] App Instalado!');
    e.waitUntil(
        caches.open(CACHE_NAME)
        .then((cache) => {
            return cache.addAll(urlsToCache);
        })
    );
});

// Ativação do Service Worker e Limpeza de Caches Antigos
self.addEventListener('activate', (e) => {
    console.log('[Service Worker] Ativado');
    e.waitUntil(
        caches.keys().then((cacheNames) => {
            return Promise.all(
                cacheNames.map((cacheName) => {
                    if (cacheName !== CACHE_NAME) {
                        console.log('[Service Worker] Removendo cache antigo:', cacheName);
                        return caches.delete(cacheName);
                    }
                })
            );
        })
    );
});

// Interceptação de rede: Estratégia Stale-While-Revalidate para os arquivos do app
self.addEventListener('fetch', (e) => {
    // Não fazemos cache de chamadas da API e Planilhas do Google
    if (e.request.url.includes('script.google.com') || e.request.url.includes('docs.google.com')) {
        return;
    }

    e.respondWith(
        caches.match(e.request).then((cachedResponse) => {
            const fetchPromise = fetch(e.request).then((networkResponse) => {
                caches.open(CACHE_NAME).then((cache) => {
                    cache.put(e.request, networkResponse.clone());
                });
                return networkResponse;
            }).catch(() => { /* Tratar fallback de conexão, se necessário */ });
            
            return cachedResponse || fetchPromise;
        })
    );
});

// Ouve mensagens da página (app.js) para pular a espera e ativar a nova versão
self.addEventListener('message', (event) => {
    if (event.data && event.data.action === 'skipWaiting') {
        self.skipWaiting();
    }
});