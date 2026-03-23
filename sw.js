// Instalação do Service Worker
self.addEventListener('install', (e) => {
    console.log('[Service Worker] App Instalado!');
});

// Permite interceptação de rede (Requisito obrigatório do Google Chrome para PWA)
self.addEventListener('fetch', (e) => {});