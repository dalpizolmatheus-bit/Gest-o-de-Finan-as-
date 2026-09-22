// Service worker do Gestor Financeiro Pessoal.
// Cacheia o "shell" do app para abrir rápido e funcionar offline (os dados
// em si vêm do Google Sheets, então lançamentos novos exigem rede).
const CACHE_NAME = "gestor-financeiro-v3";
const SHELL_FILES = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const { request } = event;

  // Nunca cacheia chamadas ao Apps Script (dados precisam estar sempre atualizados).
  if (request.url.includes("script.google.com")) {
    return;
  }

  // A página principal (o HTML) é sempre buscada da rede primeiro, e só cai
  // pro cache se estiver offline. Assim, toda atualização publicada no
  // GitHub aparece já na primeira abertura, sem precisar abrir duas vezes
  // ou trocar a versão do cache a cada mudança.
  const isDocument = request.mode === "navigate" || request.destination === "document";
  if (isDocument) {
    event.respondWith(
      fetch(request)
        .then((response) => {
          if (response && response.ok) {
            const copy = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
          }
          return response;
        })
        .catch(() => caches.match(request))
    );
    return;
  }

  // Demais arquivos do shell (ícones, manifest): cache-first, com
  // atualização em segundo plano — mudam raramente, então prioriza velocidade.
  event.respondWith(
    caches.match(request).then((cached) => {
      const network = fetch(request)
        .then((response) => {
          if (response && response.ok) {
            const copy = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
          }
          return response;
        })
        .catch(() => cached);
      return cached || network;
    })
  );
});
