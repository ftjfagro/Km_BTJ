import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { VitePWA } from 'vite-plugin-pwa'
export default defineConfig({
  base: '/Km_BTJ/', // GitHub Pages de projeto — sem isso os caminhos quebram
  build: {
    outDir: 'dist',
  },
  plugins: [
    react(),
    VitePWA({
      // ATUALIZADO: 'prompt' exigia o usuário notar o banner e tocar nele pra
      // atualizar — se isso não acontecesse antes do próximo login, o app
      // corria o risco de tentar carregar com uma "casca" antiga (HTML)
      // apontando pra arquivos JS que o Vite já tinha substituído no último
      // build, resultando em tela branca até um refresh manual forçar a
      // busca da versão nova. 'autoUpdate' aplica a versão nova sozinho assim
      // que detecta — troca: se um deploy acontecer bem no meio de você
      // digitar um km (raro), o campo ainda não salvo pode se perder no reload.
      registerType: 'autoUpdate',
      // O manifest agora é o arquivo estático public/manifest.webmanifest
      // (novo ícone velocímetro navy BTJ) — o plugin não gera mais o dele.
      manifest: false,
      includeAssets: [
        'icons/favicon-32.png',
        'icons/apple-touch-icon.png',
        'icons/icon-192.png',
        'icons/icon-512.png',
        'icons/icon-maskable-192.png',
        'icons/icon-maskable-512.png',
        'icons/icon.svg',
      ],
      workbox: {
        // Guarda o app inteiro (JS, CSS, HTML, ícones) para abrir 100% offline.
        globPatterns: ['**/*.{js,css,html,ico,png,svg,webmanifest,woff,woff2}'],
        // Nunca cacheia as chamadas ao Apps Script (dados sempre frescos quando online).
        navigateFallbackDenylist: [/^\/macros/],
        // Garante que a versão nova assume o controle imediatamente (sem
        // esperar todas as abas antigas fecharem) e limpa caches de builds
        // anteriores — reforça o autoUpdate acima e evita "sobra" de uma
        // versão antiga enquanto a nova já deveria estar no comando.
        skipWaiting: true,
        clientsClaim: true,
        cleanupOutdatedCaches: true,
        runtimeCaching: [
          {
            // Mapa/GPS (reverse geocoding): usa cache como reserva se offline.
            urlPattern: /^https:\/\/nominatim\.openstreetmap\.org\/.*/i,
            handler: 'NetworkFirst',
            options: {
              cacheName: 'gps-cache',
              expiration: { maxEntries: 30, maxAgeSeconds: 60 * 60 * 24 * 7 },
            },
          },
        ],
      },
      devOptions: { enabled: false },
    }),
  ],
})
