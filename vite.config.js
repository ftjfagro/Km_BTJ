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
      registerType: 'prompt', // avisa quando há versão nova (não atualiza sozinho no meio do uso)
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
