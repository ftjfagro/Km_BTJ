import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { VitePWA } from 'vite-plugin-pwa'

export default defineConfig({
  base: './',
  build: {
    outDir: 'dist',
  },
  plugins: [
    react(),
    VitePWA({
      registerType: 'prompt', // avisa quando há versão nova (não atualiza sozinho no meio do uso)
      includeAssets: ['favicon.ico', 'icon-16.png', 'icon-32.png', 'icon-180.png', 'logo.png'],
      manifest: {
        name: 'Km BTJ',
        short_name: 'Km BTJ',
        description: 'Registro de quilometragem e reembolso — BTJ',
        start_url: './',
        scope: './',
        display: 'standalone',
        orientation: 'portrait',
        background_color: '#001F3E',
        theme_color: '#001F3E',
        icons: [
          { src: 'icon-192.png', sizes: '192x192', type: 'image/png', purpose: 'any' },
          { src: 'icon-512.png', sizes: '512x512', type: 'image/png', purpose: 'any' },
          { src: 'icon-512.png', sizes: '512x512', type: 'image/png', purpose: 'maskable' },
        ],
      },
      workbox: {
        // Guarda o app inteiro (JS, CSS, HTML, ícones) para abrir 100% offline.
        globPatterns: ['**/*.{js,css,html,ico,png,svg,woff,woff2}'],
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
