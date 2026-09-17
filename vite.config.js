import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { VitePWA } from "vite-plugin-pwa";
import { nodePolyfills } from "vite-plugin-node-polyfills";

const BASE = "/Aeon-stock-app/";

// base must match the GitHub Pages project path: https://<user>.github.io/Aeon-stock-app/
export default defineConfig({
  base: BASE,
  resolve: {
    alias: {
      // El "main" del package.json de exceljs apunta a su build para Node (usa fs/stream de
      // verdad); el build para navegador es un bundle aparte pensado para <script> tags. Sin
      // este alias, Vite puede terminar resolviendo el de Node al hacer `vite build` y fallar
      // al no poder resolver esos módulos nativos.
      exceljs: "exceljs/dist/exceljs.min.js",
    },
  },
  plugins: [
    react(),
    // El bundle de navegador de exceljs (armado con browserify) todavía espera Buffer/process
    // como globales, algo que el navegador no provee — sin este polyfill tira
    // "Buffer is not defined" en tiempo de ejecución.
    nodePolyfills(),
    VitePWA({
      registerType: "autoUpdate",
      includeAssets: ["icons/apple-touch-icon.png"],
      workbox: {
        // Por default el service worker de la PWA reescribe CUALQUIER navegación (click en un
        // <a>, escribir la URL) hacia index.html, para que la app funcione offline como SPA.
        // Eso también atrapaba los links a archivos estáticos propios — manual.html,
        // presentacion-comercial.pdf, el futuro catalogo.pdf — sirviendo la app en vez del
        // archivo real. Como la app no tiene rutas propias (todo es un solo path con estado
        // en React), alcanza con excluir del fallback cualquier URL que termine en extensión.
        navigateFallbackDenylist: [/\.[^/]+$/],
      },
      manifest: {
        name: "AEON",
        short_name: "AEON",
        description: "Control de stock e inventario AEON",
        start_url: BASE,
        scope: BASE,
        display: "standalone",
        background_color: "#F5F6F8",
        theme_color: "#565A5F",
        icons: [
          { src: "icons/icon-192.png", sizes: "192x192", type: "image/png" },
          { src: "icons/icon-512.png", sizes: "512x512", type: "image/png" },
          { src: "icons/icon-512-maskable.png", sizes: "512x512", type: "image/png", purpose: "maskable" },
        ],
      },
    }),
  ],
});
