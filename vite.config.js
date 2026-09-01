import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { VitePWA } from "vite-plugin-pwa";

const BASE = "/Aeon-stock-app/";

// base must match the GitHub Pages project path: https://<user>.github.io/Aeon-stock-app/
export default defineConfig({
  base: BASE,
  plugins: [
    react(),
    VitePWA({
      registerType: "autoUpdate",
      includeAssets: ["icons/apple-touch-icon.png"],
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
