import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { VitePWA } from "vite-plugin-pwa";

export default defineConfig({
  plugins: [
    react(),
    VitePWA({
      registerType: "autoUpdate",
      includeAssets: ["icons/icon-192.png", "icons/icon-512.png", "icons/icon-180.png", "icons/icon-167.png", "icons/icon-152.png"],
      manifest: {
        name: "GeoAcervo — Gestão de Minerais e Rochas",
        short_name: "GeoAcervo",
        description: "Sistema de gestão de acervo de minerais e rochas para laboratório de geologia.",
        theme_color: "#0B3D2E",
        background_color: "#0B3D2E",
        display: "standalone",
        orientation: "portrait",
        start_url: "/",
        scope: "/",
        icons: [
          { src: "icons/icon-192.png", sizes: "192x192", type: "image/png" },
          { src: "icons/icon-512.png", sizes: "512x512", type: "image/png" },
          { src: "icons/icon-512.png", sizes: "512x512", type: "image/png", purpose: "maskable" }
        ]
      }
    })
  ]
});
