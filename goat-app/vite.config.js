import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { VitePWA } from "vite-plugin-pwa";

export default defineConfig({
  plugins: [
    react(),
    VitePWA({
      registerType: "autoUpdate",
      includeAssets: ["apple-touch-icon.png"],
      manifest: {
        name: "GOAT · Frutos Secos Premium",
        short_name: "GOAT",
        description: "App de gestión de inventario, pedidos, ventas y costos para GOAT.",
        theme_color: "#1c1917",
        background_color: "#1c1917",
        display: "standalone",
        orientation: "portrait",
        start_url: "/",
        scope: "/",
        icons: [
          { src: "icon-192.png", sizes: "192x192", type: "image/png" },
          { src: "icon-512.png", sizes: "512x512", type: "image/png" },
          { src: "icon-maskable-512.png", sizes: "512x512", type: "image/png", purpose: "maskable" },
        ],
      },
      workbox: {
        // Cachea los archivos de la app para que abra rápido y funcione
        // aunque no haya conexión (los datos en sí siguen necesitando
        // internet para sincronizar con Supabase).
        globPatterns: ["**/*.{js,css,html,ico,png,svg}"],
      },
    }),
  ],
});
