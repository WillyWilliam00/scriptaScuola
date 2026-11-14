import path from "path"
import tailwindcss from "@tailwindcss/vite"
import react from "@vitejs/plugin-react"
import { defineConfig } from "vite"

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), tailwindcss()],
  build: {
    rollupOptions: {
      output: {
        manualChunks: {
          "vendor-react": ["react", "react-dom", "react-router-dom"],
          "vendor-tanstack": ["@tanstack/react-query", "@tanstack/react-form"],
        },
      },
    },
    chunkSizeWarningLimit: 500,
  },
  resolve: {
    alias: {
      "@": path.resolve(__dirname, "./src"),
      "@shared": path.resolve(__dirname, "../shared"),
    },
  },
  server: {
    host: true,
    port: 3000,
    // In sviluppo le chiamate a /api passano dal server di Vite verso il backend locale,
    // come in produzione passano da Vercel verso Render: il browser vede un solo dominio
    // e il cookie di sessione funziona allo stesso modo.
    proxy: {
      "/api": {
        target: process.env.VITE_BACKEND_URL || "http://localhost:3001",
        changeOrigin: false,
      },
    },
  },
})
