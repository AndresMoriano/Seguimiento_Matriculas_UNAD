import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Configuración de compilación. base: "./" hace que las rutas sean relativas,
// para que funcione igual en Netlify o en cualquier subcarpeta.
export default defineConfig({
  plugins: [react()],
  base: "./",
  build: {
    outDir: "dist",
    chunkSizeWarningLimit: 1200,
  },
});
