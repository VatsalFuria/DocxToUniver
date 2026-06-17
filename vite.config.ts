import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],

  // Univer uses some Node.js built-ins; tell Vite to polyfill them
  optimizeDeps: {
    // Pre-bundle these so Vite doesn't struggle with their ESM/CJS mix
    include: [
      "@univerjs/presets",
      "@univerjs/presets/preset-docs-core",
      "jszip",
    ],
  },

  build: {
    // Univer bundles can be large; increase the chunk size warning threshold
    chunkSizeWarningLimit: 4000,

    rollupOptions: {
      output: {
        // Split Univer into its own chunk to keep the app chunk small
        manualChunks: {
          univer: ["@univerjs/presets"],
        },
      },
    },
  },
});
