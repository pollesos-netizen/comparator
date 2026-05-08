import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { nodePolyfills } from "vite-plugin-node-polyfills";

export default defineConfig({
  base: "./",
  plugins: [react(), nodePolyfills({ include: ["buffer"] })],
  optimizeDeps: {
    include: ["cfb"],
  },
  worker: {
    format: "es",
  },
  build: {
    rollupOptions: {
      output: {
        manualChunks: (id) => {
          if (id.includes("pdf.worker")) return "pdfworker";
        },
      },
    },
  },
});