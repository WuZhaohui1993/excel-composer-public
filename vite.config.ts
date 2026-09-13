/// <reference types="vitest" />
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  test: {
    exclude: ["node_modules/**", "dist/**", "dist-server/**"]
  },
  server: {
    port: 5173,
    proxy: {
      "/api": "http://127.0.0.1:3001"
    }
  },
  build: {
    outDir: "dist",
    sourcemap: true
  }
});
