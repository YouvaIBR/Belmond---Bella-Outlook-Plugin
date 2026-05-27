import { defineConfig } from "vite";

export default defineConfig({
  root: ".",
  base: process.env.NODE_ENV === "production" ? "/Belmond---Bella-Outlook-Plugin/" : "/",
  build: {
    outDir: "dist",
    rollupOptions: {
      input: "index.html",
    },
  },
  server: {
    port: 3000,
    https: false,
  },
});
