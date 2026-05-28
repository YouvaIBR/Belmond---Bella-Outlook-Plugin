import { defineConfig } from "vite";

export default defineConfig({
  root: ".",
  base: process.env.NODE_ENV === "production" ? "/Belmond---Bella-Outlook-Plugin/" : "/",
  build: {
    outDir: "dist",
    rollupOptions: {
      input: {
        taskpane: "index.html",
        commands: "commands.html",
        replyDialog: "reply-dialog.html",
      },
    },
  },
  server: {
    port: 3000,
    https: false,
  },
});
