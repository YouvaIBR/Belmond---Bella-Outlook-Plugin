import { defineConfig } from "vite";

export default defineConfig({
  root: ".",
  base: "/",
  build: {
    outDir: "dist",
    rollupOptions: {
      input: {
        taskpane: "index.html",
        commands: "commands.html",
        launchEvents: "launchEvents.html",
      },
    },
  },
  server: {
    port: 3000,
    https: false,
  },
});
