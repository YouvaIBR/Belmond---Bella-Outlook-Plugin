import { defineConfig } from "vite";

// `base` defaults to "/" (local dev, mock, Firebase). The GitHub Pages build
// serves under the repo sub-path, so the workflow sets VITE_BASE accordingly.
export default defineConfig({
  root: ".",
  base: process.env.VITE_BASE ?? "/",
  build: {
    outDir: "dist",
    rollupOptions: {
      input: {
        taskpane: "index.html",
        workday: "workday.html",
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
