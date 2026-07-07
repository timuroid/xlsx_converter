import { defineConfig } from "vite";

export default defineConfig({
  build: {
    rollupOptions: {
      input: {
        index: "index.html",
      },
      output: {
        entryFileNames: "assets/[name]-[hash].js",
      },
    },
  },
});
