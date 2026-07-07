import { defineConfig } from "vite";

export default defineConfig({
  build: {
    emptyOutDir: true,
    lib: {
      entry: "src/transformer-test-entry.ts",
      formats: ["es"],
      fileName: () => "transformer-test-entry.mjs",
    },
    outDir: "dist-test",
  },
});
