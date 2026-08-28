import { defineConfig } from "vite";

export default defineConfig({
  base: "/pdf-workbench/",
  build: {
    outDir: "dist",
    emptyOutDir: true
  }
});
