import { defineConfig } from "vite";

export default defineConfig({
  root: ".",
  base: process.env.VITE_BASE_PATH || "/",
  build: {
    outDir: "dist",
  },
});
