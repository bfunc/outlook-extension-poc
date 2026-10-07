import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// The site is hosted under https://<host>/static/outlook/, so every asset URL is built with that prefix.
export default defineConfig({
  base: "/static/outlook/",
  plugins: [react()],
  build: { target: "es2020" },
});
