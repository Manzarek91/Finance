import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// base: "./" rende i percorsi degli asset relativi, così il sito funziona
// su GitHub Pages a prescindere dal nome del repository
// (https://tuonome.github.io/nome-repo/).
export default defineConfig({
  plugins: [react()],
  base: "./",
});
