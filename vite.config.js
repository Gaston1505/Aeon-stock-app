import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// base must match the GitHub Pages project path: https://<user>.github.io/Aeon-stock-app/
export default defineConfig({
  plugins: [react()],
  base: "/Aeon-stock-app/",
});
