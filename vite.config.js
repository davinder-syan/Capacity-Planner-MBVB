import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// If you deploy this to GitHub Pages under a repo subpath (https://username.github.io/repo-name/),
// set `base` below to "/repo-name/". For Vercel/Netlify or a custom domain, leave it as "/".
export default defineConfig({
  plugins: [react()],
  base: "/"
});
