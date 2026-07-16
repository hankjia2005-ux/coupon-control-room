import vinext from "vinext";
import { defineConfig } from "vite";

// GitHub Pages is a static deployment. Keep this config self-contained so the
// public build never resolves local Sites/Cloudflare-only modules.
export default defineConfig({
  plugins: [vinext()],
});
