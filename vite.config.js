import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// base: './' makes the build use relative asset paths, which is what
// GitHub Pages needs when the site is served from a subfolder like
// https://<user>.github.io/<repo>/
export default defineConfig({
  plugins: [react()],
  base: './',
});
