import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// Sur GitHub Pages, l'URL est https://<user>.github.io/<repo>/
// On lit le nom du repo depuis GITHUB_REPOSITORY (injecté par GH Actions)
// En local (npm run dev / build), base = '/'
const repoName = process.env.GITHUB_REPOSITORY?.split('/')[1];

export default defineConfig(({ command }) => ({
  plugins: [react()],
  base: command === 'build' && repoName ? `/${repoName}/` : '/',
}));
