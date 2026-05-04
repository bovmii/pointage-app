import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// Sur GitHub Pages, l'URL est https://<user>.github.io/<repo>/
// On lit le nom du repo depuis la variable d'env GITHUB_REPOSITORY (fournie par GH Actions)
// En local (npm run dev), base = '/'
const repoName = process.env.GITHUB_REPOSITORY?.split('/')[1];

export default defineConfig({
  plugins: [react()],
  base: process.env.NODE_ENV === 'production' && repoName ? `/${repoName}/` : '/',
});
