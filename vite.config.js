import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  // ⚠️  Sostituisci <REPO> con il nome del tuo repository quando deployi su GitHub Pages
  base: process.env.NODE_ENV === 'production' ? '/<REPO>/' : '/'
});