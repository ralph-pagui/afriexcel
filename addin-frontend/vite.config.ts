import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  // Base relative pour garantir que les assets se chargent toujours
  // peu importe où l'add-in est hébergé (Netlify, Vercel, localhost ou sous-dossier)
  base: './',
  build: {
    outDir: 'dist',
    target: 'es2015', // Rétrocompatibilité avec Excel 2016 / WebView2 / WebViews plus anciennes
    assetsDir: 'assets',
    sourcemap: false,
    rollupOptions: {
      output: {
        manualChunks: {
          vendor: ['react', 'react-dom'],
          icons: ['lucide-react']
        }
      }
    }
  },
  server: {
    port: 3000,
    https: false // Dev local HTTP ou HTTPS
  }
});
