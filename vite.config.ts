import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { loadTls } from './server/tls';

export default defineConfig(({ command }) => ({
  plugins: [react()],
  server: {
    host: 'localhost',
    port: Number(process.env.PORT || 3001),
    strictPort: true,
    fs: {
      deny: ['.env', '.env.*', '*.{crt,pem,key}', '**/.local/**', '**/certs/**', '**/.git/**'],
    },
    https: process.env.HTTP_ONLY === '1' || command === 'build' ? undefined : loadTls(),
    proxy: {
      '/api': { target: `http://127.0.0.1:${process.env.API_PORT || 4318}`, changeOrigin: true },
    },
  },
  build: {
    rollupOptions: { input: { index: 'index.html', taskpane: 'taskpane.html' } },
    chunkSizeWarningLimit: 1300,
  },
}));
