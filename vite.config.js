import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  base: './', // чтобы собранное приложение открывалось через file:// в Electron
  plugins: [react()],
  server: { port: 5173, strictPort: true },
  build: { chunkSizeWarningLimit: 2000 },
  test: { include: ['tests/**/*.test.js'] },
});
