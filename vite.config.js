import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  base: './', // чтобы собранное приложение открывалось через file:// в Electron
  plugins: [react()],
  // Явно IPv4 и свой порт: с host 'localhost' Vite мог встать на [::1], пока 127.0.0.1:5173
  // занят другим проектом, и Electron молча открывал чужое приложение. Порт синхронизирован с "dev" в package.json.
  server: { host: '127.0.0.1', port: 5183, strictPort: true },
  build: { chunkSizeWarningLimit: 2000 },
  test: { include: ['tests/**/*.test.js'] },
});
