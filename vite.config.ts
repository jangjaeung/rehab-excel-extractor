import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

/** 개발 서버 포트 */
const DEV_PORT = 5173;

export default defineConfig({
  plugins: [react()],
  server: {
    port: DEV_PORT,
    strictPort: true,
  },
  build: {
    outDir: 'dist',
    emptyOutDir: true,
  },
});
