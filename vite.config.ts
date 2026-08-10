import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

/** Electron 이 file:// 로 index.html 을 열기 때문에 상대 경로 빌드가 필요하다. */
const RENDERER_PORT = 5173;

export default defineConfig({
  plugins: [react()],
  base: './',
  server: {
    port: RENDERER_PORT,
    strictPort: true,
  },
  build: {
    outDir: 'dist',
    emptyOutDir: true,
  },
});
