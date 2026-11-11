import { fileURLToPath, URL } from 'node:url'

import vue from '@vitejs/plugin-vue'
import { defineConfig } from 'vite'

export default defineConfig({
  plugins: [vue()],
  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url)),
    },
  },
  server: {
    // 개발 중 /api 요청은 로컬 백엔드(8080)로 넘긴다.
    proxy: {
      '/api': 'http://localhost:8080',
    },
  },
})
