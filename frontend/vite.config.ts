import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

export default defineConfig({
  plugins: [react(), tailwindcss()],
  server: {
    // 외부(사내망) 접근 차단 — 루프백에만 바인딩한다.
    // CLI 의 --host 옵션이 이 값을 덮으므로 실행 명령에도 --host 를 붙이지 않는다.
    host: '127.0.0.1',
    port: 5173,
    proxy: {
      '/api': 'http://localhost:8001',
    },
  },
})
