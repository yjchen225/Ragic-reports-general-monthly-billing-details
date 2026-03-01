import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  base: "/Ragic-reports-general-monthly-billing-details/", // ← 改成你的 repo 名稱（大小寫要一致）
  plugins: [react()],
})
