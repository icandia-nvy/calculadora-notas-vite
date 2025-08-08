// vite.config.js
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// Usa el nombre EXACTO del repo:
const repo = 'calculadora-notas-vite'

export default defineConfig({
  base: process.env.GITHUB_PAGES ? `/${repo}/` : '/',
  plugins: [react()],
})