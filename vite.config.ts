import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import electron from 'vite-plugin-electron'
import renderer from 'vite-plugin-electron-renderer'
import { copyFileSync, mkdirSync, existsSync } from 'fs'
import { fileURLToPath } from 'url'
import path from 'path'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

function copyPreloadPlugin() {
  const src = path.join(__dirname, 'preload.cjs')
  const destDir = path.join(__dirname, 'dist-electron')
  const dest = path.join(destDir, 'preload.cjs')
  return {
    name: 'copy-preload-cjs',
    buildStart() {
      if (!existsSync(destDir)) mkdirSync(destDir, { recursive: true })
      copyFileSync(src, dest)
    },
    writeBundle() {
      if (!existsSync(destDir)) mkdirSync(destDir, { recursive: true })
      copyFileSync(src, dest)
    },
  }
}

export default defineConfig({
  server: {
    port: 3000,
    strictPort: true,
    host: '0.0.0.0',
  },
  base: './',
  plugins: [
    copyPreloadPlugin(),
    react(),
    tailwindcss(),
    electron([
      {
        entry: 'main.ts',
      },
    ]),
    renderer(),
  ],
})