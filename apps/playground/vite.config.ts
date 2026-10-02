import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vite'

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..')

export default defineConfig({
  server: {
    port: 5173,
    open: false,
    fs: {
      // Allow importing monorepo package sources
      allow: [root],
    },
  },
  resolve: {
    // Dev: load TypeScript sources so edits hot-reload without rebuilding dist/
    alias: {
      '@almadocx/canvas': path.join(root, 'packages/canvas/src/index.ts'),
      '@almadocx/core': path.join(root, 'packages/core/src/index.ts'),
    },
  },
  optimizeDeps: {
    exclude: ['@almadocx/canvas', '@almadocx/core'],
  },
})
