import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vitest/config'

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..')

export default defineConfig({
  resolve: {
    alias: {
      '@almadocx/core': path.join(root, 'packages/core/src/index.ts'),
    },
  },
  test: {
    include: ['tests/**/*.test.ts'],
    environment: 'jsdom',
    coverage: {
      provider: 'v8',
      reporter: ['text', 'json-summary', 'lcov'],
      reportsDirectory: './coverage',
      include: ['src/**/*.ts'],
      exclude: [
        'src/**/*.d.ts',
        // Public barrel only re-exports.
        'src/index.ts',
        // Full editor shell is exercised via harness/manual smoke; unit coverage
        // focuses on selection, hit-test, a11y, and sanitize modules.
        'src/editor.ts',
        'src/render/paint.ts',
      ],
      thresholds: {
        statements: 100,
        branches: 100,
        functions: 100,
        lines: 100,
      },
    },
  },
})
