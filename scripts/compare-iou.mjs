#!/usr/bin/env node
/**
 * LibreOffice vs Almadocx visual compare for fixtures/iou-form.docx
 * Usage: node scripts/compare-iou.mjs
 */
import { mkdirSync } from 'node:fs'
import { resolve, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { execSync } from 'node:child_process'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const out = '/tmp/iou-compare'
mkdirSync(join(out, 'lo'), { recursive: true })
mkdirSync(join(out, 'alma'), { recursive: true })
mkdirSync(join(out, 'diff'), { recursive: true })

const fixture = join(root, 'fixtures/iou-form.docx')
execSync(`soffice --headless --convert-to pdf --outdir ${join(out, 'lo')} "${fixture}"`, {
  stdio: 'inherit',
})
execSync(`pdftoppm -png -r 96 ${join(out, 'lo', 'iou-form.pdf')} ${join(out, 'lo', 'page')}`, {
  stdio: 'inherit',
})
execSync(`node ${join(root, 'scripts/render-fixture.mjs')} "${fixture}" ${join(out, 'alma')}`, {
  stdio: 'inherit',
})

for (const n of [1, 2]) {
  const a = join(out, 'alma', `page-${n}.png`)
  const l = join(out, 'lo', `page-${n}.png`)
  try {
    execSync(`magick compare -metric MAE "${a}" "${l}" null:`, { encoding: 'utf8', stdio: 'pipe' })
  } catch (e) {
    // ImageMagick compare exits 1 when images differ; metric is on stderr.
    const metric = (e.stderr || e.stdout || '').toString().trim()
    console.log(`page-${n} MAE:`, metric)
  }
  execSync(
    `magick compare -metric RMSE "${a}" "${l}" "${join(out, 'diff', `p${n}-diff.png`)}" || true`,
    { stdio: 'inherit', shell: '/bin/bash' },
  )
  execSync(`magick "${a}" "${l}" +append "${join(out, 'diff', `p${n}-side.png`)}"`, {
    stdio: 'inherit',
  })
}
console.log('Side-by-side:', join(out, 'diff/p1-side.png'), join(out, 'diff/p2-side.png'))
