import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { basename, dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { checkoutPinned } from '../../src/pinned-checkout.ts';
import { assertFormatMatchesExtension } from '../../src/workbook-format.ts';

const UPSTREAM = 'https://github.com/LibreOffice/core.git';
const COMMIT = '195620c2';
const SUBDIR = 'sc/qa/unit/data/functions';

const here = dirname(fileURLToPath(import.meta.url));
const cache = join(here, '..', '..', '.cache', 'libreoffice');
const workbooks = join(here, 'workbooks');

if (!existsSync(join(cache, '.git'))) {
  mkdirSync(dirname(cache), { recursive: true });
  execFileSync('git', [ 'clone', '--filter=blob:none', '--sparse', '--no-checkout', UPSTREAM, cache ], { stdio: 'inherit' });
  execFileSync('git', [ '-C', cache, 'sparse-checkout', 'set', SUBDIR ], { stdio: 'inherit' });
}
checkoutPinned(cache, COMMIT);
const commit = execFileSync('git', [ '-C', cache, 'rev-parse', 'HEAD' ], { encoding: 'utf8' }).trim();
const date = execFileSync('git', [ '-C', cache, 'log', '-1', '--format=%cI' ], { encoding: 'utf8' }).trim();

rmSync(workbooks, { recursive: true, force: true });
const source = join(cache, SUBDIR);
const files: Array<{ file: string, category: string, savedBy: string }> = [];
for (const category of readdirSync(source, { withFileTypes: true }).filter(d => d.isDirectory()).map(d => d.name).sort()) {
  const fodsDir = category === 'fods' ? join(source, category) : join(source, category, 'fods');
  if (!existsSync(fodsDir)) {
    continue;
  }
  const fods = readdirSync(fodsDir).filter(f => f.endsWith('.fods')).sort().map(f => join(fodsDir, f));
  const outDir = join(workbooks, category);
  mkdirSync(outDir, { recursive: true });
  execFileSync('soffice', [ '--headless', '--convert-to', 'xlsx', '--outdir', outDir, ...fods ], { stdio: [ 'ignore', 'ignore', 'inherit' ] });
  for (const f of fods) {
    const name = basename(f, '.fods') + '.xlsx';
    if (existsSync(join(outDir, name))) {
      assertFormatMatchesExtension(join(outDir, name));
      files.push({ file: `${category}/${name}`, category, savedBy: 'LibreOffice (soffice --convert-to xlsx)' });
    }
    else {
      console.error(`conversion produced no output for ${category}/${basename(f)}`);
    }
  }
  console.error(`${category}: ${fods.length} fods`);
}

const soffice = execFileSync('soffice', [ '--version' ], { encoding: 'utf8' }).trim();

writeFileSync(join(here, 'manifest.json'), JSON.stringify({
  suite: 'libreoffice',
  upstream: UPSTREAM,
  commit,
  commitDate: date,
  subdir: SUBDIR,
  license: 'MPL-2.0 (repository); fods files carry no per-file header',
  shape: 'self-checking workbook: column A formula, column B expected literal, column C check, sheet 1 summary',
  converter: soffice,
  importedAt: new Date().toISOString(),
  files,
}, null, 2) + '\n');

console.log(`libreoffice: ${files.length} workbooks at ${commit.slice(0, 8)} (${date})`);
