import { execFileSync } from 'node:child_process';
import { cpSync, existsSync, mkdirSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { checkoutPinned } from '../../src/pinned-checkout.ts';
import { assertFormatMatchesExtension } from '../../src/workbook-format.ts';

const UPSTREAM = 'https://github.com/ironcalc/IronCalc.git';
const COMMIT = '2128db12ab30b77fdebde9cba994273a26438ae3';
const SUBDIR = 'xlsx/tests/calc_tests';

const here = dirname(fileURLToPath(import.meta.url));
const cache = join(here, '..', '..', '.cache', 'ironcalc');
const workbooks = join(here, 'workbooks');

if (!existsSync(join(cache, '.git'))) {
  mkdirSync(dirname(cache), { recursive: true });
  execFileSync('git', [ 'clone', '--filter=blob:none', '--sparse', '--no-checkout', UPSTREAM, cache ], { stdio: 'inherit' });
  execFileSync('git', [ '-C', cache, 'sparse-checkout', 'set', SUBDIR ], { stdio: 'inherit' });
}
checkoutPinned(cache, COMMIT);
const date = execFileSync('git', [ '-C', cache, 'log', '-1', '--format=%cI' ], { encoding: 'utf8' }).trim();

rmSync(workbooks, { recursive: true, force: true });
cpSync(join(cache, SUBDIR), workbooks, { recursive: true });

function walk (dir: string): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap(entry => {
    const path = join(dir, entry.name);
    return entry.isDirectory() ? walk(path) : entry.name.endsWith('.xlsx') ? [ path ] : [];
  });
}

function savedBy (path: string): string {
  try {
    const xml = execFileSync('unzip', [ '-p', path, 'docProps/app.xml' ], { encoding: 'utf8', stdio: [ 'ignore', 'pipe', 'ignore' ] });
    return /<Application>([^<]*)<\/Application>/.exec(xml)?.[1] ?? '';
  }
  catch {
    return '';
  }
}

const files = walk(workbooks).sort().map(path => {
  assertFormatMatchesExtension(path);
  const rel = relative(workbooks, path);
  const category = rel.includes('/') ? rel.split('/')[0] : '(root)';
  return { file: rel, category, savedBy: savedBy(path) };
});

writeFileSync(join(here, 'manifest.json'), JSON.stringify({
  suite: 'ironcalc',
  upstream: UPSTREAM,
  commit: COMMIT,
  commitDate: date,
  subdir: SUBDIR,
  license: 'MIT OR Apache-2.0',
  shape: 'cached-value workbook',
  importedAt: new Date().toISOString(),
  files,
}, null, 2) + '\n');

console.log(`ironcalc: ${files.length} workbooks at ${COMMIT.slice(0, 8)} (${date})`);
