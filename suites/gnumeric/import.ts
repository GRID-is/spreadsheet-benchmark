import { execFileSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { checkoutPinned } from '../../src/pinned-checkout.ts';
import { assertFormatMatchesExtension } from '../../src/workbook-format.ts';

const UPSTREAM = 'https://github.com/GNOME/gnumeric.git';
const COMMIT = '09a37c89f701712d8c71a71058c751650778a9c1';
const SUBDIR = 'samples';

const here = dirname(fileURLToPath(import.meta.url));
const cache = join(here, '..', '..', '.cache', 'gnumeric');
const workbooks = join(here, 'workbooks');

const families: Record<string, string[]> = {
  functions: [ 'statfuns', 'mathfuns', 'finfuns', 'logfuns', 'textfuns', 'engfuns', 'datefuns', 'infofuns', 'lookfuns', 'dbfuns' ].map(n => `excel/${n}.xls`),
  yalta2008: [ 'excel/yalta2008.xls' ],
  overview: [ 'excel12/database.xlsx', 'excel12/countif.xlsx', 'excel12/ifs-funcs.xlsx' ],
};

if (!existsSync(join(cache, '.git'))) {
  mkdirSync(dirname(cache), { recursive: true });
  execFileSync('git', [ 'clone', '--filter=blob:none', '--sparse', '--no-checkout', UPSTREAM, cache ], { stdio: 'inherit' });
}
const patterns = Object.values(families).flat().map(name => `/${SUBDIR}/${name}`);
execFileSync('git', [ '-C', cache, 'sparse-checkout', 'set', '--no-cone', ...patterns ], { stdio: 'inherit' });
checkoutPinned(cache, COMMIT);
const date = execFileSync('git', [ '-C', cache, 'log', '-1', '--format=%cI' ], { encoding: 'utf8' }).trim();

function savedBy (path: string): string {
  if (!path.endsWith('.xlsx')) {
    return '';
  }
  try {
    const xml = execFileSync('unzip', [ '-p', path, 'docProps/app.xml' ], { encoding: 'utf8', stdio: [ 'ignore', 'pipe', 'ignore' ] });
    return /<Application>([^<]*)<\/Application>/.exec(xml)?.[1] ?? '';
  }
  catch {
    return '';
  }
}

rmSync(workbooks, { recursive: true, force: true });
const files: Array<{ file: string, category: string, savedBy: string }> = [];
for (const [ category, names ] of Object.entries(families)) {
  mkdirSync(join(workbooks, category), { recursive: true });
  for (const name of names) {
    const target = join(workbooks, category, name.split('/')[1]);
    copyFileSync(join(cache, SUBDIR, name), target);
    assertFormatMatchesExtension(target);
    files.push({ file: `${category}/${name.split('/')[1]}`, category, savedBy: savedBy(target) });
  }
  console.error(`${category}: ${names.length} workbooks`);
}

writeFileSync(join(here, 'manifest.json'), JSON.stringify({
  suite: 'gnumeric',
  upstream: UPSTREAM,
  commit: COMMIT,
  commitDate: date,
  subdir: SUBDIR,
  license: 'GPL-2.0-only OR GPL-3.0-only',
  shape: 'self-checking workbooks from test/t1000-t1023: functions = A4 reads "All ok"; yalta2008 = D10:D99 TRUE; overview = Overview!C2:C99 all 0',
  importedAt: new Date().toISOString(),
  files,
}, null, 2) + '\n');

console.log(`gnumeric: ${files.length} workbooks at ${COMMIT.slice(0, 8)} (${date})`);
