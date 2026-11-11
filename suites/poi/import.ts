import { execFileSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { checkoutPinned } from '../../src/pinned-checkout.ts';
import { assertFormatMatchesExtension } from '../../src/workbook-format.ts';

const UPSTREAM = 'https://github.com/apache/poi.git';
const COMMIT = '6327f056ba3cfdae326a0261a236460592b80f5c';
const SUBDIR = 'test-data/spreadsheet';
const RUNNERS = 'poi/src/test/java/org/apache/poi/ss/formula/functions';

const here = dirname(fileURLToPath(import.meta.url));
const cache = join(here, '..', '..', '.cache', 'poi');
const workbooks = join(here, 'workbooks');

if (!existsSync(join(cache, '.git'))) {
  mkdirSync(dirname(cache), { recursive: true });
  execFileSync('git', [ 'clone', '--filter=blob:none', '--sparse', '--no-checkout', UPSTREAM, cache ], { stdio: 'inherit' });
}
execFileSync('git', [ '-C', cache, 'sparse-checkout', 'set', '--no-cone', `/${RUNNERS}/` ], { stdio: 'inherit' });
checkoutPinned(cache, COMMIT);
const date = execFileSync('git', [ '-C', cache, 'log', '-1', '--format=%cI' ], { encoding: 'utf8' }).trim();

function functionWorkbooks (): string[] {
  const dir = join(cache, RUNNERS);
  return readdirSync(dir)
    .filter(f => f.endsWith('FromSpreadsheet.java') || f.endsWith('FromSpreadsheet2013.java'))
    .flatMap(f => {
      const match = /dataFromClass\(\w+\.class,\s*"([^"]+\.xlsx?)"/.exec(readFileSync(join(dir, f), 'utf8'));
      return match ? [ match[1] ] : [];
    })
    .sort();
}

const families: Record<string, string[]> = {
  eval: [ 'FormulaEvalTestData.xls', 'FormulaEvalTestData_Copy.xlsx' ],
  functions: functionWorkbooks(),
  // POI tests these two in .xls and .xlsx alike. The .xls has the same cells and maps to the same
  // fixture, so only the .xlsx is imported.
  matrix: [ 'MatrixFormulaEvalTestData.xlsx' ],
  multisheet: [ 'FormulaSheetRange.xlsx' ],
};

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

const patterns = Object.values(families).flat().map(name => `/${SUBDIR}/${name}`);
execFileSync('git', [ '-C', cache, 'sparse-checkout', 'add', ...patterns ], { stdio: 'inherit' });

rmSync(workbooks, { recursive: true, force: true });
const files: Array<{ file: string, category: string, savedBy: string }> = [];
for (const [ category, names ] of Object.entries(families)) {
  mkdirSync(join(workbooks, category), { recursive: true });
  for (const name of names) {
    const target = join(workbooks, category, name);
    copyFileSync(join(cache, SUBDIR, name), target);
    assertFormatMatchesExtension(target);
    files.push({ file: `${category}/${name}`, category, savedBy: savedBy(target) });
  }
  console.error(`${category}: ${names.length} workbooks`);
}

writeFileSync(join(here, 'manifest.json'), JSON.stringify({
  suite: 'poi',
  upstream: UPSTREAM,
  commit: COMMIT,
  commitDate: date,
  subdir: SUBDIR,
  license: 'Apache-2.0',
  shape: 'expected-literal workbook: eval = 4-row blocks, formula row then literal row, columns D onward; functions = column B formula, column C literal, column A marker; matrix = 4-row blocks, formulas H:J, literals K:M; multisheet = sheet "test", column D formula, column C literal',
  importedAt: new Date().toISOString(),
  files,
}, null, 2) + '\n');

console.log(`poi: ${files.length} workbooks at ${COMMIT.slice(0, 8)} (${date})`);
