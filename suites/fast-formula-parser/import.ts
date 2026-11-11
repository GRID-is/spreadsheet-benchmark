import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { runInNewContext } from 'node:vm';
import { type CaseSpec, writeCaseWorkbook } from '../../src/case-workbook.ts';
import type { CellValue } from '../../src/model/case.ts';
import { checkoutPinned } from '../../src/pinned-checkout.ts';
import { assertFormatMatchesExtension } from '../../src/workbook-format.ts';

const UPSTREAM = 'https://github.com/LesterLyu/fast-formula-parser.git';
const COMMIT = 'ed38a50f759c1ed918444e94e8b5876e6d1e873b';
const SUBDIR = 'test';

const here = dirname(fileURLToPath(import.meta.url));
const cache = join(here, '..', '..', '.cache', 'fast-formula-parser');
const workbooks = join(here, 'workbooks');

const FORMULA_ERRORS = { DIV0: '#DIV/0!', NA: '#N/A', NAME: '#NAME?', NULL: '#NULL!', NUM: '#NUM!', REF: '#REF!', VALUE: '#VALUE!' };

if (!existsSync(join(cache, '.git'))) {
  mkdirSync(dirname(cache), { recursive: true });
  execFileSync('git', [ 'clone', '--filter=blob:none', '--sparse', '--no-checkout', UPSTREAM, cache ], { stdio: 'inherit' });
  execFileSync('git', [ '-C', cache, 'sparse-checkout', 'set', SUBDIR ], { stdio: 'inherit' });
}
checkoutPinned(cache, COMMIT);
const date = execFileSync('git', [ '-C', cache, 'log', '-1', '--format=%cI' ], { encoding: 'utf8' }).trim();

function evaluate<T> (source: string): T {
  const module = { exports: {} };
  runInNewContext(source, { require: () => FORMULA_ERRORS, FormulaError: FORMULA_ERRORS, module });
  return module.exports as T;
}

function rowsOf (runnerSource: string): Array<Array<CellValue | undefined>> {
  const block = /const data = (\[[\s\S]*?\n\]);/.exec(runnerSource)?.[1];
  return block ? evaluate(`module.exports = ${block};`) : [];
}

function cases (testcaseSource: string): CaseSpec[] {
  const groups = evaluate<Record<string, Record<string, CellValue>>>(testcaseSource);
  return Object.entries(groups).flatMap(([ group, formulas ]) => Object.entries(formulas).map(([ formula, expected ]) => ({ group, formula, expected }))).map((c, i) => ({ ...c, ref: `AA${i + 1}` }));
}

const categories = [
  { name: 'operators', dir: join(cache, SUBDIR, 'operators') },
  ...readdirSync(join(cache, SUBDIR, 'formulas'), { withFileTypes: true }).filter(d => d.isDirectory() && d.name !== 'web').map(d => ({ name: d.name, dir: join(cache, SUBDIR, 'formulas', d.name) })),
];

rmSync(workbooks, { recursive: true, force: true });
mkdirSync(workbooks, { recursive: true });
const files: Array<{ file: string, category: string, cases: number, savedBy: string }> = [];
const skipped: Record<string, number> = {};
for (const { name, dir } of categories.sort((a, b) => a.name.localeCompare(b.name))) {
  const specs = cases(readFileSync(join(dir, 'testcase.js'), 'utf8'));
  const target = join(workbooks, `${name}.xlsx`);
  const written = await writeCaseWorkbook(target, [ { sheet: name, rows: rowsOf(readFileSync(join(dir, `${name}.js`), 'utf8')), cases: specs } ]);
  for (const [ reason, n ] of Object.entries(written.skipped)) {
    skipped[reason] = (skipped[reason] ?? 0) + n;
  }
  const count = written.blocks.reduce((n, b) => n + b.cases.length, 0);
  console.error(`${name}: ${count} cases`);
  if (!count) {
    continue;
  }
  assertFormatMatchesExtension(target);
  files.push({ file: `${name}.xlsx`, category: name, cases: count, savedBy: 'engine-compatibility (case workbook)' });
}

writeFileSync(join(here, 'manifest.json'), JSON.stringify({
  suite: 'fast-formula-parser',
  upstream: UPSTREAM,
  commit: COMMIT,
  commitDate: date,
  subdir: SUBDIR,
  license: 'MIT',
  shape: 'web (WEBSERVICE, network) excluded; case workbook per category: the category data at A1, one case per row in column AA with the upstream literal in BA and the function name in CA; upstream evaluates each formula at A1, so ROW() and COLUMN() differ',
  skipped,
  importedAt: new Date().toISOString(),
  files,
}, null, 2) + '\n');

console.log(`fast-formula-parser: ${files.length} workbooks, ${files.reduce((n, f) => n + f.cases, 0)} cases at ${COMMIT.slice(0, 8)} (${date})`);
