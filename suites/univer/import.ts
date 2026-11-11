import { execFileSync, spawnSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { basename, dirname, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { type Block, type CaseSpec, LITERAL_OFFSET, type Size, columnLetters, writeCaseWorkbook } from '../../src/case-workbook.ts';
import type { CellValue } from '../../src/model/case.ts';
import { checkoutPinned } from '../../src/pinned-checkout.ts';
import { assertFormatMatchesExtension } from '../../src/workbook-format.ts';

const UPSTREAM = 'https://github.com/dream-num/univer.git';
const COMMIT = '4e6a56e45c8647d9741139af7ebf2f725945d7f3';
const SUBDIR = 'packages/engine-formula';

const here = dirname(fileURLToPath(import.meta.url));
const cache = join(here, '..', '..', '.cache', 'univer');
const pkg = join(cache, SUBDIR);
const workbooks = join(here, 'workbooks');
const captureDir = join(cache, 'capture');

type Scalar = string | number | boolean | null;
type Arg = { kind: 'scalar', value: Scalar } | { kind: 'empty' } | { kind: 'error', value: string } | { kind: 'array', rows: Scalar[][] } | { kind: 'unsupported', what: string };
type Record_ = { file: string, test: string, name: string, args: Arg[], result: Arg };

const CASE_COLUMN = 25;

if (!existsSync(join(cache, '.git'))) {
  mkdirSync(dirname(cache), { recursive: true });
  execFileSync('git', [ 'clone', '--filter=blob:none', '--no-checkout', UPSTREAM, cache ], { stdio: 'inherit' });
}
checkoutPinned(cache, COMMIT);
const date = execFileSync('git', [ '-C', cache, 'log', '-1', '--format=%cI' ], { encoding: 'utf8' }).trim();
if (!existsSync(join(pkg, 'node_modules'))) {
  execFileSync('pnpm', [ 'install', '--filter', '@univerjs/engine-formula...', '--frozen-lockfile' ], { cwd: cache, stdio: 'inherit' });
}

for (const name of [ 'capture.setup.ts', 'vitest.capture.config.ts' ]) {
  copyFileSync(join(here, name), join(pkg, name));
}
rmSync(captureDir, { recursive: true, force: true });
mkdirSync(captureDir, { recursive: true });
const vitest = spawnSync('pnpm', [ 'exec', 'vitest', 'run', '--config', 'vitest.capture.config.ts' ], {
  cwd: pkg,
  encoding: 'utf8',
  env: { ...process.env, UNIVER_CAPTURE_OUT: join(captureDir, 'calls'), CI: 'true' },
});
const vitestSummary = /Tests\s+.*\(\d+\)/.exec(vitest.stdout)?.[0].replace(/\s+/g, ' ') ?? `exit ${vitest.status}`;
console.error(`vitest: ${vitestSummary}`);

const records: Record_[] = readdirSync(captureDir).flatMap(f => readFileSync(join(captureDir, f), 'utf8').split('\n').filter(Boolean).map(line => JSON.parse(line)));

function text (value: Scalar): string {
  if (typeof value === 'string') {
    return `"${value.replace(/"/g, '""')}"`;
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE';
  }
  return value === null ? '' : String(value);
}

function toCellValue (value: Scalar): CellValue | undefined {
  return value === null ? undefined : value;
}

const skipped: Record<string, number> = {};
const skip = (reason: string): void => {
  skipped[reason] = (skipped[reason] ?? 0) + 1;
};

// Univer spills a formula whose result is an array, and its function tests pass arrays to scalar
// arguments and expect an array result, so a case is a dynamic array formula with room to spill
// Univer's result.
function spillOf (result: Arg): Size {
  if (result.kind !== 'array') {
    return { rows: 1, cols: 1 };
  }
  return { rows: Math.max(1, result.rows.length), cols: Math.max(1, ...result.rows.map(r => r?.length ?? 0)) };
}

function blockOf (record: Record_, index: number): Block | undefined {
  const rows: Array<Array<CellValue | undefined>> = [];
  const parts: string[] = [];
  for (const arg of record.args) {
    if (arg.kind === 'unsupported') {
      skip(`argument ${arg.what}`);
      return undefined;
    }
    if (arg.kind === 'array') {
      const argRows = arg.rows.map(r => r ?? []);
      const width = Math.max(0, ...argRows.map(r => r.length));
      if (width >= CASE_COLUMN || argRows.length === 0) {
        skip('array shape');
        return undefined;
      }
      const top = rows.length + 1;
      for (const row of argRows) {
        rows.push(row.map(toCellValue));
      }
      rows.push([]);
      parts.push(`A${top}:${columnLetters(width - 1)}${top + argRows.length - 1}`);
    }
    else {
      parts.push(arg.kind === 'empty' ? '' : arg.kind === 'error' ? arg.value : text(arg.value));
    }
  }
  const result = record.result;
  const spill = spillOf(result);
  if (spill.cols > LITERAL_OFFSET) {
    skip('array wider than the case columns');
    return undefined;
  }
  const expected: CellValue = result.kind === 'scalar' ? result.value : result.kind === 'error' ? result.value : result.kind === 'array' ? result.rows[0]?.[0] ?? null : null;
  const cases: CaseSpec[] = [ { ref: `${columnLetters(CASE_COLUMN)}1`, formula: `${record.name}(${parts.join(',')})`, expected, group: record.name, spill } ];
  const title = record.test.split(' > ').slice(1).join(' ').replace(/[[\]:*?/\\']/g, ' ').replace(/\s+/g, ' ').trim();
  return { sheet: `${index} ${title}`.slice(0, 31).trim(), rows, cases };
}

type FileEntry = { file: string, category: string, blocks: Block[] };
const byFile = new Map<string, FileEntry>();
for (const record of records) {
  const rel = relative(join(pkg, 'src', 'functions'), record.file);
  const category = rel.split('/')[0];
  const stem = rel.split('/')[1];
  const key = `${category}/${stem}.xlsx`;
  const entry = byFile.get(key) ?? { file: key, category, blocks: [] };
  const block = blockOf(record, entry.blocks.length + 1);
  if (block) {
    entry.blocks.push(block);
    byFile.set(key, entry);
  }
}
const entries = [ ...byFile.values() ].sort((a, b) => a.file.localeCompare(b.file));
writeFileSync(join(here, 'cases.json'), JSON.stringify({ commit: COMMIT, files: entries }, null, 1) + '\n');

rmSync(workbooks, { recursive: true, force: true });
const files: Array<{ file: string, category: string, cases: number, savedBy: string }> = [];
for (const { file, category, blocks } of entries) {
  const target = join(workbooks, file);
  mkdirSync(dirname(target), { recursive: true });
  const written = await writeCaseWorkbook(target, blocks);
  for (const [ reason, n ] of Object.entries(written.skipped)) {
    skipped[reason] = (skipped[reason] ?? 0) + n;
  }
  if (!written.blocks.length) {
    continue;
  }
  assertFormatMatchesExtension(target);
  files.push({ file, category, cases: written.blocks.length, savedBy: 'engine-compatibility (case workbook)' });
}

writeFileSync(join(here, 'manifest.json'), JSON.stringify({
  suite: 'univer',
  upstream: UPSTREAM,
  commit: COMMIT,
  commitDate: date,
  subdir: SUBDIR,
  license: 'Apache-2.0',
  vitest: vitestSummary,
  shape: 'captured by running the engine-formula vitest specs with capture.setup.ts, which records every top-level function call as arguments and result into cases.json; one workbook per spec file, one sheet per call: array arguments stacked from A1 and referenced as ranges, scalar arguments inline, the call in Z1 as a dynamic array formula with room to spill Univer\'s result, Univer\'s own result (its top-left value, for an array) as the upstream literal in AZ1 and the function name in BZ1',
  skipped,
  importedAt: new Date().toISOString(),
  files,
}, null, 2) + '\n');

console.error(`skipped: ${JSON.stringify(skipped)}`);
console.log(`univer: ${files.length} workbooks, ${files.reduce((n, f) => n + f.cases, 0)} cases at ${COMMIT.slice(0, 8)} (${date})`);
