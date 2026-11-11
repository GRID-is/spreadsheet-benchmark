import { execFileSync, spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { basename, dirname, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { writeCaseWorkbook } from '../../src/case-workbook.ts';
import { checkoutPinned } from '../../src/pinned-checkout.ts';
import { assertFormatMatchesExtension } from '../../src/workbook-format.ts';
import { type ArrayFormulasByTest, type Skipped, blocksOf } from './blocks.ts';
import { LOCALE_TESTS } from './own-syntax.ts';

const UPSTREAM = 'https://github.com/ironcalc/IronCalc.git';
const COMMIT = '2128db12ab30b77fdebde9cba994273a26438ae3';
const SUBDIR = 'base/src/test';

const here = dirname(fileURLToPath(import.meta.url));
const cache = join(here, '..', '..', '.cache', 'ironcalc');
const workbooks = join(here, 'workbooks');
const capturePath = join(cache, 'capture.jsonl');

if (!existsSync(join(cache, '.git'))) {
  mkdirSync(dirname(cache), { recursive: true });
  execFileSync('git', [ 'clone', '--filter=blob:none', '--sparse', '--no-checkout', UPSTREAM, cache ], { stdio: 'inherit' });
}
// The capture builds the crate, and Cargo loads every member of the workspace.
execFileSync('git', [ '-C', cache, 'sparse-checkout', 'add', 'base', 'xlsx', 'bindings' ], { stdio: 'inherit' });
checkoutPinned(cache, COMMIT);
const date = execFileSync('git', [ '-C', cache, 'log', '-1', '--format=%cI' ], { encoding: 'utf8' }).trim();

// A formula entered with set_user_input, as _set does, is a dynamic array formula when IronCalc's
// static analysis cannot prove it scalar, and spills its result; Excel evaluates an ordinary formula
// with implicit intersection instead. To tell which formulas are dynamic array formulas, and how far
// they spill, run the tests with _get_text patched to record the array formulas of its sheet.
const utilPath = join(cache, 'base', 'src', 'test', 'util.rs');
const util = readFileSync(utilPath, 'utf8');
const getText = '        let CellReferenceIndex { sheet, row, column } = self._parse_reference(cell);\n        self._get_text_at(sheet, row, column)\n';
if (util.split(getText).length !== 2) {
  throw new Error('util.rs patch anchor not found exactly once; update suites/ironcalc-unit/import.ts for the pinned commit');
}
writeFileSync(utilPath, util.replace(getText, getText.replace('\n', '\n        _capture_array_formulas(self, sheet);\n')) + `
fn _capture_array_formulas(model: &Model, sheet: u32) {
    let Ok(path) = std::env::var("IRONCALC_CAPTURE_OUT") else { return };
    let Ok(worksheet) = model.workbook.worksheet(sheet) else { return };
    let mut anchors = Vec::new();
    for (row, column, cell) in worksheet.sheet_data.cells() {
        if let Cell::ArrayFormula { r, kind, .. } = &*cell {
            anchors.push(format!("[{row},{column},{},{},\\"{kind:?}\\"]", r.0, r.1));
        }
    }
    let test = std::thread::current().name().unwrap_or("").to_string();
    let line = format!("{{\\"test\\":\\"{test}\\",\\"sheet\\":{sheet},\\"anchors\\":[{}]}}\\n", anchors.join(","));
    use std::io::Write;
    if let Ok(mut file) = std::fs::OpenOptions::new().create(true).append(true).open(path) {
        let _ = file.write_all(line.as_bytes());
    }
}
`);
writeFileSync(capturePath, '');
let run;
try {
  run = spawnSync('cargo', [ 'test', '--quiet', '-p', 'ironcalc_base', '--lib', 'test::' ], {
    cwd: cache,
    encoding: 'utf8',
    maxBuffer: 64 * 1024 * 1024,
    env: { ...process.env, IRONCALC_CAPTURE_OUT: capturePath },
  });
}
finally {
  writeFileSync(utilPath, util);
}
if (run.error) {
  throw new Error(`cargo test could not run (the import needs a Rust toolchain): ${run.error.message}`);
}
console.error(`cargo test: ${/test result: .*/.exec(run.stdout)?.[0] ?? `exit ${run.status}`}`);
if (run.status !== 0) {
  console.error(`${run.stdout}${run.stderr}`.split('\n').slice(-15).join('\n'));
  console.error('warning: test run exited non-zero; continuing with whatever was captured');
}

// The array formulas of each test, by its path (test::<module>::<fn>), from its last _get_text
const arrayFormulas: ArrayFormulasByTest = new Map();
for (const line of readFileSync(capturePath, 'utf8').split('\n')) {
  if (line) {
    const { test, sheet, anchors } = JSON.parse(line) as { test: string, sheet: number, anchors: Array<[number, number, number, number, string]> };
    arrayFormulas.set(test, { sheet, anchors: anchors.map(([ row, column, width, height, kind ]) => ({ row, column, width, height, kind })) });
  }
}

function walk (dir: string): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap(entry => {
    const path = join(dir, entry.name);
    return entry.isDirectory() ? walk(path) : entry.name.startsWith('test_') && entry.name.endsWith('.rs') ? [ path ] : [];
  });
}

rmSync(workbooks, { recursive: true, force: true });
const files: Array<{ file: string, category: string, cases: number, savedBy: string }> = [];
const skipped: Skipped = {};
for (const path of walk(join(cache, SUBDIR)).sort()) {
  const rel = relative(join(cache, SUBDIR), path);
  const category = rel.includes('/') ? rel.split('/')[0] : 'general';
  const stem = basename(rel, '.rs').replace(/^test_/, '');
  const module = `test::${rel.replace(/\.rs$/, '').replaceAll('/', '::')}`;
  const blocks = blocksOf(readFileSync(path, 'utf8'), module, LOCALE_TESTS.has(basename(path)), skipped, arrayFormulas);
  if (!blocks.length) {
    continue;
  }
  mkdirSync(join(workbooks, category), { recursive: true });
  const target = join(workbooks, category, `${stem}.xlsx`);
  const written = await writeCaseWorkbook(target, blocks);
  for (const [ reason, n ] of Object.entries(written.skipped)) {
    skipped[reason] = (skipped[reason] ?? 0) + n;
  }
  if (!written.blocks.length) {
    continue;
  }
  assertFormatMatchesExtension(target);
  const count = written.blocks.reduce((n, b) => n + b.cases.length, 0);
  files.push({ file: `${category}/${stem}.xlsx`, category, cases: count, savedBy: 'engine-compatibility (case workbook)' });
}

writeFileSync(join(here, 'manifest.json'), JSON.stringify({
  suite: 'ironcalc-unit',
  upstream: UPSTREAM,
  commit: COMMIT,
  commitDate: date,
  subdir: SUBDIR,
  license: 'MIT OR Apache-2.0',
  shape: 'case workbook per Rust test file, one sheet per #[test] fn: every model._set input in its cell, the _get_text expectation (formatted text, parsed to number, boolean or error where it reads as one) 26 columns right of the asserted formula cell and the fn name 52 columns right; a formula that IronCalc evaluates as a dynamic array formula, captured by running the tests with _get_text patched, is one over the range it spills into; tests that call other model methods, evaluate more than once, or address other sheets are left out, as are tests of formulas in a German locale and of LAMBDA parameters in brackets, which are IronCalc syntax',
  skipped,
  importedAt: new Date().toISOString(),
  files,
}, null, 2) + '\n');

console.error(`skipped: ${JSON.stringify(skipped)}`);
console.log(`ironcalc-unit: ${files.length} workbooks, ${files.reduce((n, f) => n + f.cases, 0)} cases at ${COMMIT.slice(0, 8)} (${date})`);
