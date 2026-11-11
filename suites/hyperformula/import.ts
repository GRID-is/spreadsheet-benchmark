import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { runInNewContext } from 'node:vm';
import { type Block, type CaseSpec, LITERAL_OFFSET, type Size, columnLetters, parseRef, writeCaseWorkbook } from '../../src/case-workbook.ts';
import type { CellValue } from '../../src/model/case.ts';
import { checkoutPinned } from '../../src/pinned-checkout.ts';
import { assertFormatMatchesExtension } from '../../src/workbook-format.ts';
import { withoutArrayFormula } from './array-formula.ts';

const UPSTREAM = 'https://github.com/handsontable/hyperformula.git';
// The last release with public tests: from 3.2.0, test/ holds only a smoke test, and the full suite
// is available on request
const COMMIT = 'd30eeab3fadbb88d7bdede4ee12478baf5d0a24f';
const SUBDIR = 'test/unit/interpreter';
// The release tagged at COMMIT, run at import to learn which formulas HyperFormula evaluates as
// arrays and the size of their results
const HF_VERSION = '3.1.1';

const here = dirname(fileURLToPath(import.meta.url));
const cache = join(here, '..', '..', '.cache', 'hyperformula');
const workbooks = join(here, 'workbooks');
const engineDir = join(here, '..', '..', '.cache', `hyperformula-${HF_VERSION}`);

const ERROR_STRINGS: Record<string, string> = {
  DIV_BY_ZERO: '#DIV/0!', NAME: '#NAME?', VALUE: '#VALUE!', NUM: '#NUM!', NA: '#N/A',
  CYCLE: '#CYCLE!', REF: '#REF!', SPILL: '#SPILL!', LIC: '#LICENSE!', ERROR: '#ERROR!',
};
const MUTATING = /\b(setCellContents|setSheetContent|addSheet|removeSheet|addRows|removeRows|addColumns|removeColumns|buildFromSheets|buildEmpty|undo|redo|moveCells|moveRows|moveColumns|paste|copy|cut|batch|setNamedExpression|addNamedExpression|rebuildAndRecalculate|suspendEvaluation)\b/;

if (!existsSync(join(cache, '.git'))) {
  mkdirSync(dirname(cache), { recursive: true });
  execFileSync('git', [ 'clone', '--filter=blob:none', '--sparse', '--no-checkout', UPSTREAM, cache ], { stdio: 'inherit' });
}
// Set on every import, so that a cache cloned for an older SUBDIR follows it
execFileSync('git', [ '-C', cache, 'sparse-checkout', 'set', SUBDIR ], { stdio: 'inherit' });
checkoutPinned(cache, COMMIT);
const date = execFileSync('git', [ '-C', cache, 'log', '-1', '--format=%cI' ], { encoding: 'utf8' }).trim();
if (!existsSync(join(engineDir, 'node_modules', 'hyperformula'))) {
  mkdirSync(engineDir, { recursive: true });
  execFileSync('npm', [ 'install', '--prefix', engineDir, '--no-save', '--no-audit', '--no-fund', `hyperformula@${HF_VERSION}` ], { stdio: 'inherit' });
}

type Address = { sheet: number, row: number, col: number };
type Engine = {
  getCellType (address: Address): string,
  getCellValue (address: Address): unknown,
  // Internal API, from which the size of an array result is read
  dependencyGraph: { getCell (address: Address): { width?: number, height?: number } | undefined },
  destroy (): void,
};
// A dynamic import by path, typed here, because the release is installed in .cache at import time:
// it is not a dependency of this repository.
const { HyperFormula } = await import(pathToFileURL(join(engineDir, 'node_modules', 'hyperformula', 'es', 'index.mjs')).href) as {
  HyperFormula: { version: string, buildFromArray (rows: unknown[][], config: object): Engine },
};
if (HyperFormula.version !== HF_VERSION) {
  throw new Error(`${engineDir} holds HyperFormula ${HyperFormula.version}, not ${HF_VERSION}`);
}

function balanced (source: string, start: number, open: string, close: string): string | undefined {
  let depth = 0;
  let quote: string | undefined;
  for (let i = start; i < source.length; i++) {
    const ch = source[i];
    if (quote) {
      if (ch === '\\') {
        i++;
      }
      else if (ch === quote) {
        quote = undefined;
      }
      continue;
    }
    if (ch === '\'' || ch === '"' || ch === '`') {
      quote = ch;
    }
    else if (ch === open) {
      depth++;
    }
    else if (ch === close && --depth === 0) {
      return source.slice(start, i + 1);
    }
  }
  return undefined;
}

const context = {
  EmptyValue: null,
  ErrorType: Object.fromEntries(Object.keys(ERROR_STRINGS).map(k => [ k, k ])),
  ErrorMessage: new Proxy({}, { get: () => '' }),
  detailedError: (type: string) => ERROR_STRINGS[type],
  detailedErrorWithOrigin: (type: string) => ERROR_STRINGS[type],
};

function evaluate (expression: string): unknown {
  return runInNewContext(`(${expression})`, { ...context });
}

function toCellValue (raw: unknown): CellValue | undefined {
  if (raw === null || raw === undefined) {
    return undefined;
  }
  if (typeof raw === 'number' || typeof raw === 'boolean') {
    return raw;
  }
  if (typeof raw !== 'string') {
    throw new Error(`unsupported array value ${JSON.stringify(raw)}`);
  }
  if (raw === '') {
    return undefined;
  }
  if (raw.startsWith('\'')) {
    return raw.slice(1);
  }
  if (/^-?\d+(\.\d+)?(e[+-]?\d+)?$/i.test(raw)) {
    return Number(raw);
  }
  if (/^(true|false)$/i.test(raw)) {
    return raw.toLowerCase() === 'true';
  }
  return raw;
}

function literal (raw: unknown): CellValue {
  if (raw === null || raw === undefined || typeof raw === 'number' || typeof raw === 'boolean' || typeof raw === 'string') {
    return raw ?? null;
  }
  throw new Error(`unsupported expectation ${JSON.stringify(raw)}`);
}

function sheetName (index: number, title: string): string {
  return `${index} ${title.replace(/[[\]:*?/\\']/g, ' ').replace(/\s+/g, ' ').trim()}`.slice(0, 31).trim();
}

// The formulas of rows that HyperFormula evaluates as arrays, by reference, with the size of their
// results; undefined when other cells of rows block the result of one.
function arraysOf (rows: unknown[][], useArrayArithmetic: boolean): Record<string, Size> | undefined {
  const engine = HyperFormula.buildFromArray(rows, { licenseKey: 'gpl-v3', useArrayArithmetic });
  try {
    const arrays: Record<string, Size> = {};
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].length; c++) {
        const address = { sheet: 0, row: r, col: c };
        if (engine.getCellType(address) !== 'ARRAYFORMULA') {
          continue;
        }
        const result = engine.getCellValue(address);
        if (typeof result === 'object' && result !== null && 'type' in result && result.type === 'SPILL') {
          return undefined;
        }
        const { width = 1, height = 1 } = engine.dependencyGraph.getCell(address) ?? {};
        arrays[`${columnLetters(c)}${r + 1}`] = { rows: height, cols: width };
      }
    }
    return arrays;
  }
  finally {
    engine.destroy();
  }
}

type Skipped = Record<string, number>;

function blocksOf (source: string, group: string, skipped: Skipped): Block[] {
  const blocks: Block[] = [];
  const starts = [ ...source.matchAll(/\bit\((['`])((?:\\.|(?!\1).)*)\1/g) ];
  starts.forEach((start, i) => {
    const body = source.slice(start.index, starts[i + 1]?.index ?? source.length);
    const builds = [ ...body.matchAll(/buildFromArray\(/g) ];
    const skip = (reason: string): void => {
      skipped[reason] = (skipped[reason] ?? 0) + 1;
    };
    if (builds.length !== 1 || MUTATING.test(body)) {
      skip(builds.length !== 1 ? 'not one buildFromArray' : 'mutating api');
      return;
    }
    const arrayStart = (builds[0].index ?? 0) + 'buildFromArray('.length;
    const arrayLiteral = balanced(body, arrayStart, '[', ']');
    let rawRows: unknown;
    try {
      rawRows = arrayLiteral ? evaluate(arrayLiteral) : undefined;
    }
    catch {
      rawRows = undefined;
    }
    if (!Array.isArray(rawRows) || !rawRows.every(Array.isArray)) {
      skip('array not a literal');
      return;
    }
    let rows: Array<Array<CellValue | undefined>>;
    try {
      rows = rawRows.map(row => row.map(toCellValue));
    }
    catch {
      skip('array value');
      return;
    }
    const cases: CaseSpec[] = [];
    for (const m of body.matchAll(/expect\(engine\.getCellValue\(adr\('([A-Z]+\d+)'\)\)\)\.(toEqual|toBe|toBeCloseTo|toEqualError)\(/g)) {
      const ref = m[1];
      const args = balanced(body, (m.index ?? 0) + m[0].length - 1, '(', ')');
      const { col, row } = { col: ref.replace(/\d+$/, ''), row: Number(ref.replace(/^[A-Z]+/, '')) };
      const colIndex = [ ...col ].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
      const arrayValue = rows[row - 1]?.[colIndex];
      if (typeof arrayValue !== 'string' || !arrayValue.startsWith('=') || !args) {
        skip('expectation not on a formula');
        continue;
      }
      if (cases.some(c => c.ref === ref)) {
        continue;
      }
      try {
        const values = evaluate(`[${args.slice(1, -1)}]`);
        cases.push({ ref, expected: literal(Array.isArray(values) ? values[0] : values), group });
      }
      catch {
        skip('expectation value');
      }
    }
    if (!cases.length) {
      return;
    }
    const rest = body.slice(arrayStart + (arrayLiteral ?? '').length);
    const config = /^\s*,\s*\{/.test(rest) ? balanced(rest, rest.indexOf('{'), '{', '}') : undefined;
    const arrays = arraysOf(rawRows, /\buseArrayArithmetic\s*:\s*true\b/.test(config ?? ''));
    if (arrays === undefined) {
      skip('array result blocked');
      return;
    }
    if (Object.entries(arrays).some(([ ref, { cols } ]) => parseRef(ref).col + cols > LITERAL_OFFSET)) {
      skip('spills into the literal columns');
      return;
    }
    const spills: Record<string, Size> = {};
    rows.forEach((row, r) => row.forEach((value, c) => {
      const ref = `${columnLetters(c)}${r + 1}`;
      const translated = typeof value === 'string' && value.startsWith('=') ? withoutArrayFormula(value) : undefined;
      if (translated !== undefined) {
        row[c] = translated;
      }
      const spill = arrays[ref] ?? (translated !== undefined ? { rows: 1, cols: 1 } : undefined);
      if (spill) {
        spills[ref] = spill;
      }
    }));
    blocks.push({ sheet: sheetName(blocks.length + 1, start[2]), rows, cases, spills });
  });
  return blocks;
}

rmSync(workbooks, { recursive: true, force: true });
const files: Array<{ file: string, category: string, cases: number, savedBy: string }> = [];
const skipped: Skipped = {};
for (const spec of readdirSync(join(cache, SUBDIR)).filter(f => f.endsWith('.spec.ts')).sort()) {
  const stem = spec.replace(/\.spec\.ts$/, '');
  const category = stem.startsWith('function-') ? 'functions' : 'other';
  const group = stem.replace(/^function-/, '').toUpperCase();
  const blocks = blocksOf(readFileSync(join(cache, SUBDIR, spec), 'utf8'), group, skipped);
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
  suite: 'hyperformula',
  upstream: UPSTREAM,
  commit: COMMIT,
  commitDate: date,
  subdir: SUBDIR,
  license: 'GPL-3.0-only',
  shape: `case workbook per spec file, one sheet per it block: the buildFromArray array at A1 with its formulas, the upstream literal 26 columns right of each asserted cell and the spec name 52 columns right; formulas that HyperFormula ${HF_VERSION} evaluates as arrays are dynamic array formulas, as is a formula with ARRAYFORMULA(x), written with x in its place; it blocks that use more than one buildFromArray, other sheets, or mutation APIs are left out`,
  skipped,
  importedAt: new Date().toISOString(),
  files,
}, null, 2) + '\n');

console.error(`skipped: ${JSON.stringify(skipped)}`);
console.log(`hyperformula: ${files.length} workbooks, ${files.reduce((n, f) => n + f.cases, 0)} cases at ${COMMIT.slice(0, 8)} (${date})`);
