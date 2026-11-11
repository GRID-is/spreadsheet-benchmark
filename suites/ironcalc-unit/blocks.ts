import { type Block, type CaseSpec, LITERAL_OFFSET, type Size, columnLetters, parseRef } from '../../src/case-workbook.ts';
import type { CellValue } from '../../src/model/case.ts';
import { hasBracketedLambdaParameter } from './own-syntax.ts';

export type ArrayFormula = { row: number, column: number, width: number, height: number, kind: string };

export type ArrayFormulasByTest = Map<string, { sheet: number, anchors: ArrayFormula[] }>;

function rustString (raw: string): string {
  return JSON.parse(`"${raw.replace(/\\\n\s*/g, '')}"`);
}

function input (text: string): CellValue {
  if (/^-?\d+(\.\d+)?(e[+-]?\d+)?$/i.test(text)) {
    return Number(text);
  }
  if (/^(true|false)$/i.test(text)) {
    return text.toLowerCase() === 'true';
  }
  return text === '' ? null : text;
}

function expectation (text: string): CellValue {
  return /^#[A-Z0-9/!?]+$/.test(text) ? text : input(text);
}

export type Skipped = Record<string, number>;

export function blocksOf (source: string, module: string, localeTest: boolean, skipped: Skipped, arrayFormulas: ArrayFormulasByTest): Block[] {
  const blocks: Block[] = [];
  // Each sheet is numbered among the tests that have cases, the skipped ones included, so that
  // skipping a test renames no other sheet, which would call for a new Excel pass.
  let tests = 0;
  const starts = [ ...source.matchAll(/#\[test\]\s*\nfn (\w+)\(\)/g) ];
  starts.forEach((start, i) => {
    const body = source.slice(start.index, starts[i + 1]?.index ?? source.length);
    const skip = (reason: string, n = 1): void => {
      skipped[reason] = (skipped[reason] ?? 0) + n;
    };
    // Also a call written as a chain, with `model` and `.method(` on separate lines
    const methods = [ ...body.matchAll(/\bmodel\s*\.(\w+)\(/g) ].map(m => m[1]);
    const foreign = methods.filter(m => ![ '_set', '_get_text', '_get_formula', 'evaluate' ].includes(m));
    if (foreign.length || methods.filter(m => m === 'evaluate').length !== 1) {
      skip(foreign.length ? 'other api' : 'not one evaluate');
      return;
    }
    const sets = [ ...body.matchAll(/model\._set\(\s*"([^"]+)",\s*(?:r#"([^]*?)"#|"((?:\\.|\\\n|[^"\\])*)")\s*,?\s*\)/g) ]
      .map(m => [ m[1], m[2] ?? rustString(m[3]) ] as const);
    if (sets.some(([ ref ]) => !/^[A-Z]+\d+$/.test(ref))) {
      skip('other sheet');
      return;
    }
    const rows: Array<Array<CellValue | undefined>> = [];
    for (const [ ref, text ] of sets) {
      const { col, row } = parseRef(ref);
      if (col >= 26) {
        skip('set cells too wide');
        return;
      }
      (rows[row - 1] ??= [])[col] = input(text);
    }
    const cases: CaseSpec[] = [];
    for (const [ , ref, raw ] of body.matchAll(/_get_text\("([A-Z]+\d+)"\),\s*\*?"((?:\\.|[^"\\])*)"/g)) {
      const { col, row } = parseRef(ref);
      const cell = rows[row - 1]?.[col];
      if (typeof cell !== 'string' || !cell.startsWith('=')) {
        skip('expectation not on a formula');
        continue;
      }
      if (!cases.some(c => c.ref === ref)) {
        cases.push({ ref, expected: expectation(rustString(raw)), group: start[1] });
      }
    }
    if (!cases.length) {
      return;
    }
    const number = ++tests;
    if (localeTest) {
      skip('upstream locale test, not an Excel formula', cases.length);
      return;
    }
    if (sets.some(([ , text ]) => hasBracketedLambdaParameter(text))) {
      skip('not Excel syntax', cases.length);
      return;
    }
    const captured = arrayFormulas.get(`${module}::${start[1]}`);
    if (!captured || captured.sheet !== 0) {
      throw new Error(`${module}::${start[1]}: no array formulas captured on its first sheet`);
    }
    const spills: Record<string, Size> = {};
    for (const { row, column, width, height, kind } of captured.anchors) {
      const ref = `${columnLetters(column - 1)}${row}`;
      if (kind !== 'Dynamic') {
        throw new Error(`${module}::${start[1]}!${ref}: ${kind} array formula without set_user_array_formula`);
      }
      if (column - 1 + width > LITERAL_OFFSET) {
        skip('spills into the literal columns');
        return;
      }
      const spill = { rows: height, cols: width };
      const spec = cases.find(c => c.ref === ref);
      if (spec) {
        spec.spill = spill;
      }
      else {
        spills[ref] = spill;
      }
    }
    blocks.push({ sheet: `${number} ${start[1]}`.slice(0, 31), rows: rows.map(row => row ?? []), cases, spills });
  });
  return blocks;
}
