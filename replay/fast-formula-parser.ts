// Replay a report's cases through fast-formula-parser. FFP is a per-formula
// evaluator, not a workbook engine: each case formula is evaluated against a
// data provider serving Excel's cached values for every other cell, so there
// is no recalculation chain and no cross-cell error propagation of FFP's own.
// Excel's spilled cells are the exception: the cell input leaves them out, so
// the provider serves them empty.
//
// An array formula, dynamic or CSE, is evaluated with allowReturnArray, and its
// value is the top-left element of the result, as in Excel. FFP keeps no sheet
// to spill into, so the reader answers a spilled cell from its anchor's result
// and reports the result's size as the spill range. Any other formula is
// evaluated without allowReturnArray, by FFP's own scalar rule: a single-column
// range answers its first cell and any other range #VALUE!, where Excel answers
// the implicit intersection with the formula's cell. Its spill range is its own
// cell.
import { createRequire } from 'node:module';
import FormulaParser from 'fast-formula-parser';
import { areaRange, areaRefs, spillArea } from '../src/model/spill.ts';
import { readCellInput, refToCoords, replay, type CaseCell, type Engine, type EngineValue, type InputCell } from './common.ts';

const require = createRequire(import.meta.url);
const version: string = require('fast-formula-parser/package.json').version;
const { FormulaError } = FormulaParser;

type ProviderValue = InstanceType<typeof FormulaError> | InputCell['v'];

function toEngineValue (v: InputCell['v']): ProviderValue {
  if (typeof v === 'string' && /^#[A-Z0-9/!?._]+[!?]$/.test(v)) return new FormulaError(v);
  return v ?? null;
}

// Whether v is an error answer from FFP: a FormulaError, or any object with a string `error`
// property.
function isFormulaError (v: unknown): v is { error: string, message?: unknown } {
  return v instanceof FormulaError || (v != null && typeof v === 'object' && 'error' in v && typeof v.error === 'string');
}

// A case's value as the replay records it.
function toCaseValue (v: unknown): EngineValue {
  while (Array.isArray(v)) v = v[0]?.[0] ?? null;
  if (isFormulaError(v)) {
    // FFP reports unimplemented functions as #ERROR!; Excel would say #NAME?,
    // and the classifier counts #NAME? as unsupported rather than fail.
    return /is not implemented/.test(String(v.message ?? '')) ? '#NAME?' : v.error;
  }
  if (v instanceof Error) throw v;
  if (v != null && typeof v === 'object') return String(v);
  return typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean' ? v : null;
}

// FFP's answer as an array by row: a scalar is a 1x1 array.
function toRows (v: unknown): unknown[][] {
  if (!Array.isArray(v)) return [ [ v ] ];
  const rows = v.map(row => Array.isArray(row) ? row : [ row ]);
  return rows.length > 0 && rows[0].length > 0 ? rows : [ [ null ] ];
}

export const engine = {
  name: 'fast-formula-parser',
  version,
  async buildFile (fixture: string) {
    const sheets = new Map<string, Map<string, ProviderValue>>();
    // The fixture's formula text of each formula cell and the array formulas among them, by
    // `sheet!ref`, and the anchor of each other cell of an array formula's range.
    const formulas = new Map<string, string>();
    const arrayFormulas = new Set<string>();
    const anchors = new Map<string, string>();
    for (const sheet of await readCellInput(fixture)) {
      const values = new Map<string, ProviderValue>();
      for (const [ref, cell] of Object.entries(sheet.cells)) {
        const { row, col } = refToCoords(ref);
        values.set(`${row + 1},${col + 1}`, toEngineValue(cell.v));
        if (cell.f == null) continue;
        formulas.set(`${sheet.name}!${ref}`, String(cell.f));
        if (cell.F != null) {
          arrayFormulas.add(`${sheet.name}!${ref}`);
          for (const other of areaRefs(spillArea(ref, cell.F)).flat()) {
            if (other !== ref) anchors.set(`${sheet.name}!${other}`, ref);
          }
        }
      }
      sheets.set(sheet.name, values);
    }
    const cellAt = (sheet: string, row: number, col: number) => sheets.get(sheet)?.get(`${row},${col}`) ?? null;
    const parser = new FormulaParser({
      onCell: ({ sheet, row, col }) => cellAt(sheet, row, col),
      onRange: ({ sheet, from, to }) => {
        const rows = [];
        for (let r = from.row; r <= to.row; r++) {
          const line = [];
          for (let c = from.col; c <= to.col; c++) line.push(cellAt(sheet, r, c));
          rows.push(line);
        }
        return rows;
      },
    });
    // FFP's answer to the formula at each cell read, by `sheet!ref`, as an array by row: an array
    // formula's whole array, any other formula's 1x1 scalar.
    const answers = new Map<string, unknown[][]>();
    const answer = (sheet: string, ref: string, formula: string | undefined): unknown[][] | undefined => {
      const key = `${sheet}!${ref}`;
      let rows = answers.get(key);
      if (rows === undefined) {
        const f = formula ?? formulas.get(key);
        if (f === undefined) return undefined;
        const { row, col } = refToCoords(ref);
        let v: unknown;
        try {
          v = parser.parse(f, { sheet, row: row + 1, col: col + 1 }, arrayFormulas.has(key));
        }
        catch (err) {
          v = err;
        }
        rows = arrayFormulas.has(key) ? toRows(v) : [ [ v ] ];
        answers.set(key, rows);
      }
      return rows;
    };
    // A case's cell answers the top-left of its formula's array, and a spilled cell the element of
    // its anchor's array at its offset; a cell outside the array answers nothing. Each formula is
    // evaluated once: from its case's formula text when the case is read first, as the job process
    // does, else from the fixture's.
    const value = (c: CaseCell & { formula?: string }): EngineValue => {
      const own = answer(c.sheet, c.ref, c.formula);
      if (own !== undefined) return toCaseValue(own[0][0]);
      const anchor = anchors.get(`${c.sheet}!${c.ref}`);
      const rows = anchor === undefined ? undefined : answer(c.sheet, anchor, undefined);
      if (anchor === undefined || rows === undefined) return null;
      const at = refToCoords(c.ref);
      const origin = refToCoords(anchor);
      return toCaseValue(rows[at.row - origin.row]?.[at.col - origin.col] ?? null);
    };
    // The spill range is the size of the array that FFP answered, even for a CSE formula, whose
    // array Excel fits to the formula's range: the replay does not yet read the cell input's cse
    // mark, and fitting every array to its range would hide a spill of the wrong size.
    const spillRange = (c: CaseCell): string | undefined => {
      const rows = answer(c.sheet, c.ref, undefined);
      if (rows === undefined) return undefined;
      const { row, col } = refToCoords(c.ref);
      return areaRange({ top: row, left: col, bottom: row + rows.length - 1, right: col + rows.reduce((width, r) => Math.max(width, r.length), 0) - 1 });
    };
    return Object.assign(value, { spillRange });
  },
} satisfies Engine;

if (import.meta.main) {
  await replay(engine, import.meta.url);
}
