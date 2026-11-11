// Replay a report's cases through HyperFormula: rebuild each fixture's cells
// (formulas + literals) in an HF instance and read the case cells back.
//
// An array formula of the fixture, dynamic or CSE, is given to HyperFormula in ARRAYFORMULA, which
// evaluates its expression as an array, as Excel does. Other formulas are given as they are, and
// HyperFormula's default config (useArrayArithmetic false) intersects a range in a scalar argument
// with the formula's row or column, as Excel does in an ordinary formula. HyperFormula spills a CSE
// formula's array at its own size, where Excel fits it to the formula's range.
import { HyperFormula } from 'hyperformula';
import { harnessTiming, readCellInput, refToCoords, replay, type CaseCell, type Engine, type InputCell, type InputSheet } from './common.ts';

type Content = string | InputCell['v'];

// HyperFormula's content for each sheet of a fixture's cell input: each array formula in
// ARRAYFORMULA, unless wrap returns false for its cell, and every other formula as it is.
export function hyperFormulaSheets (input: InputSheet[], wrap = (_sheet: string, _ref: string) => true): Record<string, Content[][]> {
  const sheets: Record<string, Content[][]> = {};
  for (const sheet of input) {
    const rows: Content[][] = [];
    for (const [ref, cell] of Object.entries(sheet.cells)) {
      const { row, col } = refToCoords(ref);
      (rows[row] ??= [])[col] = cell.f == null ? cell.v : cell.F != null && wrap(sheet.name, ref) ? `=ARRAYFORMULA(${cell.f})` : '=' + cell.f;
    }
    for (let i = 0; i < rows.length; i++) rows[i] ??= [];
    sheets[sheet.name] = rows;
  }
  return sheets;
}

const CONFIG = { licenseKey: 'gpl-v3' };

// Builds a HyperFormula workbook from a fixture's cell input, with every array formula in
// ARRAYFORMULA. HyperFormula 3.5.0 throws while building a workbook that has some formulas in
// ARRAYFORMULA, such as FILTER(A1:A2, {TRUE;FALSE})+30 or an XLOOKUP of an array, instead of
// answering an error in their cell. When the build throws, the workbook is built once per array
// formula with only that one in ARRAYFORMULA, and then once more with the formulas that threw given
// as ordinary formulas. If none threw on its own, the first build's error is thrown. Only that last
// build counts as HyperFormula's time: the builds before it are the harness's.
export function buildHyperFormula (input: InputSheet[]): HyperFormula {
  const t0 = Date.now();
  try {
    return HyperFormula.buildFromSheets(hyperFormulaSheets(input), CONFIG);
  }
  catch (error) {
    const throwing = new Set<string>();
    for (const sheet of input) {
      for (const [ ref, cell ] of Object.entries(sheet.cells)) {
        if (cell.f == null || cell.F == null) continue;
        try {
          HyperFormula.buildFromSheets(hyperFormulaSheets(input, (s, r) => s === sheet.name && r === ref), CONFIG).destroy();
        }
        catch {
          throwing.add(`${sheet.name}!${ref}`);
        }
      }
    }
    if (throwing.size === 0) throw error;
    harnessTiming.ms += Date.now() - t0;
    return HyperFormula.buildFromSheets(hyperFormulaSheets(input, (s, r) => !throwing.has(`${s}!${r}`)), CONFIG);
  }
}

export const engine = {
  name: 'hyperformula',
  version: HyperFormula.version,
  async buildFile (fixture: string) {
    const hf = buildHyperFormula(await readCellInput(fixture));
    return (c: CaseCell) => {
      const sheetId = hf.getSheetId(c.sheet);
      if (sheetId == null) return null;
      const { row, col } = refToCoords(c.ref);
      let v = hf.getCellValue({ sheet: sheetId, row, col });
      if (v != null && typeof v === 'object') v = v.value ?? String(v);
      return v;
    };
  },
} satisfies Engine;

if (import.meta.main) {
  await replay(engine, import.meta.url);
}
