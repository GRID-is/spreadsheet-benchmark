import { type CellValue as RecordedValue, type Spill } from './model/case.ts';
import { areaRange, areaRefs, isSingleCell, spillArea } from './model/spill.ts';

export const DEFAULT_EPS = 5e-8;

export function numbersAreClose (x: number, y: number, eps: number): boolean {
  const norm = Math.sqrt(x * x + y * y);
  if (norm === 0) {
    return true;
  }
  const d = Math.abs(x - y);
  return d < eps || d / norm < eps;
}

export type CellValue = string | number | boolean | null | undefined;

export function valuesMatch (expected: CellValue, actual: CellValue, eps: number): boolean {
  if (typeof expected === 'number' && typeof actual === 'number') {
    return numbersAreClose(expected, actual, eps);
  }
  if (expected == null || expected === '') {
    return actual == null || actual === '';
  }
  return expected === actual;
}

/**
 * The spill of the case at `anchor`: Excel's spill range there and the engine's, with the values of
 * both over Excel's range. `excel.range` is the range of the fixture's array formula at `anchor`,
 * absent where it has none. `engine.range` is undefined where the engine reports no spill range.
 * Undefined where neither range spans more than `anchor`: the case is then its cell alone.
 */
export function observeSpill (
  anchor: string,
  excel: { range?: string | null, value: (ref: string) => CellValue },
  engine: { range?: string | null, value: (ref: string) => CellValue },
  eps: number,
): Spill | undefined {
  const excelArea = spillArea(anchor, excel.range);
  const engineArea = engine.range === undefined ? undefined : spillArea(anchor, engine.range);
  if (isSingleCell(excelArea) && (engineArea === undefined || isSingleCell(engineArea))) {
    return undefined;
  }
  const refs = areaRefs(excelArea);
  const read = (value: (ref: string) => CellValue) => refs.map(row => row.map((ref): RecordedValue => value(ref) ?? null));
  const spill: Spill = { excelRange: areaRange(excelArea), excelValues: read(excel.value) };
  if (engineArea !== undefined) {
    spill.apiaryRange = areaRange(engineArea);
  }
  const apiaryValues = read(engine.value);
  spill.apiaryValues = apiaryValues;
  spill.mismatches = spill.excelValues.flatMap((row, r) => row.filter((v, c) => (r > 0 || c > 0) && !valuesMatch(v, apiaryValues[r]?.[c], eps))).length;
  return spill;
}

const FUNCTION_CALL = /(?<![A-Z0-9_.!'])([A-Z_][A-Z0-9_.]*)\s*\(/gi;

export function functionNames (formula: string): string[] {
  const names = new Set<string>();
  for (const match of formula.matchAll(FUNCTION_CALL)) {
    names.add((match[1] ?? '').toUpperCase().replace(/^_XLFN\./, ''));
  }
  return [ ...names ];
}

const VOLATILE = new Set([ 'RAND', 'RANDBETWEEN', 'RANDARRAY', 'NOW', 'TODAY' ]);

export function isVolatile (formula: string): boolean {
  return functionNames(formula).some(name => VOLATILE.has(name));
}
