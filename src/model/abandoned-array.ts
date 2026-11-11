import { callArguments } from './formula-text.ts';

// The most elements an array function's result has been measured to compute in Excel. At
// 134,209,536 Excel abandoned the calculation and the cell showed 0, with no error.
const LARGEST_COMPUTED = 100_010_000;

// SEQUENCE and RANDARRAY answer #VALUE! for a count past this, of rows or of columns, whatever the
// array's size.
const COUNT_CAP = 1_048_576;

// A number, signed or not, in any parentheses.
const NUMERIC_CONSTANT = /^[(\s]*([-+]?\s*(?:\d+\.?\d*|\.\d+)(?:E[-+]?\d+)?)[)\s]*$/i;

// The value of a numeric constant argument, `fallback` where it is omitted, or NaN for anything
// else, such as a reference.
function constant (arg: string | undefined, fallback: number): number {
  if (arg == null || arg.trim() === '') {
    return fallback;
  }
  const match = NUMERIC_CONSTANT.exec(arg)?.[1];
  return match == null ? NaN : Number(match.replace(/\s/g, ''));
}

// The element count of a rows × columns array, or NaN where a count is past `cap` or below 1, which
// builds no array, so that two negative counts make no size.
function size (rows: number, columns: number, cap = Infinity): number {
  return rows < 1 || columns < 1 || rows > cap || columns > cap ? NaN : rows * columns;
}

// The element count of the array that a call to each function builds, where its sizes are numeric
// constants and its counts at least 1 and within any cap; NaN otherwise.
const ARRAY_SIZES: Array<[ string, (args: string[]) => number ]> = [
  [ 'SEQUENCE', args => size(constant(args[0], NaN), constant(args[1], 1), COUNT_CAP) ],
  [ 'RANDARRAY', args => size(constant(args[0], 1), constant(args[1], 1), COUNT_CAP) ],
  [ 'MAKEARRAY', args => size(constant(args[0], NaN), constant(args[1], 1)) ],
  [ 'MUNIT', args => size(constant(args[0], NaN), constant(args[0], NaN)) ],
];

// The arguments of a call to `name` that is the whole of `text`, or null.
function wholeCall (text: string, name: string): string[] | null {
  const call = text.trim();
  if (!call.toUpperCase().startsWith(`${name}(`)) {
    return null;
  }
  const args = callArguments(call, name)[0];
  return args != null && call.length === name.length + args.join(',').length + 2 ? args : null;
}

// Whether the formula builds an array larger than Excel has been measured to compute, with sizes
// given as numeric constants, in a context where Excel has been measured to leave the cell at 0:
// the call is the whole formula, or the whole argument of ROWS. Inside a larger formula, such as an
// IF branch not taken, a 0 may be the formula's own answer.
export function buildsAbandonedArray (formula: string): boolean {
  const rows = wholeCall(formula, 'ROWS');
  const call = (rows?.length === 1 ? rows[0] : null) ?? formula;
  return ARRAY_SIZES.some(([ name, sizeOf ]) => {
    const args = wholeCall(call, name);
    return args != null && sizeOf(args) > LARGEST_COMPUTED;
  });
}
