import { callArguments } from './formula-text.ts';

// Invalid return types that Excel answers with a week number or #NUM!, varying from one
// recalculation to the next. Only the measured ones: Excel's answer to each of these changed
// between recalculations, except 150, whose cached value in the fixture differs between reports.
// Other invalid return types stay in until measured unstable; 0, 0.5, 3 and 9 were measured
// answering #NUM! every time.
const UNSTABLE = new Set([ -2, -1, -0.5, 22, 99, 100, 150, 1000, 32767, 34903, 34903.5, 999999 ]);

// A number, signed or not, in any parentheses.
const NUMERIC_CONSTANT = /^[(\s]*([-+]?\s*(?:\d+\.?\d*|\.\d+)(?:E[-+]?\d+)?)[)\s]*$/i;
// A DATE call with three unsigned integer constants, in any parentheses.
const DATE_CONSTANT = /^[(\s]*DATE\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)[)\s]*$/i;

// The serial number of DATE(year, month, day) in the 1900 date system, for a date from March 1900
// on, where the serial is the count of days since 1899-12-30; null for an earlier date.
function dateSerial (year: number, month: number, day: number): number | null {
  const ms = Date.UTC(year < 1900 ? year + 1900 : year, month - 1, day);
  return ms < Date.UTC(1900, 2, 1) ? null : (ms - Date.UTC(1899, 11, 30)) / 86_400_000;
}

// The return type a WEEKNUM argument gives as a constant: a number, or a DATE call of constants.
function constantReturnType (arg: string): number | null {
  const constant = NUMERIC_CONSTANT.exec(arg)?.[1];
  if (constant != null) {
    return Number(constant.replace(/\s/g, ''));
  }
  const date = DATE_CONSTANT.exec(arg);
  return date == null ? null : dateSerial(Number(date[1]), Number(date[2]), Number(date[3]));
}

// Whether the formula calls WEEKNUM with a constant return type that Excel answers unstably. A
// return type given by reference, other expression, text or boolean does not count: only the
// formula text is at hand.
export function callsUnstableWeeknum (formula: string): boolean {
  return callArguments(formula, 'WEEKNUM').some(args => {
    const returnType = constantReturnType(args[1] ?? '');
    return returnType != null && UNSTABLE.has(returnType);
  });
}
