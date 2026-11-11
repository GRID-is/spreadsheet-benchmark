import { type Case, type Spill } from './case.ts';
import { lookup, type Catalogue } from './catalogue.ts';
import { readsEnvironment } from './environment.ts';
import { userFunctionNames } from './formula-text.ts';
import { callsUnstableWeeknum } from './unstable-weeknum.ts';
import { buildsAbandonedArray } from './abandoned-array.ts';

export type Status = 'pass' | 'fail' | 'unsupported' | 'error' | 'skip';
export type SkipReason = 'excel-dropped-formula' | 'excel-read-as-user-function' | 'excel-lacks-function' | 'volatile' | 'volatile-precedent' | 'environment' | 'excel-unstable' | 'excel-abandoned-array';

// What kind of disagreement a failed case is. Only `wrong-value` is a wrong
// answer in the everyday sense; `wrong-error` means both sides refused the
// input but named the refusal differently (#NUM! against #VALUE!), which some
// comparisons want to tolerate. The last two are of a case whose own cell
// matches: `wrong-spill-size` where the engine reports a spill range other
// than Excel's, `wrong-spilled-value` where a spilled cell differs.
export type FailKind = 'wrong-value' | 'wrong-error' | 'error-for-value' | 'value-for-error' | 'wrong-spill-size' | 'wrong-spilled-value';

// An error value, by the spelling engines and Excel use. The terminator is what
// separates a real error from text that merely starts with a #: LEFT(A8,3) over a
// #DIV/0! cell yields the string "#DI", and A1&B1 over #N/A yields "#N/A0". The
// three without one are named, #N/A and #GETTING_DATA from Excel and #FMT from
// LibreOffice, whose TEXT() answers it when it cannot format the value.
// The #"NAME" form is Gnumeric's syntax for an error it has no constant for: an
// Excel-only code comes back as #"#CALC!", and a formula Gnumeric cannot parse is
// replaced by ERROR("<the formula text>"), which answers #"<the formula text>".
// Both are errors rather than text: in the xlsx Gnumeric writes, those cells are
// t="e".
const ERROR_VALUE = /^#(?:"[\s\S]*"|N\/A|GETTING_DATA|FMT|[A-Z0-9_/]+[!?])$/;

export function isErrorValue (value: unknown): boolean {
  return typeof value === 'string' && ERROR_VALUE.test(value);
}

export function failKindOf (excelValue: unknown, engineValue: unknown): FailKind {
  const excelError = isErrorValue(excelValue);
  const engineError = isErrorValue(engineValue);
  return excelError && engineError ? 'wrong-error' : engineError ? 'error-for-value' : excelError ? 'value-for-error' : 'wrong-value';
}

// How a spill disagrees with Excel's, or undefined where it agrees: the engine's spill range, where
// it reports one, is Excel's, and every spilled cell of Excel's spill range matches.
function spillFailKind (spill: Spill): FailKind | undefined {
  if (spill.apiaryRange !== undefined && spill.apiaryRange !== spill.excelRange) {
    return 'wrong-spill-size';
  }
  return spill.mismatches === 0 ? undefined : 'wrong-spilled-value';
}

export type Classification = {
  status: Status,
  skipReason?: SkipReason,
  failKind?: FailKind,
  parity: boolean,
  propagatedError: boolean,
  neverSupported: boolean,
};

export function classify (c: Case, catalogue: Catalogue): Classification {
  const entries = c.functions.map(name => lookup(catalogue, name));
  const outer = entries[0];
  const neverSupported = entries.some(e => e.apiary === 'never');
  const propagatedError = c.apiaryErrorOrigin != null && c.apiaryErrorOrigin !== '' && c.apiaryErrorOrigin !== `${c.sheet}!${c.ref}`;
  const skip = (skipReason: SkipReason): Classification => ({ status: 'skip', skipReason, parity: false, propagatedError, neverSupported });
  if (c.loadError != null) {
    return { status: 'error', parity: true, propagatedError, neverSupported };
  }
  if (c.excelDroppedFormula) {
    return skip('excel-dropped-formula');
  }
  // The file named a function Excel knows without the prefix Excel needs, so
  // Excel took it for an unknown user function and answered #NAME?. That
  // answer reflects how the file was written, not what the function computes.
  // Apiary 18.3.0 and later keep the prefix in `formula`, and then record no `fileFormula`.
  if (userFunctionNames(c.fileFormula ?? c.formula).some(name => lookup(catalogue, name).inExcel)) {
    return skip('excel-read-as-user-function');
  }
  if (outer && !outer.inExcel) {
    return skip('excel-lacks-function');
  }
  if (entries.some(e => e.nondeterministic)) {
    return skip('volatile');
  }
  // Excel's value comes from the value a volatile cell had at the Excel pass, which an engine
  // evaluating the fixture later need not reproduce.
  if (c.volatilePrecedent != null) {
    return skip('volatile-precedent');
  }
  // Excel's answer describes the machine and the Excel session of the Excel pass.
  if (readsEnvironment(c.formula)) {
    return skip('environment');
  }
  // Excel's answer varies from one recalculation to the next, so the fixture holds one of them.
  if (callsUnstableWeeknum(c.formula)) {
    return skip('excel-unstable');
  }
  // Excel abandons an array function's calculation past about 1E8 elements and shows 0 without an
  // error, which is no answer to compare with.
  if (c.excelValue === 0 && buildsAbandonedArray(c.formula)) {
    return skip('excel-abandoned-array');
  }
  if (!c.match) {
    return c.apiaryValue === '#NAME?'
      ? { status: 'unsupported', parity: true, propagatedError, neverSupported }
      : { status: 'fail', failKind: failKindOf(c.excelValue, c.apiaryValue), parity: true, propagatedError, neverSupported };
  }
  const spillFail = c.spill === undefined ? undefined : spillFailKind(c.spill);
  if (spillFail !== undefined) {
    return { status: 'fail', failKind: spillFail, parity: true, propagatedError, neverSupported };
  }
  return { status: 'pass', parity: true, propagatedError, neverSupported };
}
