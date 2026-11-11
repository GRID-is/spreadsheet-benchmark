import { blankNames } from './formula-text.ts';

// CELL("filename"), the path of the workbook file, or INFO of the current directory, the count of
// open sheets, the operating system and its version, or the version of the application; each with
// its info type as a whole literal string argument, not the start of an expression.
const ENVIRONMENT_CALL = /(?<![\w.])(?:CELL\s*\(\s*"filename"(?=\s*[,)])|INFO\s*\(\s*"(?:directory|numfile|osversion|release|system)"(?=\s*\)))/i;

// Whether the formula asks where it is evaluated, which depends on where and in what the workbook
// is open, not on its contents. A sheet or column name that reads as such a call does not count.
export function readsEnvironment (formula: string): boolean {
  return ENVIRONMENT_CALL.test(blankNames(formula));
}
