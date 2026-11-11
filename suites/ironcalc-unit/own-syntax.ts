// IronCalc's tests of its own formula syntax, which no xlsx file can contain, are left out.

// The test files of formulas in a German locale or language: German function names, and ; or , as
// the separators of arguments and decimals. An xlsx file has English names and Excel's separators.
export const LOCALE_TESTS = new Set([ 'test_locale.rs', 'test_language.rs' ]);

// Whether the formula declares a LAMBDA parameter in brackets, as IronCalc does an optional one:
// LAMBDA(a, [b], ...). Excel has no such syntax. A bracket in the body, such as a structured
// reference, is not a parameter: a parameter is followed by a comma.
export function hasBracketedLambdaParameter (formula: string): boolean {
  return /\bLAMBDA\(\s*(?:[^,()[\]]+,\s*)*\[[^\]]*\]\s*,/i.test(formula);
}
