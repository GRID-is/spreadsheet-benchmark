// The positions of the characters of a formula outside its quoted text: its strings and quoted
// sheet names, in which a doubled quote stands for one quote.
function* unquoted (formula: string): Generator<number> {
  let quote: string | undefined;
  for (let i = 0; i < formula.length; i++) {
    const ch = formula[i];
    if (quote) {
      if (ch === quote) {
        quote = undefined;
      }
    }
    else if (ch === '"' || ch === '\'') {
      quote = ch;
    }
    else {
      yield i;
    }
  }
}

function argumentCount (args: string): number {
  if (args.trim() === '') {
    return 0;
  }
  let count = 1;
  let depth = 0;
  for (const i of unquoted(args)) {
    const ch = args[i];
    if (ch === '(' || ch === '{') {
      depth++;
    }
    else if (ch === ')' || ch === '}') {
      depth--;
    }
    else if (depth === 0 && ch === ',') {
      count++;
    }
  }
  return count;
}

// The position of the first ARRAYFORMULA call at or after from, outside quoted text; -1 if none.
function arrayFormulaCall (formula: string, from: number): number {
  for (const i of unquoted(formula)) {
    if (i >= from && /^ARRAYFORMULA\(/i.test(formula.slice(i, i + 13)) && !/\w/.test(formula[i - 1] ?? '')) {
      return i;
    }
  }
  return -1;
}

// The position of the parenthesis that closes the one at open; -1 if none.
function closingParenthesis (formula: string, open: number): number {
  let depth = 0;
  for (const i of unquoted(formula)) {
    if (i < open) {
      continue;
    }
    if (formula[i] === '(') {
      depth++;
    }
    else if (formula[i] === ')' && --depth === 0) {
      return i;
    }
  }
  return -1;
}

// Excel has no ARRAYFORMULA. HyperFormula evaluates its argument as an array formula, as Excel does
// a dynamic array formula of that argument. Returns the formula without its ARRAYFORMULA calls, or
// undefined when it has none or one of them does not have exactly one argument.
export function withoutArrayFormula (formula: string): string | undefined {
  let result = formula;
  for (let start = arrayFormulaCall(result, 0); start !== -1; start = arrayFormulaCall(result, start)) {
    const open = start + 'ARRAYFORMULA'.length;
    const close = closingParenthesis(result, open);
    if (close === -1) {
      return undefined;
    }
    const args = result.slice(open + 1, close);
    if (argumentCount(args) !== 1) {
      return undefined;
    }
    const before = result.slice(0, start);
    const after = result.slice(close + 1);
    const bare = /(^=|[(,])\s*$/.test(before) && /^\s*([),]|$)/.test(after);
    result = before + (bare ? args.trim() : `(${args})`) + after;
  }
  return result === formula ? undefined : result;
}
