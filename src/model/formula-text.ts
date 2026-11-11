// Comparing the formula text a file stores with Apiary's reading of it.
//
// Excel stores formulas in a file form that differs from the typed form by
// convention alone: storage prefixes (_xlfn., _xlws., _xlpm., _xlnm.),
// SINGLE(x) for @x, ANCHORARRAY(C1) for C1#, A1:A1 for A1, CRLF line breaks.
// Those differences say nothing about whether Apiary read the formula
// correctly, so the stored text is reduced to the typed form first and only
// what remains counts as a difference. Apiary's side is left as it is: a
// storage prefix surviving in Apiary's text is itself a misreading.

export function storedFormulaDiffers (stored: string, apiary: string): boolean {
  return typedForm(stored) !== apiary.replace(/\r\n/g, '\n');
}

export function typedForm (stored: string): string {
  const text = stored.replace(/\r\n/g, '\n');
  const unprefixed = mapOutsideSpans(text, s => s
    .replace(/(?<![\w.])(?:_xl(?:fn|ws|pm|nm)\.)+/g, '')
    // a sheet name that reads as a cell reference is quoted when typed
    .replace(/(?<![\w'.:!$])([A-Za-z]{1,3}\d+)!/g, "'$1'!"));
  const rewritten = rewriteCalls(rewriteCalls(unprefixed, 'SINGLE', arg => `@${operand(arg)}`), 'ANCHORARRAY', arg => `${operand(arg)}#`);
  return mapOutsideSpans(rewritten, s => s.replace(/(?<![\w$.:])(\$?[A-Z]{1,3}\$?\d+):\1(?![\w$:])/g, '$1'));
}

// Function names a file marks with _xludf., Excel's prefix for a user-defined
// function it does not recognise. A name Excel does know ends up here when
// the authoring tool omitted the prefix Excel needs (LibreOffice writing
// MINIFS rather than _xlfn.MINIFS), and Excel then answers #NAME?.
export function userFunctionNames (stored: string): string[] {
  const names: string[] = [];
  mapOutsideSpans(stored, s => {
    for (const m of s.matchAll(/(?<![\w.])_xludf\.([A-Za-z][\w.]*)\s*\(/g)) {
      names.push((m[1] as string).toUpperCase());
    }
    return s;
  });
  return names;
}

// The formula with each quoted sheet name and bracketed reference blanked to as many spaces, so
// that a pattern can match its function calls and string literals but no text that only names a
// sheet, a table column or another workbook.
export function blankNames (formula: string): string {
  let out = '';
  for (let i = 0; i < formula.length; i++) {
    const ch = formula[i] as string;
    if (SPAN_START.includes(ch)) {
      const end = spanEnd(formula, i);
      out += ch === '"' ? formula.slice(i, end) : ' '.repeat(end - i);
      i = end - 1;
    }
    else {
      out += ch;
    }
  }
  return out;
}

// The argument texts of each call of the function `name` in a formula, nested calls included,
// untrimmed. A call with a storage prefix (_xlfn.NAME) or within a longer name does not count.
export function callArguments (formula: string, name: string): string[][] {
  const calls: string[][] = [];
  const callee = `${name.toUpperCase()}(`;
  for (let i = 0; i < formula.length; i++) {
    if (SPAN_START.includes(formula[i] as string)) {
      i = spanEnd(formula, i) - 1;
    }
    else if (formula.slice(i, i + callee.length).toUpperCase() === callee && !/[\w.]/.test(i === 0 ? '' : formula[i - 1] as string)) {
      const args = splitArguments(formula, i + name.length);
      if (args) {
        calls.push(args);
      }
    }
  }
  return calls;
}

// The texts between the commas of the parenthesised argument list that opens at `open`, or null
// where it does not close.
function splitArguments (text: string, open: number): string[] | null {
  const args: string[] = [];
  let depth = 0;
  let start = open + 1;
  for (let i = open; i < text.length; i++) {
    const ch = text[i];
    if (SPAN_START.includes(ch as string)) {
      i = spanEnd(text, i) - 1;
    }
    else if (ch === '(' || ch === '{') {
      depth++;
    }
    else if (ch === ',' && depth === 1) {
      args.push(text.slice(start, i));
      start = i + 1;
    }
    else if ((ch === ')' || ch === '}') && --depth === 0) {
      args.push(text.slice(start, i));
      return args;
    }
  }
  return null;
}

// Apiary writes @ and # against a single operand without parentheses, and
// parenthesises nothing else; a compound argument keeps its grouping here, so
// a reading that loses it (@SEQUENCE(2,2) * 42 for SINGLE(SEQUENCE(2,2) * 42))
// still counts as a difference.
function operand (arg: string): string {
  const a = arg.trim();
  if (/^[-+]?(?:"(?:[^"]|"")*"|\{[^{}]*\}|(?:'(?:[^']|'')*'!)?[^\s()+\-*/^&=<>,]+)$/.test(a)) {
    return a;
  }
  const call = /^@*[A-Za-z_][\w.]*\(/.exec(a);
  if (call && matchingParen(a, call[0].length - 1) === a.length - 1) {
    return a;
  }
  return `(${a})`;
}

// Replace NAME(arg) with render(arg), matching the closing parenthesis while
// skipping string literals, quoted sheet names and bracketed references, where
// a formula may quote parentheses.
function rewriteCalls (text: string, name: string, render: (arg: string) => string): string {
  let out = '';
  let i = 0;
  while (i < text.length) {
    if (SPAN_START.includes(text[i] as string)) {
      const end = spanEnd(text, i);
      out += text.slice(i, end);
      i = end;
      continue;
    }
    const before = i === 0 ? '' : text[i - 1] as string;
    if (text.startsWith(`${name}(`, i) && !/[\w.]/.test(before)) {
      const open = i + name.length;
      const close = matchingParen(text, open);
      if (close >= 0) {
        out += render(rewriteCalls(text.slice(open + 1, close), name, render));
        i = close + 1;
        continue;
      }
    }
    out += text[i];
    i++;
  }
  return out;
}

function matchingParen (text: string, open: number): number {
  let depth = 0;
  for (let i = open; i < text.length; i++) {
    const ch = text[i];
    if (SPAN_START.includes(ch as string)) {
      i = spanEnd(text, i) - 1;
    }
    else if (ch === '(') {
      depth++;
    }
    else if (ch === ')' && --depth === 0) {
      return i;
    }
  }
  return -1;
}

const SPAN_START = `"'[`;

// Index just past a span holding no formula syntax: a string literal or quoted
// sheet name, where a doubled quote escapes one, or a bracketed reference
// (Table1[Item '#]), which nests and where ' escapes the next character.
function spanEnd (text: string, start: number): number {
  const quote = text[start];
  if (quote === '[') {
    let depth = 0;
    for (let i = start; i < text.length; i++) {
      if (text[i] === "'") {
        i++;
      }
      else if (text[i] === '[') {
        depth++;
      }
      else if (text[i] === ']' && --depth === 0) {
        return i + 1;
      }
    }
    return text.length;
  }
  let i = start + 1;
  while (i < text.length) {
    if (text[i] === quote) {
      if (text[i + 1] === quote) {
        i += 2;
        continue;
      }
      return i + 1;
    }
    i++;
  }
  return text.length;
}

function mapOutsideSpans (text: string, fn: (segment: string) => string): string {
  const starts = /["'[]/g;
  let out = '';
  let i = 0;
  while (i < text.length) {
    starts.lastIndex = i;
    const start = starts.exec(text)?.index ?? -1;
    if (start < 0) {
      out += fn(text.slice(i));
      break;
    }
    out += fn(text.slice(i, start));
    const end = spanEnd(text, start);
    out += text.slice(start, end);
    i = end;
  }
  return out;
}
