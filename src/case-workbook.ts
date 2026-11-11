import { writeFileSync } from 'node:fs';
import { basename } from 'node:path';
import { ZipArchive } from '@borgar/zip';
import { Model, formulaParserReady } from '@grid-is/apiary';
import { parse } from '@grid-is/waspiary';
import type { Adapter } from './adapter.ts';
import type { CellValue } from './model/case.ts';

export type Size = { rows: number, cols: number };
// `spill` writes the formula as a dynamic array formula and keeps a range of that size free of other
// cells. Without it, the formula is an ordinary one, which Excel evaluates with implicit intersection.
export type CaseSpec = { ref: string, formula?: string, expected: CellValue, group: string, spill?: Size };
// `spills`: the same for the formulas in `rows`, by reference. `sheets`: more sheets that the
// block's formulas read, which have no cases and are written after its sheet. `names`: defined
// names scoped to each sheet of the block, so that blocks can define the same name differently.
export type Block = { sheet: string, rows: Array<Array<CellValue | undefined>>, cases: CaseSpec[], spills?: Record<string, Size>, sheets?: DataSheet[], names?: BlockName[] };
export type DataSheet = Pick<Block, 'sheet' | 'rows' | 'spills'>;

export const LITERAL_OFFSET = 26;
export const GROUP_OFFSET = 52;

const EXCEL_ERRORS = [ '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!', '#FIELD!', '#CONNECT!', '#BLOCKED!', '#UNKNOWN!' ];

type JsfCell = { v?: CellValue, t?: 'e', f?: string, F?: string };

function literal (value: CellValue): JsfCell {
  if (typeof value === 'string' && value.startsWith('#')) {
    return EXCEL_ERRORS.includes(value) ? { v: value, t: 'e' } : { f: `"${value.replace(/"/g, '""')}"` };
  }
  return { v: value };
}

function cellFor (value: CellValue): JsfCell {
  return typeof value === 'string' && value.startsWith('=') ? { f: value.slice(1) } : literal(value);
}

export function columnLetters (index: number): string {
  let letters = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  }
  return letters;
}

export function parseRef (ref: string): { col: number, row: number } {
  const match = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!match) {
    throw new Error(`not a cell reference: ${ref}`);
  }
  let col = 0;
  for (const ch of match[1]) {
    col = col * 26 + ch.charCodeAt(0) - 64;
  }
  return { col: col - 1, row: Number(match[2]) };
}

export function shifted (ref: string, columns: number): string {
  const { col, row } = parseRef(ref);
  return `${columnLetters(col + columns)}${row}`;
}

function spillRange (ref: string, { rows, cols }: Size): string {
  const { col, row } = parseRef(ref);
  const end = `${columnLetters(col + cols - 1)}${row + rows - 1}`;
  return end === ref ? ref : `${ref}:${end}`;
}

function cellsOf ({ sheet, rows, cases, spills = {} }: Block): Record<string, JsfCell> {
  const cells: Record<string, JsfCell> = {};
  const spillAreas = [ ...Object.entries(spills), ...cases.flatMap(({ ref, spill }) => spill ? [ [ ref, spill ] as const ] : []) ]
    .map(([ anchor, size ]) => ({ anchor, ...parseRef(anchor), ...size }));
  // The spill whose range holds ref, other than one anchored at ref
  const findSpiller = (ref: string): string | undefined => {
    const { col, row } = parseRef(ref);
    return spillAreas.find(a => a.anchor !== ref && col >= a.col && col < a.col + a.cols && row >= a.row && row < a.row + a.rows)?.anchor;
  };
  const put = (ref: string, cell: JsfCell, what: string): void => {
    if (ref in cells) {
      throw new Error(`${sheet}!${ref}: ${what} overwrites an existing cell`);
    }
    const spiller = findSpiller(ref);
    if (spiller !== undefined) {
      throw new Error(`${sheet}!${ref}: ${what} blocks the spill of ${spiller}`);
    }
    cells[ref] = cell;
  };
  spillAreas.forEach((a, i) => spillAreas.slice(i + 1).forEach(b => {
    if (a.col < b.col + b.cols && b.col < a.col + a.cols && a.row < b.row + b.rows && b.row < a.row + a.rows) {
      throw new Error(`${sheet}: spill of ${b.anchor} overlaps the spill of ${a.anchor}`);
    }
  }));
  rows.forEach((row, r) => row.forEach((value, c) => {
    if (c >= LITERAL_OFFSET) {
      throw new Error(`${sheet}: rows wider than ${LITERAL_OFFSET} columns collides with the literal columns`);
    }
    if (value !== undefined) {
      put(`${columnLetters(c)}${r + 1}`, cellFor(value), 'value');
    }
  }));
  for (const { ref, formula, expected, group } of cases) {
    if (formula !== undefined) {
      put(ref, { f: formula }, 'case');
    }
    else if (cells[ref]?.f === undefined) {
      throw new Error(`${sheet}!${ref}: case has no formula`);
    }
    put(shifted(ref, LITERAL_OFFSET), literal(expected), 'literal');
    put(shifted(ref, GROUP_OFFSET), { v: group }, 'group');
  }
  for (const area of spillAreas) {
    if (cells[area.anchor]?.f === undefined) {
      throw new Error(`${sheet}!${area.anchor}: spill is not on a formula`);
    }
    cells[area.anchor].F = spillRange(area.anchor, area);
  }
  return cells;
}

export type BlockName = { name: string, value: string };
type ScopedName = BlockName & { scope: string };

function sheetsOf (block: Block): string[] {
  return [ block.sheet, ...(block.sheets ?? []).map(s => s.sheet) ];
}

// The defined names of the workbook: those of each block, scoped to each of its sheets
function definedNames (blocks: Block[]): ScopedName[] {
  return blocks.flatMap(b => sheetsOf(b).flatMap(scope => (b.names ?? []).map(n => ({ ...n, scope }))));
}

// The skip reasons of writeCaseWorkbook
export const UNWRITABLE_CASE = 'formula the xlsx writer cannot write';
export const UNWRITABLE_CONTEXT = 'context formula the xlsx writer cannot write';

// What writeCaseWorkbook wrote: the blocks, less the cases it skipped, and how many it skipped for
// each reason
export type Written = { blocks: Block[], skipped: Record<string, number> };

function xmlText (text: string): string {
  return text.replace(/&(lt|gt|quot|apos|amp|#\d+|#x[\da-f]+);/gi, (_, entity: string) => {
    const named: Record<string, string> = { lt: '<', gt: '>', quot: '"', apos: "'", amp: '&' };
    if (entity[0] !== '#') {
      return named[entity.toLowerCase()];
    }
    return String.fromCodePoint(entity[1].toLowerCase() === 'x' ? parseInt(entity.slice(2), 16) : Number(entity.slice(1)));
  });
}

// Apiary's xlsx writer (@grid-is/jsf2xlsx 2.4.2) writes a formula that it cannot normalize, such as
// one that does not parse, reads another workbook ([Other.xlsx]Sheet1!A1) or contains an invalid
// array literal ({1,2;nope}), as IFERROR(__xludf.DUMMYFUNCTION("<formula>"), <cached value>),
// without its spill, so Excel answers the cached value. The <formula> there is Apiary's rewrite of
// the one given (a1 becomes A1), so a wrapper is recognized by its start alone.
const WRAPPER = 'IFERROR(__xludf.DUMMYFUNCTION(';

// The cells, by sheet, whose formula the writer wrapped. Throws on a defined name that the writer
// wrapped or dropped: it drops one that reads a sheet the workbook lacks, or reads cells without
// naming their sheet.
async function unwritable (xlsx: Uint8Array, cells: Map<string, Record<string, JsfCell>>, names: ScopedName[]): Promise<Map<string, Set<string>>> {
  const zip = new ZipArchive(xlsx.buffer.slice(xlsx.byteOffset, xlsx.byteOffset + xlsx.byteLength) as ArrayBuffer);
  const text = async (entry: string): Promise<string> => new TextDecoder().decode(await zip.read(entry) ?? new Uint8Array());
  const rels = await text('xl/_rels/workbook.xml.rels');
  const workbook = await text('xl/workbook.xml');
  const sheets = [ ...workbook.matchAll(/<sheet\s[^>]*\bname="([^"]*)"/g) ].map(([ , name ]) => xmlText(name));
  const writtenNames = [ ...workbook.matchAll(/<definedName\s([^>]*)>([^<]*)<\/definedName>/g) ].map(([ , attrs, formula ]) => {
    const sheetId = /\blocalSheetId="(\d+)"/.exec(attrs)?.[1];
    return { name: xmlText(/\bname="([^"]*)"/.exec(attrs)?.[1] ?? ''), scope: sheetId === undefined ? undefined : sheets[Number(sheetId)], formula: xmlText(formula) };
  });
  const label = ({ name, scope }: { name: string, scope?: string }): string => scope === undefined ? name : `${scope}!${name}`;
  // A case that reads a wrapped defined name answers the wrapper's value, and which cases read it
  // is not known here
  const wrapped = writtenNames.filter(n => n.formula.startsWith(WRAPPER)).map(label);
  if (wrapped.length) {
    throw new Error(`the xlsx writer cannot write the defined name ${wrapped.join(', ')}`);
  }
  const kept = new Set(writtenNames.map(n => label(n).toUpperCase()));
  const dropped = names.map(label).filter(n => !kept.has(n.toUpperCase()));
  if (dropped.length) {
    throw new Error(`the xlsx writer drops the defined name ${dropped.join(', ')}`);
  }
  const found = new Map<string, Set<string>>();
  for (const [ , attrs ] of workbook.matchAll(/<sheet\s([^>]*)>/g)) {
    const sheet = xmlText(/\bname="([^"]*)"/.exec(attrs)?.[1] ?? '');
    const id = /\br:id="([^"]*)"/.exec(attrs)?.[1];
    const target = [ ...rels.matchAll(/<Relationship\s([^>]*)>/g) ].map(([ , a ]) => a).find(a => a.includes(`Id="${id}"`));
    const path = /\bTarget="([^"]*)"/.exec(target ?? '')?.[1];
    const written = cells.get(sheet);
    if (path === undefined || written === undefined) {
      throw new Error(`${sheet}: sheet not found in the written workbook`);
    }
    let ref = '';
    for (const [ , cellAttrs, formula ] of (await text(`xl/${path.replace(/^\/?xl\//, '')}`)).matchAll(/<c\s([^>]*)>|<f\b[^>]*>([^<]*)<\/f>/g)) {
      if (cellAttrs !== undefined) {
        ref = /\br="([^"]*)"/.exec(cellAttrs)?.[1] ?? '';
        continue;
      }
      const given = written[ref]?.f;
      if (given !== undefined && !given.startsWith(WRAPPER) && xmlText(formula).startsWith(WRAPPER)) {
        found.set(sheet, (found.get(sheet) ?? new Set()).add(ref));
      }
    }
  }
  return found;
}

// A rectangle of cells: columns from 0, rows from 1, both ends included.
export type Area = { left: number, top: number, right: number, bottom: number };

function columnIndex (letters: string): number {
  return parseRef(`${letters}1`).col;
}

// The area that a reference node of Waspiary's syntax tree refers to; undefined for a node that is
// not a reference to fixed cells. The sheet of a prefixed reference is not told apart.
function areaOf (node: any): Area | undefined {
  if (typeof node?.cell === 'string') {
    const { col, row } = parseRef(node.cell.replaceAll('$', ''));
    return { left: col, top: row, right: col, bottom: row };
  }
  if (Array.isArray(node?.colrange)) {
    const [ a, b ] = node.colrange.map((c: string) => columnIndex(c.replaceAll('$', '')));
    return { left: Math.min(a, b), top: 1, right: Math.max(a, b), bottom: Infinity };
  }
  if (Array.isArray(node?.rowrange)) {
    const [ a, b ] = node.rowrange.map((r: string) => Number(r.replaceAll('$', '')));
    return { left: 0, top: Math.min(a, b), right: Infinity, bottom: Math.max(a, b) };
  }
  if (node?.prefixed !== undefined) {
    return areaOf(node.prefixed);
  }
  if (Array.isArray(node?.range)) {
    const [ a, b ] = node.range.map(areaOf);
    return a && b && { left: Math.min(a.left, b.left), top: Math.min(a.top, b.top), right: Math.max(a.right, b.right), bottom: Math.max(a.bottom, b.bottom) };
  }
  return undefined;
}

function areaAt (ref: string, { rows, cols }: Size = { rows: 1, cols: 1 }): Area {
  const { col, row } = parseRef(ref);
  return { left: col, top: row, right: col + cols - 1, bottom: row + rows - 1 };
}

function overlaps (a: Area, b: Area): boolean {
  return a.left <= b.right && b.left <= a.right && a.top <= b.bottom && b.top <= a.bottom;
}

const REFERENCE_NODES = [ 'cell', 'colrange', 'rowrange', 'prefixed', 'range' ];

/**
 * Whether the formula may read a cell of the areas: it has a reference that overlaps one, or one
 * that its text does not fix (one of the defined names, a table, INDIRECT or OFFSET), or it does
 * not parse. A name that is not defined reads no cell.
 */
export function mayRead (formula: string, areas: Area[], definedNames: string[] = []): boolean {
  const defined = new Set(definedNames.map(n => n.toUpperCase()));
  const visit = (node: any, bound: Set<string>): boolean => {
    if (node === null || typeof node !== 'object') {
      return false;
    }
    if (Array.isArray(node)) {
      return node.some(n => visit(n, bound));
    }
    if ('parse_error' in node || 'structured' in node) {
      return true;
    }
    if (REFERENCE_NODES.some(key => key in node)) {
      const area = areaOf(node);
      return area === undefined || areas.some(a => overlaps(a, area));
    }
    if (typeof node.name === 'string') {
      return defined.has(node.name.toUpperCase()) && !bound.has(node.name.toUpperCase());
    }
    if (typeof node.call === 'string' && /^(_xlfn\.)?(INDIRECT|OFFSET)$/i.test(node.call)) {
      return true;
    }
    if ('let' in node) {
      const inner = new Set(bound);
      for (const [ name, value ] of Object.entries(node.let)) {
        if (visit(value, inner)) {
          return true;
        }
        inner.add(name.toUpperCase());
      }
      return visit(node.expr, inner);
    }
    if ('lambda' in node) {
      return visit(node.lambda, new Set([ ...bound, ...node.params.map((p: string) => p.toUpperCase()) ]));
    }
    return Object.values(node).some(n => visit(n, bound));
  };
  return visit(parse(formula), new Set());
}

// The blocks less what the writer cannot write. A case on such a formula is skipped, and so is
// that formula, and so is a case that may read it, directly or through other formulas, since it
// would read an empty cell. A block with such a formula in its rows that no case sits on, or on one
// of its other sheets, is skipped whole: its cases would read the cached value of that formula, not
// its result. Only the block's own formulas are followed, as the importers write blocks that do not
// read each other's sheets.
function writable (blocks: Block[], found: Map<string, Set<string>>, definedNames: string[]): Written {
  const skipped: Record<string, number> = {};
  const skip = (reason: string, n: number): void => {
    if (n > 0) {
      skipped[reason] = (skipped[reason] ?? 0) + n;
    }
  };
  const kept = blocks.flatMap((block): Block[] => {
    const refs = found.get(block.sheet);
    if ((block.sheets ?? []).some(s => found.has(s.sheet))) {
      skip(UNWRITABLE_CONTEXT, block.cases.length);
      return [];
    }
    if (refs === undefined) {
      return [ block ];
    }
    const caseRefs = new Set(block.cases.map(c => c.ref));
    if ([ ...refs ].some(ref => !caseRefs.has(ref))) {
      skip(UNWRITABLE_CONTEXT, block.cases.length);
      return [];
    }
    const rows = block.rows.map(row => [ ...row ]);
    const spills = { ...block.spills };
    for (const ref of refs) {
      const { col, row } = parseRef(ref);
      if (rows[row - 1]?.[col] !== undefined) {
        rows[row - 1][col] = undefined;
      }
      delete spills[ref];
    }
    const removed = block.cases.filter(c => refs.has(c.ref));
    const formulas = new Map<string, string>();
    block.rows.forEach((row, r) => row.forEach((value, c) => {
      if (typeof value === 'string' && value.startsWith('=')) {
        formulas.set(`${columnLetters(c)}${r + 1}`, value.slice(1));
      }
    }));
    for (const { ref, formula } of block.cases) {
      if (formula !== undefined) {
        formulas.set(ref, formula);
      }
    }
    const areas = removed.map(({ ref, spill }) => areaAt(ref, spill));
    const lost = new Set(refs);
    for (let more = true; more;) {
      more = false;
      for (const [ ref, formula ] of formulas) {
        if (!lost.has(ref) && mayRead(formula, areas, definedNames)) {
          areas.push(areaAt(ref, block.cases.find(c => c.ref === ref)?.spill ?? block.spills?.[ref]));
          lost.add(ref);
          more = true;
        }
      }
    }
    const cases = block.cases.filter(c => !lost.has(c.ref));
    skip(UNWRITABLE_CASE, removed.length);
    skip(UNWRITABLE_CONTEXT, block.cases.length - removed.length - cases.length);
    return cases.length ? [ { ...block, rows, cases, ...(block.spills ? { spills } : {}) } ] : [];
  });
  return { blocks: kept, skipped };
}

async function toXlsx (path: string, blocks: Block[]): Promise<{ xlsx: Uint8Array, cells: Map<string, Record<string, JsfCell>> }> {
  const cells = new Map(blocks.flatMap(block => [
    [ block.sheet, cellsOf(block) ] as const,
    ...(block.sheets ?? []).map(s => [ s.sheet, cellsOf({ ...s, cases: [] }) ] as const),
  ]));
  if (new Set([ ...cells.keys() ].map(name => name.toUpperCase())).size !== blocks.reduce((n, b) => n + sheetsOf(b).length, 0)) {
    throw new Error('two sheets have the same name');
  }
  const names = definedNames(blocks);
  // Named by the file name alone, which CELL("filename") returns in brackets, so that the cached
  // values are the same wherever the workbook is written.
  const model = Model.fromJSF({ name: basename(path), sheets: [ ...cells ].map(([ name, sheetCells ]) => ({ name, cells: sheetCells })), ...(names.length ? { names } : {}) });
  return { xlsx: await model.getWorkbooks()[0].toXLSX('nodebuffer'), cells };
}

/**
 * Writes the blocks as a case workbook, less what the xlsx writer cannot write (see `writable`),
 * and returns what it wrote. Writes no file when no case is left, and throws on a defined name the
 * writer cannot write or drops.
 */
export async function writeCaseWorkbook (path: string, blocks: Block[]): Promise<Written> {
  await formulaParserReady;
  let { xlsx, cells } = await toXlsx(path, blocks);
  let written: Written = { blocks, skipped: {} };
  const found = await unwritable(xlsx, cells, definedNames(blocks));
  if (found.size) {
    written = writable(blocks, found, definedNames(blocks).map(n => n.name));
    if (!written.blocks.length) {
      return written;
    }
    ({ xlsx, cells } = await toXlsx(path, written.blocks));
    const left = [ ...await unwritable(xlsx, cells, definedNames(written.blocks)) ].flatMap(([ sheet, refs ]) => [ ...refs ].map(ref => `${sheet}!${ref}`));
    if (left.length) {
      throw new Error(`the xlsx writer still cannot write ${left.join(', ')}`);
    }
  }
  writeFileSync(path, xlsx);
  return written;
}

export const caseWorkbookAdapter: Adapter = {
  isCase: (_sheetIndex, ref, sheet) => parseRef(ref).col < GROUP_OFFSET && sheet.value(shifted(ref, GROUP_OFFSET)) !== undefined,
  upstreamRef: ref => shifted(ref, LITERAL_OFFSET),
};
