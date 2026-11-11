import type { Adapter, SheetContext } from '../../src/adapter.ts';

type Layout = {
  isCase: (sheetIndex: number, col: number, row: number, sheet: SheetContext) => boolean,
  upstreamRef: (col: number, row: number) => string,
};

const COL_B = 1;
const COL_C = 2;
const COL_D = 3;
const COL_H = 7;
const COL_J = 9;

function parseRef (ref: string): { col: number, row: number } {
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

function refOf (col: number, row: number): string {
  let letters = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  }
  return `${letters}${row}`;
}

function label (sheet: SheetContext, col: number, row: number): string | undefined {
  const v = sheet.value(refOf(col, row));
  return typeof v === 'string' ? v : undefined;
}

function blockStarts (sheet: SheetContext, first: number, labelCol: number, sentinel: string): number[] {
  const starts: number[] = [];
  for (let row = first; row <= sheet.lastRow; row += 4) {
    const name = label(sheet, labelCol, row);
    if (name === undefined || name === sentinel) {
      break;
    }
    starts.push(row);
  }
  return starts;
}

function endRow (sheet: SheetContext, first: number, sentinel: string): number {
  for (let row = first; row <= sheet.lastRow; row++) {
    if (label(sheet, 0, row)?.toLowerCase() === sentinel.toLowerCase()) {
      return row;
    }
  }
  return Infinity;
}

const evalLayout: Layout = {
  isCase: (sheetIndex, col, row, sheet) =>
    sheetIndex === 0 && col >= COL_D &&
    [ ...blockStarts(sheet, 23, COL_B, '<END-OF-FUNCTIONS>'), ...blockStarts(sheet, 96, COL_B, '<END-OF-FUNCTIONS>') ].includes(row),
  upstreamRef: (col, row) => refOf(col, row + 1),
};

const functionsLayout: Layout = {
  isCase: (sheetIndex, col, row, sheet) =>
    sheetIndex >= 1 && col === COL_B && row >= 5 && row < endRow(sheet, 5, '<end>') &&
    label(sheet, 0, row)?.toLowerCase() !== '<skip>',
  upstreamRef: (_col, row) => refOf(COL_C, row),
};

const matrixLayout: Layout = {
  isCase: (sheetIndex, col, row, sheet) =>
    sheetIndex === 0 && col >= COL_H && col <= COL_J &&
    blockStarts(sheet, 2, 0, '<END>').some(start => row >= start && row < start + 3),
  upstreamRef: (col, row) => refOf(col + 3, row),
};

const multisheetLayout: Layout = {
  isCase: (_sheetIndex, col, row, sheet) => {
    const end = endRow(sheet, 11, '<END>');
    return sheet.name === 'test' && col === COL_D && row >= 11 && row < end && label(sheet, 0, row) !== undefined;
  },
  upstreamRef: (_col, row) => refOf(COL_C, row),
};

const layouts: Record<string, Layout> = {
  eval: evalLayout,
  functions: functionsLayout,
  matrix: matrixLayout,
  multisheet: multisheetLayout,
};

function layoutOf (file: string): Layout {
  const layout = layouts[file.split('/')[0]];
  if (!layout) {
    throw new Error(`no POI layout for ${file}`);
  }
  return layout;
}

const adapter: Adapter = {
  isCase: (sheetIndex, ref, sheet) => {
    const { col, row } = parseRef(ref);
    return layoutOf(sheet.file).isCase(sheetIndex, col, row, sheet);
  },
  upstreamRef: (ref, sheet) => {
    const { col, row } = parseRef(ref);
    return layoutOf(sheet.file).upstreamRef(col, row);
  },
};

export default adapter;
