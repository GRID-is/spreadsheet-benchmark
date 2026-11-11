import type { Adapter, SheetContext } from '../../src/adapter.ts';

const TEST_COLUMNS = [ 'B', 'D', 'F' ];

function parseRef (ref: string): { col: string, row: number } {
  const match = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!match) {
    throw new Error(`not a cell reference: ${ref}`);
  }
  return { col: match[1], row: Number(match[2]) };
}

function headerRow (sheet: SheetContext): number {
  for (let row = 1; row <= sheet.lastRow; row++) {
    if (sheet.value(`A${row}`) === 'Function') {
      return row;
    }
  }
  return Infinity;
}

function functionsCase (sheetIndex: number, ref: string, sheet: SheetContext): boolean {
  const { col, row } = parseRef(ref);
  return sheetIndex === 0 && TEST_COLUMNS.includes(col) && row > headerRow(sheet) && sheet.value(`A${row}`) !== 'Total';
}

const adapter: Adapter = {
  isCase: (sheetIndex, ref, sheet) => sheet.file.startsWith('functions/') ? functionsCase(sheetIndex, ref, sheet) : true,
  upstreamRef: (ref, sheet) => {
    if (!sheet.file.startsWith('functions/')) {
      return undefined;
    }
    const { col, row } = parseRef(ref);
    return `${String.fromCharCode(col.charCodeAt(0) + 1)}${row}`;
  },
};

export default adapter;
