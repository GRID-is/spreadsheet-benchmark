export type CellValue = string | number | boolean | null;

// The spill (CONTEXT.md) of a case whose formula answered an array, in Excel or in the engine under
// test. Both value arrays cover Excel's spill range, by row, its anchor first.
export type Spill = {
  // Excel's spill range, `C1:D3`, as the fixture stores it; the anchor alone where only the engine's
  // answer spilled
  excelRange: string,
  excelValues: CellValue[][],
  // The engine's spill range, where the engine under test reports one
  apiaryRange?: string,
  // Absent where the engine failed to load the workbook
  apiaryValues?: CellValue[][],
  // How many spilled cells of Excel's spill range differ under `eps`; absent with `apiaryValues`
  mismatches?: number,
};

export type Case = {
  suite: string,
  category: string,
  file: string,
  sheet: string,
  ref: string,
  id: string,
  formula: string,
  // The formula text as the file stores it, present only when it differs in
  // meaning from `formula`, which is Apiary's reading (see formula-text.ts).
  fileFormula?: string,
  functions: string[],
  excelValue: CellValue,
  apiaryValue: CellValue,
  match: boolean,
  eps: number,
  excelDroppedFormula: boolean,
  apiaryErrorOrigin?: string,
  // A nearest volatile cell, `Sheet!A1`, that the formula reads, directly or through other cells,
  // or may read through a reference OFFSET computes; the case's own cell when its formula is
  // volatile only through a defined name; empty when it can reach one only through a reference
  // INDIRECT computes, a formula that does not parse or a table that cannot be found. Absent when
  // the formula depends on none, and when it calls a volatile function itself (see
  // volatile-precedents.ts).
  volatilePrecedent?: string,
  upstreamLiteral?: CellValue,
  // Present where Excel's spill range or the engine's spans more than the case's cell
  spill?: Spill,
  loadError?: string,
  oracleApp: string,
  apiaryVersion: string,
};
