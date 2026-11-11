// Types for the engine packages that the runners import but the repository does not depend on:
// they are installed ad hoc (`npm install --no-save`, see replay/README.md), so the type check
// cannot rely on their own types. Each declaration covers only what its runner uses. A
// declaration here takes precedence over an installed package's own types, so the type check
// gives the same result whether or not the package is installed.

declare module 'hyperformula' {
  type RawCellContent = string | number | boolean | null | undefined;
  // A formula error is a DetailedCellError, whose value is the error string.
  type CellValue = string | number | boolean | null | { value: string };
  export class HyperFormula {
    static readonly version: string;
    static buildFromSheets (sheets: Record<string, RawCellContent[][]>, config?: { licenseKey?: string }): HyperFormula;
    getSheetId (sheetName: string): number | undefined;
    getCellValue (address: { sheet: number, row: number, col: number }): CellValue;
    destroy (): void;
  }
}

declare module 'fast-formula-parser' {
  type Position = { sheet: string, row: number, col: number };
  type CellRange = { sheet: string, from: { row: number, col: number }, to: { row: number, col: number } };
  class FormulaError extends Error {
    constructor (error: string, msg?: string);
    // The error string, such as '#DIV/0!'.
    readonly error: string;
  }
  export default class FormulaParser {
    static readonly FormulaError: typeof FormulaError;
    constructor (config: { onCell (ref: Position): unknown, onRange (ref: CellRange): unknown[][] });
    parse (formula: string, position: Position, allowReturnArray?: boolean): unknown;
  }
}

declare module 'playwright' {
  type Route = { fulfill (response: { status: number, contentType: string, body: string }): Promise<void> };
  type Page = {
    route (url: string, handler: (route: Route) => unknown): Promise<void>,
    goto (url: string): Promise<unknown>,
    addScriptTag (options: { path: string }): Promise<unknown>,
    evaluate<R, A> (pageFunction: (arg: A) => R | Promise<R>, arg: A): Promise<R>,
  };
  type Browser = { newPage (): Promise<Page>, close (): Promise<void> };
  export const chromium: { launch (options: { channel?: string, headless?: boolean }): Promise<Browser> };
}
