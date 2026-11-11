// Headless cell-input replay using Univer's core, docs, sheets and formula plugins.
// Loaded through CommonJS because engine-render's ESM opentype.js import fails in Node.
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, resolve } from 'node:path';
import { readCellInput, refToCoords, replay, type CaseCell, type Engine, type InputCell } from './common.ts';

type UniverCellData = { v?: string | number | boolean | null, t?: number, f?: string };
type FWorksheet = { getSheetName (): string, getRange (row: number, col: number): { getCellData (): UniverCellData | null } };
type UniverAPI = {
  getFormula (): { onCalculationResultApplied (timeout?: number): Promise<void> },
  createWorkbook (data: { id: string, sheetOrder: string[], sheets: Record<string, object> }): { getSheets (): FWorksheet[] },
};
type UniverPlugin = { new (...args: never[]): object, pluginName: string, packageName: string, version: string, type: number };
type UniverInstance = { registerPlugin (plugin: UniverPlugin): void, dispose (): void };

const require = createRequire(process.env.UNIVER_PACKAGE_ROOT
  ? resolve(process.env.UNIVER_PACKAGE_ROOT, 'package.json')
  : import.meta.url);
const { Univer, LocaleType, LogLevel }: {
  Univer: new (options: { locale: string, locales: Record<string, object>, logLevel: number }) => UniverInstance,
  LocaleType: { EN_US: string },
  LogLevel: { WARN: number },
} = require('@univerjs/core');
const { FUniver }: { FUniver: { newAPI (univer: UniverInstance): UniverAPI } } = require('@univerjs/core/facade');
const { UniverFormulaEnginePlugin }: { UniverFormulaEnginePlugin: UniverPlugin } = require('@univerjs/engine-formula');
const { UniverDocsPlugin }: { UniverDocsPlugin: UniverPlugin } = require('@univerjs/docs');
const { UniverSheetsPlugin }: { UniverSheetsPlugin: UniverPlugin } = require('@univerjs/sheets');
const { UniverSheetsFormulaPlugin }: { UniverSheetsFormulaPlugin: UniverPlugin } = require('@univerjs/sheets-formula');
const sheetsFormulaEnUS: object = require('@univerjs/sheets-formula/locale/en-US');
require('@univerjs/sheets/facade');
require('@univerjs/engine-formula/facade');
require('@univerjs/sheets-formula/facade');

const packages: Record<string, string> = Object.fromEntries([
  'core', 'docs', 'engine-formula', 'engine-render', 'protocol', 'rpc', 'sheets', 'sheets-formula', 'themes',
].map(name => {
  const id = `@univerjs/${name}`;
  return [ id, JSON.parse(readFileSync(resolve(dirname(require.resolve(id)), '../../package.json'), 'utf8')).version ];
}));
const version = packages['@univerjs/core'];
if (Object.values(packages).some(value => value !== version)) {
  throw new Error(`Univer packages must have one version: ${JSON.stringify(packages)}`);
}

(globalThis as { Image?: unknown }).Image ??= class Image {};

const NUMBER_TYPE = 2;
const BOOLEAN_TYPE = 3;

function cellFor (cell: InputCell): UniverCellData {
  if (cell.f != null) {
    return { f: `=${cell.f}` };
  }
  if (typeof cell.v === 'boolean') {
    return { v: cell.v ? 1 : 0, t: BOOLEAN_TYPE };
  }
  return { v: cell.v };
}

let previous: UniverInstance | null = null;

export const engine = {
  name: 'univer',
  version,
  options: { packages: JSON.stringify(packages) },
  async buildFile (fixture: string) {
    if (previous) {
      previous.dispose();
      previous = null;
    }
    const input = await readCellInput(fixture);
    const sheets: Record<string, object> = {};
    const sheetOrder: string[] = [];
    let hasFormulas = false;
    input.forEach((sheet, i) => {
      const id = `s${i}`;
      sheetOrder.push(id);
      const cellData: Record<number, Record<number, UniverCellData>> = {};
      let maxRow = 0;
      let maxCol = 0;
      for (const [ref, cell] of Object.entries(sheet.cells)) {
        if (cell.v == null && cell.f == null) continue;
        const { row, col } = refToCoords(ref);
        (cellData[row] ??= {})[col] = cellFor(cell);
        if (cell.f != null) hasFormulas = true;
        if (row > maxRow) maxRow = row;
        if (col > maxCol) maxCol = col;
      }
      sheets[id] = { id, name: sheet.name, cellData, rowCount: Math.max(maxRow + 10, 100), columnCount: Math.max(maxCol + 5, 40) };
    });
    const univer = new Univer({
      locale: LocaleType.EN_US,
      locales: { [LocaleType.EN_US]: sheetsFormulaEnUS },
      logLevel: LogLevel.WARN,
    });
    previous = univer;
    try {
      univer.registerPlugin(UniverFormulaEnginePlugin);
      univer.registerPlugin(UniverDocsPlugin);
      univer.registerPlugin(UniverSheetsPlugin);
      univer.registerPlugin(UniverSheetsFormulaPlugin);
      const univerAPI = FUniver.newAPI(univer);
      const workbook = univerAPI.createWorkbook({ id: 'replay', sheetOrder, sheets });
      if (hasFormulas) await univerAPI.getFormula().onCalculationResultApplied(30000);
      const byName = new Map<string, FWorksheet>();
      for (const sheet of workbook.getSheets()) {
        byName.set(sheet.getSheetName(), sheet);
      }
      return (c: CaseCell) => {
        const sheet = byName.get(c.sheet);
        if (!sheet) return null;
        const { row, col } = refToCoords(c.ref);
        const cell = sheet.getRange(row, col).getCellData();
        if (cell == null) return null;
        if (cell.t === BOOLEAN_TYPE) return cell.v === 1 || cell.v === true;
        if (cell.t === NUMBER_TYPE && typeof cell.v === 'string' && cell.v.trim() !== '' && Number.isFinite(Number(cell.v))) return Number(cell.v);
        return cell.v ?? null;
      };
    }
    catch (error) {
      univer.dispose();
      previous = null;
      throw error;
    }
  },
} satisfies Engine;

if (import.meta.main) {
  await replay(engine, import.meta.url);
}
