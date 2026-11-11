import { existsSync } from 'node:fs';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { type CellValue } from './model/case.ts';

export type SheetContext = {
  file: string,
  name: string,
  lastRow: number,
  value: (ref: string) => CellValue | undefined,
};

export type Adapter = {
  isCase: (sheetIndex: number, ref: string, sheet: SheetContext) => boolean,
  upstreamRef?: (ref: string, sheet: SheetContext) => string | undefined,
};

export const everyFormulaCell: Adapter = {
  isCase: () => true,
};

export async function loadAdapter (suite: string): Promise<Adapter> {
  const path = resolve('suites', suite, 'adapter.ts');
  if (!existsSync(path)) {
    return everyFormulaCell;
  }
  const mod: { default: Adapter } = await import(pathToFileURL(path).href);
  return mod.default;
}
