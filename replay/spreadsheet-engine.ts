// Replay a report's cases through @grid-is/spreadsheet-engine, GRID's public engine package, reading
// the fixture itself and recalculating every formula cell. It reproduces GRID's row on the
// published benchmark with the public package, so anyone can check it without access to GRID's
// private builds.
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { ALL_FORMULA_CELLS, Model } from '@grid-is/spreadsheet-engine';
import { replay, type CaseCell, type Engine } from './common.ts';

const { version } = JSON.parse(readFileSync(resolve('node_modules', '@grid-is', 'spreadsheet-engine', 'package.json'), 'utf8'));

export const engine = {
  name: 'spreadsheet-engine',
  version,
  options: { input: 'xlsx' },
  async buildFile (fixture: string) {
    const model = await Model.fromXLSXFile(fixture);
    model.recalculate(ALL_FORMULA_CELLS);
    const value = (c: CaseCell) => {
      const v = model.getCell(c.ref, c.sheet)?.v;
      if (typeof v === 'object' && v !== null) return v.name;
      if (typeof v === 'number' && !Number.isFinite(v)) return '#NUM!';
      return v ?? null;
    };
    return Object.assign(value, { spillRange: (c: CaseCell) => model.getCell(c.ref, c.sheet)?.F ?? c.ref });
  },
} satisfies Engine;

if (import.meta.main) {
  await replay(engine, import.meta.url);
}
