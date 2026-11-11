// Replay a report's cases through SpreadJS (MESCIUS @mescius/spread-sheets).
// SpreadJS assumes a real browser, so each file runs in headless system
// Chrome via Playwright (channel 'chrome', no browser download). The page has
// a localhost origin, where SpreadJS runs without a licence key, as MESCIUS's
// licensing page allows. The fixture is imported by SpreadJS's own xlsx
// importer (spread-sheets-io), so no other reader stands in between. The importer keeps Excel's cached
// values, so every workbook is recalculated in full before it is read;
// `--verify-recalc` proves that with a poisoned cached value. The evaluation
// build prepends an "Evaluation Version" watermark sheet, hence lookups go by
// sheet name. Browser launch and script loading are excluded from buildMs via
// harnessTiming.
// Install ad hoc: npm install --no-save @mescius/spread-sheets
//   @mescius/spread-sheets-io @mescius/spread-sheets-shapes
//   @mescius/spread-sheets-charts playwright …
import { readFileSync } from 'node:fs';
import { basename, resolve } from 'node:path';
import { chromium } from 'playwright';
import { harnessTiming, refToCoords, replay, type CaseCell, type Engine, type EngineValue } from './common.ts';

// The part of the SpreadJS API that the page function below uses. GC is a global of the page, not
// of this module, which only refers to it inside the page function.
type SpreadJsSheet = {
  getRowCount (): number,
  setRowCount (count: number): void,
  getColumnCount (): number,
  setColumnCount (count: number): void,
  getValue (row: number, col: number): EngineValue | Date | object,
};
type SpreadJsWorkbook = {
  suspendPaint (): void,
  import (file: File, success: (result?: unknown) => void, error: (e: unknown) => void, options: { fileType: number }): void,
  suspendCalcService (ignoreDirty: boolean): void,
  resumeCalcService (recalcAll: boolean): void,
  getSheetCount (): number,
  getSheet (index: number): SpreadJsSheet,
  getSheetFromName (name: string): SpreadJsSheet | null,
  calculate (type: number): void,
};
declare const GC: {
  Spread: {
    Sheets: {
      Workbook: new (host: HTMLElement, options: { allowDynamicArray: boolean }) => SpreadJsWorkbook,
      FileType: { excel: number },
      CalculationType: { rebuild: number },
    },
  },
};

const version: string = JSON.parse(readFileSync(resolve('node_modules', '@mescius', 'spread-sheets', 'package.json'), 'utf8')).version;

const SCRIPTS = [
  'node_modules/@mescius/spread-sheets/dist/gc.spread.sheets.all.min.js',
  'node_modules/@mescius/spread-sheets-shapes/dist/gc.spread.sheets.shapes.min.js',
  'node_modules/@mescius/spread-sheets-charts/dist/gc.spread.sheets.charts.min.js',
  'node_modules/@mescius/spread-sheets-io/dist/gc.spread.sheets.io.min.js',
].map(p => resolve(p));

export const engine = {
  name: 'spread-sheets',
  version,
  warmUp: false,
  inProcess: false,
  async buildFile (fixture: string, cases: CaseCell[]) {
    const t0 = Date.now();
    const browser = await chromium.launch({ channel: 'chrome', headless: true });
    let values: Record<string, EngineValue>;
    try {
      const page = await browser.newPage();
      await page.route('http://localhost/**', route => route.fulfill({ status: 200, contentType: 'text/html', body: '<!doctype html><html><body></body></html>' }));
      await page.goto('http://localhost/');
      for (const script of SCRIPTS) {
        await page.addScriptTag({ path: script });
      }
      const targets = cases.map(c => {
        const { row, col } = refToCoords(c.ref);
        return { key: `${c.sheet}!${c.ref}`, sheet: c.sheet, row, col };
      });
      harnessTiming.ms += Date.now() - t0; // harness setup, not engine work
      values = await page.evaluate(async ({ b64, name, targets: wanted }) => {
        const host = document.createElement('div');
        document.body.appendChild(host);
        const wb = new GC.Spread.Sheets.Workbook(host, { allowDynamicArray: true });
        wb.suspendPaint();
        const bin = atob(b64);
        const u8 = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
        const file = new File([u8], name);
        await new Promise((resolveImport, rejectImport) => wb.import(file, resolveImport, e => rejectImport(new Error(`import: ${JSON.stringify(e)}`)), { fileType: GC.Spread.Sheets.FileType.excel }));
        // The importer sizes each sheet to its used range (a one-row test sheet
        // gets one row), so a recalculated dynamic array has nowhere to spill
        // and answers #SPILL!. Pad every sheet first. Not to Excel's dimensions:
        // with those the rebuild calculation below took minutes on the larger
        // workbooks and crashed the tab on FLOOR_CEILING. A whole-row or
        // whole-column reference spans the padded sheet, not Excel's.
        wb.suspendCalcService(false);
        for (let i = 0; i < wb.getSheetCount(); i++) {
          const sheet = wb.getSheet(i);
          sheet.setRowCount(sheet.getRowCount() + 1000);
          sheet.setColumnCount(sheet.getColumnCount() + 100);
        }
        wb.resumeCalcService(false);
        // Excel's Ctrl+Shift+Alt+F9: rebuild the dependency tree and
        // recalculate every formula, discarding the imported cached values.
        // CalculationType.all is cheaper but recomputes only the anchor of
        // a dynamic array and never spills it; rebuild spills.
        wb.calculate(GC.Spread.Sheets.CalculationType.rebuild);
        const EPOCH = Date.UTC(1899, 11, 30);
        const out: Record<string, EngineValue> = {};
        for (const t of wanted) {
          const sheet = wb.getSheetFromName(t.sheet);
          if (!sheet) {
            out[t.key] = null;
            continue;
          }
          let v = sheet.getValue(t.row, t.col);
          if (v instanceof Date) {
            v = (Date.UTC(v.getFullYear(), v.getMonth(), v.getDate(), v.getHours(), v.getMinutes(), v.getSeconds(), v.getMilliseconds()) - EPOCH) / 86400000;
          }
          else if (v !== null && typeof v === 'object') {
            v = String(v);
          }
          out[t.key] = v;
        }
        return out;
      }, { b64: readFileSync(fixture).toString('base64'), name: basename(fixture), targets });
    }
    finally {
      await browser.close();
    }
    return (c: CaseCell) => values[`${c.sheet}!${c.ref}`] ?? null;
  },
} satisfies Engine;

// Does the replay actually recalculate, or does SpreadJS keep Excel's cached
// values? Poison one cached value and see which answer comes back.
async function verifyRecalc (): Promise<void> {
  const { execFileSync } = await import('node:child_process');
  const { mkdtempSync, rmSync } = await import('node:fs');
  const { tmpdir } = await import('node:os');
  const { join } = await import('node:path');
  const source = resolve('suites/ironcalc/excel/arithmetic.xlsx');
  const dir = mkdtempSync(join(tmpdir(), 'spread-sheets-verify-'));
  try {
    const poisoned = join(dir, 'poisoned.xlsx');
    // The substitution must land, or the "poisoned" copy is the fixture unchanged
    // and E2 already holds the right answer, so the check passes without testing
    // anything. Regenerating the fixture in another shape is what would do it.
    execFileSync('python3', [ '-c', `
import zipfile, re, sys
SHEET = 'xl/worksheets/sheet1.xml'
zin = zipfile.ZipFile(sys.argv[1])
hits = 0
with zipfile.ZipFile(sys.argv[2], 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == SHEET:
            d, hits = re.subn(r'(<c r="E2"[^>]*>\\s*<f>C2\\+D2</f>\\s*<v>)[^<]+(</v>)',
                              r'\\g<1>99999\\g<2>', data.decode('utf8'), count=1)
            data = d.encode('utf8')
        zout.writestr(item, data)
if hits != 1:
    raise SystemExit(f'poisoning E2 in {SHEET} matched {hits} cells, expected 1: '
                     'the fixture is no longer the shape this check assumes')
`, source, poisoned ], { stdio: [ 'ignore', 'ignore', 'inherit' ] });
    const read = await engine.buildFile(poisoned, [ { sheet: 'Sheet1', ref: 'E2' } ]);
    const got = read({ sheet: 'Sheet1', ref: 'E2' });
    const ok = got === 3;
    console.log('E2 is =C2+D2; its cached value was poisoned to 99999, truth is 3.');
    console.log(`SpreadJS ${version} returned ${got} -> ${ok
      ? 'RECALCULATED (scores are real)'
      : 'PASSED THE CACHE THROUGH (scores would be meaningless)'}`);
    process.exitCode = ok ? 0 : 1;
  }
  finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

if (import.meta.main) {
  if (process.argv.includes('--verify-recalc')) await verifyRecalc();
  else await replay(engine, import.meta.url);
}
