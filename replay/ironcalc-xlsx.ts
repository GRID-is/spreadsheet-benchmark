// Replay a report's cases through IronCalc reading the fixture itself, via the
// xlsx importer in IronCalc's Python package (`pip install ironcalc`). This is
// an xlsx-input replay, like SpreadJS, Google Sheets, LibreOffice, Gnumeric and
// POI: IronCalc's own reader and parser see the original file, so their gaps
// count.
//
// Each fixture is evaluated in a fresh Python process (replay/ironcalc_xlsx.py),
// which reads the values back through IronCalc's typed API: numbers arrive as
// floats, so no formatting round trip is needed. Set IRONCALC_PYTHON to the
// interpreter that has the package (default: python3).
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { replay, type CaseCell, type Engine, type EngineValue } from './common.ts';
import { runLimited } from './child-process.ts';

const PYTHON = process.env.IRONCALC_PYTHON ?? 'python3';
const SCRIPT = join(dirname(fileURLToPath(import.meta.url)), 'ironcalc_xlsx.py');

function version (): string {
  const res = spawnSync(PYTHON, [ '-c', 'import importlib.metadata as m; print(m.version("ironcalc"))' ], { encoding: 'utf8' });
  if (res.status !== 0) throw new Error(`ironcalc is not importable by ${PYTHON}: ${(res.stderr || '').trim().slice(0, 200)}`);
  return res.stdout.trim();
}

let probedVersion: string | undefined;

export const engine = {
  name: 'ironcalc',
  // Probed when first read, which only the replay's own process does: the job process of each
  // fixture imports this module, and starting Python there would slow every fixture.
  get version () {
    return probedVersion ??= version();
  },
  warmUp: false,
  inProcess: false,
  options: { input: 'xlsx', reader: 'ironcalc' },
  async buildFile (fixture: string, cases: CaseCell[]) {
    const dir = mkdtempSync(join(tmpdir(), 'ironcalc-xlsx-'));
    try {
      const targets = join(dir, 'targets.json');
      const valuesPath = join(dir, 'values.json');
      writeFileSync(targets, JSON.stringify(cases.map(c => ({ sheet: c.sheet, ref: c.ref }))));
      const budget = Math.max(5000, Number(process.env.REPLAY_TIMEOUT_MS ?? 60000) - 5000);
      // The child's budget ends before the worker's timeout would, so this is
      // where a timeout surfaces.
      const res = await runLimited(PYTHON, [ SCRIPT, fixture, targets, valuesPath ], { timeoutMs: budget, maxBuffer: 256 * 1024 * 1024 });
      if (res.error) throw res.error;
      if (res.status !== 0) throw new Error(`ironcalc: ${(res.stderr || '').trim().split('\n').pop()?.slice(0, 300)}`);
      const values: Record<string, EngineValue> = JSON.parse(readFileSync(valuesPath, 'utf8'));
      return (c: CaseCell) => values[`${c.sheet}!${c.ref}`] ?? null;
    }
    finally {
      rmSync(dir, { recursive: true, force: true });
    }
  },
} satisfies Engine;

// Does load_from_xlsx + evaluate() actually recalculate, or does IronCalc keep
// Excel's cached values? Poison one cached value and see which answer comes
// back. Verified 2026-09-28 with ironcalc 0.8.3: 99999 before evaluate(), 3 after.
async function verifyRecalc (): Promise<void> {
  const { execFileSync } = await import('node:child_process');
  const { resolve } = await import('node:path');
  const source = resolve('suites/ironcalc/excel/arithmetic.xlsx');
  const dir = mkdtempSync(join(tmpdir(), 'ironcalc-verify-'));
  try {
    const poisoned = join(dir, 'poisoned.xlsx');
    execFileSync('python3', [ '-c', `
import zipfile, re, sys
zin = zipfile.ZipFile(sys.argv[1])
with zipfile.ZipFile(sys.argv[2], 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'xl/worksheets/sheet1.xml':
            d = data.decode('utf8')
            d2 = re.sub(r'(<c r="E2"[^>]*>\\s*<f>C2\\+D2</f>\\s*<v>)[^<]+(</v>)', r'\\g<1>99999\\g<2>', d, count=1)
            assert d2 != d, 'poison pattern did not match'
            data = d2.encode('utf8')
        zout.writestr(item, data)
`, source, poisoned ]);
    const read = await engine.buildFile(poisoned, [ { sheet: 'Sheet1', ref: 'E2' } ]);
    const got = read({ sheet: 'Sheet1', ref: 'E2' });
    const ok = got === 3;
    console.log('E2 is =C2+D2; its cached value was poisoned to 99999, truth is 3.');
    console.log(`IronCalc ${engine.version} (xlsx importer) returned ${got} -> ${ok
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
