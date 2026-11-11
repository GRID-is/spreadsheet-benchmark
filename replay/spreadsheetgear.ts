// Replay a report's cases through SpreadsheetGear for .NET reading the fixture itself, an
// xlsx-input replay like IronCalc's Python package and POI: SpreadsheetGear's own reader and
// parser see the original file, so their gaps count.
//
// Each fixture is evaluated in a fresh .NET process (replay/spreadsheetgear/Program.cs), which
// opens the workbook, forces a full recalculation (`CalculateFull`) and writes the values of the
// case cells as JSON. The shim is built once into .cache/spreadsheetgear with `dotnet build`
// (brew install dotnet); buildMs includes the runtime's start. Like POI's shim, the shim then reads
// its heap: the retained heap and peak live heap, but no retained external memory, which .NET does
// not account for.
//
// Licence: without SPREADSHEETGEAR_LICENSE, SpreadsheetGear runs under its no-cost licence, which
// refuses workbooks over 1,000 rows x 100 columns x 10 worksheets; those fixtures fail to load. A
// published baseline needs the signed 30-day trial or a paid licence in SPREADSHEETGEAR_LICENSE.
// The EULA (spreadsheetgear.com/company/files/eula.html, read 2026-10-10) has no benchmark clause.
import { spawnSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseHeapLine, replay, type CaseCell, type Engine, type EngineValue } from './common.ts';
import { runLimited } from './child-process.ts';
import { type HeapProbe } from '../src/job-worker.ts';

const SOURCE = join(dirname(fileURLToPath(import.meta.url)), 'spreadsheetgear');
const OUT = resolve('.cache', 'spreadsheetgear');
const DLL = join(OUT, 'SgEvaluate.dll');
const DOTNET = process.env.DOTNET ?? (existsSync('/opt/homebrew/bin/dotnet') ? '/opt/homebrew/bin/dotnet' : 'dotnet');

export function ensureBuilt () {
  const sources = [ 'Program.cs', 'SgEvaluate.csproj' ].map(f => statSync(join(SOURCE, f)).mtimeMs);
  if (existsSync(DLL) && statSync(DLL).mtimeMs >= Math.max(...sources)) return;
  const res = spawnSync(DOTNET, [ 'build', SOURCE, '-c', 'Release', '-o', OUT, '--nologo', '-v', 'quiet' ], { encoding: 'utf8' });
  if (res.error != null) throw new Error(`${DOTNET} does not run (brew install dotnet): ${res.error.message}`);
  if (res.status !== 0) throw new Error(`dotnet build failed: ${(res.stdout + res.stderr).trim().slice(-600)}`);
}

/** Throws unless dotnet runs, which the runner checks before it replays anything. */
export function checkPrerequisites (): void {
  const res = spawnSync(DOTNET, [ '--list-runtimes' ], { encoding: 'utf8' });
  if (res.error != null || res.status !== 0 || !res.stdout.includes('Microsoft.NETCore.App ')) throw new Error(`${DOTNET} does not run (brew install dotnet)`);
}

// The SpreadsheetGear package that the shim is built against, read from the project file so that
// reading the version builds nothing.
const version = /<PackageReference Include="SpreadsheetGear" Version="([^"]+)"/.exec(readFileSync(join(SOURCE, 'SgEvaluate.csproj'), 'utf8'))?.[1] ?? 'unknown';

export const engine = {
  name: 'spreadsheetgear',
  version,
  warmUp: false,
  inProcess: false,
  options: { input: 'xlsx', reader: 'spreadsheetgear', licensed: Boolean(process.env.SPREADSHEETGEAR_LICENSE) },
  async buildFile (fixture: string, cases: CaseCell[], heap?: HeapProbe) {
    ensureBuilt();
    const dir = mkdtempSync(join(tmpdir(), 'spreadsheetgear-'));
    try {
      const targets = join(dir, 'targets.json');
      const valuesPath = join(dir, 'values.json');
      writeFileSync(targets, JSON.stringify(cases.map(c => ({ sheet: c.sheet, ref: c.ref }))));
      const budget = Math.max(5000, Number(process.env.REPLAY_TIMEOUT_MS ?? 60000) - 5000);
      const res = await runLimited(DOTNET, [ DLL, fixture, targets, valuesPath ], { timeoutMs: budget, maxBuffer: 64 * 1024 * 1024 });
      if (res.error) throw res.error;
      if (res.status !== 0) {
        // .NET prints "Unhandled exception. <type>: <message>", with a multi-line message on the
        // lines that follow (the no-cost licence's refusal), then the stack.
        const lines = (res.stderr || res.stdout || '').split('\n').map(l => l.trim()).filter(l => l !== '' && !l.startsWith('at '));
        throw new Error(`spreadsheetgear: ${lines.slice(0, 2).join(' ').replace(/^Unhandled exception\. /, '').slice(0, 300)}`);
      }
      const values: Record<string, EngineValue> = JSON.parse(readFileSync(valuesPath, 'utf8'));
      for (const line of res.stdout.split('\n')) {
        const figures = parseHeapLine(line);
        if (figures) heap?.report(figures);
      }
      return (c: CaseCell) => values[`${c.sheet}!${c.ref}`] ?? null;
    }
    finally {
      rmSync(dir, { recursive: true, force: true });
    }
  },
} satisfies Engine;

// Does Open + CalculateFull recalculate, or keep Excel's cached values? Poison one cached value
// (the same cell as replay/ironcalc-xlsx.ts) and see which answer comes back.
async function verifyRecalc (): Promise<void> {
  const { execFileSync } = await import('node:child_process');
  const source = resolve('suites/ironcalc/excel/arithmetic.xlsx');
  const dir = mkdtempSync(join(tmpdir(), 'spreadsheetgear-verify-'));
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
    console.log(`SpreadsheetGear ${engine.version} returned ${got} -> ${ok
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
