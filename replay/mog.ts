// Mog (sheetmog.ai), a local native CLI. Apache-2.0, with no benchmark-disclosure restriction
// (LICENSE at fundamental-research-labs/mog v1.0.0, checked 2026-10-30).
// Convert-and-read: --recalculate is required; an open/save without it preserves imported caches.
// buildMs includes process startup, xlsx reading, whole-workbook recalculation and xlsx writing.
import { spawnSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { ZipArchive } from '@borgar/zip';
import { cachedValues, convertAndRead, guardConvertedFileExists, replay, type Engine } from './common.ts';
import { runLimited } from './child-process.ts';
import { verifyRecalc } from './verify-recalc.ts';

const MOG_CLI = process.env.MOG_CLI ?? 'mog';

function version (): string {
  const res = spawnSync(MOG_CLI, [ '--version' ], { encoding: 'utf8' });
  const match = /^mog (\d+\.\d+\.\d+(?:-[\w.-]+)?)/m.exec(res.stdout ?? '');
  if (res.error || res.status !== 0 || !match) throw new Error('cannot read mog --version');
  return match[1];
}

async function recalc (fixture: string, dir: string): Promise<string> {
  const out = join(dir, 'out.xlsx');
  const budget = Math.max(5000, Number(process.env.REPLAY_TIMEOUT_MS ?? 60000) - 5000);
  const res = await runLimited(MOG_CLI, [ '--input', fixture, '--output', out, '--recalculate' ], { timeoutMs: budget });
  if (res.error) throw res.error;
  if (res.status !== 0) {
    throw new Error(`mog --recalculate failed (${res.signal ? `killed by ${res.signal}` : `exit status ${res.status}`}): ${(res.stderr || res.stdout).trim().slice(0, 300)}`);
  }
  return guardConvertedFileExists(out, 'mog --recalculate', res);
}

let probedVersion: string | undefined;

export const engine = {
  name: 'mog',
  get version () {
    return probedVersion ??= version();
  },
  options: { input: 'xlsx', recalculate: true },
  warmUp: false,
  inProcess: false,
  buildFile: (fixture: string) => convertAndRead(recalc, fixture),
} satisfies Engine;

export function checkPrerequisites (): void {
  const res = spawnSync(MOG_CLI, [ '--version' ], { encoding: 'utf8' });
  if (res.error || res.status !== 0) {
    throw new Error('mog not found or failed to start (download a release from fundamental-research-labs/mog, or set MOG_CLI)');
  }
}

async function verify (): Promise<void> {
  await verifyRecalc(recalc, 'mog --recalculate');
  const shared = join(import.meta.dirname, 'recalc-canary-shared.xlsx');
  const readShared = await convertAndRead(recalc, shared);
  const sharedValue = readShared({ sheet: 'Sheet1', ref: 'B2' });
  const sharedOk = sharedValue === 20;
  console.log(`Shared-formula follower B2 (cache 99999, truth 20): ${sharedValue} -> ${sharedOk ? 'RECALCULATED' : 'PASSED THE CACHE THROUGH'}`);

  const fixture = join(import.meta.dirname, '../suites/ironcalc/excel/LOOKUP_AND_REFERENCE/TRANSPOSE.xlsx');
  const cells = [ 'E5', 'F5', 'E6', 'F6', 'E7', 'F7' ].map(ref => ({ sheet: 'Sheet1', ref }));
  const excel = await cachedValues(fixture);
  const readSpill = await convertAndRead(async (source, dir) => {
    const bytes = readFileSync(source);
    const zip = new ZipArchive(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
    let xml = await zip.readText('xl/worksheets/sheet1.xml');
    if (xml === undefined) throw new Error('Mog spill canary has no worksheet');
    for (const { ref } of cells) {
      let hits = 0;
      xml = xml.replace(new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*>[\\s\\S]*?</c>`), cell =>
        cell.replace(/<v>[^<]*<\/v>/, () => { hits++; return '<v>99999</v>'; }));
      if (hits !== 1) throw new Error(`Mog spill canary did not poison ${ref}`);
    }
    await zip.write('xl/worksheets/sheet1.xml', xml);
    const poisoned = join(dir, 'poisoned-spill.xlsx');
    writeFileSync(poisoned, Buffer.from(zip.toArrayBuffer()));
    return recalc(poisoned, dir);
  }, fixture);
  const spillValues = cells.map(readSpill);
  const spillOk = cells.every((cell, i) => spillValues[i] === excel(cell));
  console.log(`TRANSPOSE spill E5:F7 (all caches 99999): ${JSON.stringify(spillValues)} -> ${spillOk ? 'RECALCULATED' : 'PASSED THE CACHE THROUGH'}`);
  if (!sharedOk || !spillOk) process.exitCode = 1;
}

if (import.meta.main) {
  checkPrerequisites();
  if (process.argv.includes('--verify-recalc')) await verify();
  else await replay(engine, import.meta.url);
}
