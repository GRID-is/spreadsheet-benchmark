// Replay a report's cases through GridCraft (github.com/storytold/gridcraft), via its gridcraft-cli
// command line. Pattern: convert-and-read, like gnumeric.ts - `gridcraft-cli convert` opens the
// fixture with GridCraft's own xlsx reader and writes an xlsx whose cached values are GridCraft's
// answers, read back with cachedValues.
//
// `node replay/gridcraft.ts --verify-recalc` checks that the conversion recalculates rather than
// passing Excel's cached values through; run it after any GridCraft upgrade.
//
// buildMs covers a gridcraft-cli process spawn, GridCraft's parse and recalculation and an xlsx
// write, so it is not comparable to an in-process engine's build time.
import { spawnSync } from 'node:child_process';
import { join } from 'node:path';
import { convertAndRead, guardConvertedFileExists, replay, type Engine } from './common.ts';
import { runLimited } from './child-process.ts';
import { verifyRecalc } from './verify-recalc.ts';

const GRIDCRAFT_CLI = process.env.GRIDCRAFT_CLI ?? 'gridcraft-cli';

function version (): string {
  const res = spawnSync(GRIDCRAFT_CLI, [ '--version' ], { encoding: 'utf8' });
  const m = /(\d+\.\d+\.\d+)/.exec(`${res.stdout}${res.stderr}`);
  return m ? m[1] : 'unknown';
}

async function recalc (fixture: string, dir: string): Promise<string> {
  const out = join(dir, 'out.xlsx');
  const budget = Math.max(5000, Number(process.env.REPLAY_TIMEOUT_MS ?? 60000) - 5000);
  const res = await runLimited(GRIDCRAFT_CLI, [ 'convert', fixture, out ], { timeoutMs: budget });
  if (res.error) throw res.error;
  return guardConvertedFileExists(out, 'gridcraft-cli', res);
}

/**
 * Throws unless gridcraft-cli starts, which the runner checks before it replays anything:
 * without it, the version reads as unknown.
 */
export function checkPrerequisites (): void {
  if (spawnSync(GRIDCRAFT_CLI, [ '--version' ]).error) throw new Error('gridcraft-cli not found (download it from the GridCraft releases on GitHub, or set GRIDCRAFT_CLI)');
}

let probedVersion: string | undefined;

export const engine = {
  name: 'gridcraft',
  get version () {
    return probedVersion ??= version();
  },
  warmUp: false,
  inProcess: false,
  buildFile: (fixture: string) => convertAndRead(recalc, fixture),
} satisfies Engine;

if (import.meta.main) {
  checkPrerequisites();
  if (process.argv.includes('--verify-recalc')) await verifyRecalc(recalc, 'gridcraft-cli convert');
  else await replay(engine, import.meta.url);
}
