// Replay a report's cases through Gnumeric, via its ssconvert CLI. Pattern:
// convert-and-read, like google-sheets.ts - ssconvert --recalc opens the
// fixture with Gnumeric's own reader, recalculates every formula, and writes
// an xlsx whose cached values are Gnumeric's answers, read back with
// cachedValues. Nothing is masked by Apiary's parser on the way in: Gnumeric
// parses the original file itself.
//
// --recalc is load-bearing and verified, not assumed. Without it the cached
// values Excel wrote would pass straight through and the engine would score a
// meaningless ~100%; `node replay/gnumeric.ts --verify-recalc` proves the
// opposite by feeding ssconvert a workbook whose cached value is deliberately
// wrong and checking that the answer comes back right.
//
// buildMs covers an ssconvert process spawn, Gnumeric's own parse and an
// xlsx write, so it is not comparable to an in-process engine's build time
// (see the Performance tab's caveat panel). Reading that xlsx back is
// harness time.
import { spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { join } from 'node:path';
import { convertAndRead, guardConvertedFileExists, replay, type Engine } from './common.ts';
import { runLimited } from './child-process.ts';
import { verifyRecalc } from './verify-recalc.ts';

const SSCONVERT = process.env.SSCONVERT
  ?? [ '/opt/homebrew/bin/ssconvert', '/usr/local/bin/ssconvert' ].find(p => existsSync(p))
  ?? 'ssconvert';

function version (): string {
  const res = spawnSync(SSCONVERT, [ '--version' ], { encoding: 'utf8' });
  const m = /(\d+\.\d+\.\d+)/.exec(`${res.stdout}${res.stderr}`);
  return m ? m[1] : 'unknown';
}

// Convert one workbook and hand back the path of the recalculated copy.
async function recalc (fixture: string, dir: string): Promise<string> {
  const out = join(dir, 'out.xlsx');
  const budget = Math.max(5000, Number(process.env.REPLAY_TIMEOUT_MS ?? 60000) - 5000);
  const res = await runLimited(SSCONVERT, [ '--recalc', fixture, out ], { timeoutMs: budget });
  if (res.error) throw res.error;
  return guardConvertedFileExists(out, 'ssconvert', res);
}

/**
 * Throws unless ssconvert starts, which the runner checks before it replays anything: without
 * it, the version reads as unknown.
 */
export function checkPrerequisites (): void {
  if (spawnSync(SSCONVERT, [ '--version' ]).error) throw new Error('ssconvert not found (brew install gnumeric, or set SSCONVERT)');
}

let probedVersion: string | undefined;

export const engine = {
  name: 'gnumeric',
  // Probed when first read, which only the replay's own process does: the job process of each
  // fixture imports this module, and starting ssconvert there would slow every fixture.
  get version () {
    return probedVersion ??= version();
  },
  warmUp: false,
  inProcess: false,
  buildFile: (fixture: string) => convertAndRead(recalc, fixture),
} satisfies Engine;

if (import.meta.main) {
  checkPrerequisites();
  if (process.argv.includes('--verify-recalc')) await verifyRecalc(recalc, 'ssconvert --recalc');
  else await replay(engine, import.meta.url);
}
