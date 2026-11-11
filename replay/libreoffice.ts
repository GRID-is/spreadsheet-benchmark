// Replay a report's cases through LibreOffice Calc, via soffice --headless.
// Pattern: convert-and-read, like gnumeric.ts - LibreOffice opens the fixture
// with its own reader, recalculates every formula, and writes an xlsx whose
// cached values are LibreOffice's answers, read back with cachedValues.
// Nothing is masked by Apiary's parser on the way in: LibreOffice parses the
// original file itself.
//
// Recalculation is forced, not assumed. Calc's default "Recalculation on file
// load" for Excel-authored xlsx is "never", so a plain --convert-to passes
// Excel's cached values straight through and the engine would score a
// meaningless ~100%. The runner therefore starts soffice with its own user
// profile whose registrymodifications.xcu sets OOXMLRecalcMode (and
// ODFRecalcMode) to 0 = always recalculate. `node replay/libreoffice.ts
// --verify-recalc` proves it by feeding soffice a workbook whose cached value
// is deliberately wrong and checking that the answer comes back right; run it
// after any LibreOffice upgrade.
//
// buildMs covers a soffice process spawn, LibreOffice's own parse and an
// xlsx write, so it is not comparable to an in-process engine's build time
// (see the Performance tab's caveat panel). Reading that xlsx back is
// harness time.
import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { convertAndRead, guardConvertedFileExists, replay, type Engine } from './common.ts';
import { runLimited } from './child-process.ts';
import { verifyRecalc } from './verify-recalc.ts';

const SOFFICE = [
  process.env.SOFFICE,
  '/Applications/LibreOffice.app/Contents/MacOS/soffice',
  '/opt/homebrew/bin/soffice',
  '/usr/local/bin/soffice',
  '/usr/bin/soffice',
  '/usr/bin/libreoffice',
].find(p => p && (p === basename(p) || existsSync(p))) ?? 'soffice';

const RECALC_XCU = `<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
`;

// profile() returns the file URL of the replay's LibreOffice user profile, which sets soffice to
// recalculate every formula on load. The replay's own process makes the profile and names it
// through PROFILE_ENV to the job processes that evaluate its files, so that soffice does not set up
// a fresh profile, over a second's work, for every file. The job processes run one at a time, and
// soffice locks its profile, so concurrent instances must never share one. The tmp dir is left to
// the OS.
const PROFILE_ENV = 'REPLAY_LIBREOFFICE_PROFILE';
function profile (): string {
  let profileDir = process.env[PROFILE_ENV];
  if (!profileDir) {
    profileDir = mkdtempSync(join(tmpdir(), 'libreoffice-profile-'));
    mkdirSync(join(profileDir, 'user'), { recursive: true });
    writeFileSync(join(profileDir, 'user', 'registrymodifications.xcu'), RECALC_XCU);
    process.env[PROFILE_ENV] = profileDir;
  }
  return pathToFileURL(profileDir).href;
}

function version (): string {
  const res = spawnSync(SOFFICE, [ '--version' ], { encoding: 'utf8' });
  const m = /LibreOffice\s+(\d+\.\d+\.\d+(?:\.\d+)?)/.exec(`${res.stdout}${res.stderr}`);
  return m ? m[1] : 'unknown';
}

// Convert one workbook and hand back the path of the recalculated copy.
async function recalc (fixture: string, dir: string): Promise<string> {
  const budget = Math.max(5000, Number(process.env.REPLAY_TIMEOUT_MS ?? 60000) - 5000);
  const res = await runLimited(SOFFICE, [
    '--headless', '--norestore', '--nologo', '--nolockcheck',
    `-env:UserInstallation=${profile()}`,
    '--convert-to', 'xlsx', '--outdir', dir, fixture,
  ], { timeoutMs: budget });
  if (res.error) throw res.error;
  return guardConvertedFileExists(join(dir, basename(fixture).replace(/\.[^.]+$/, '') + '.xlsx'), 'soffice', res);
}

/**
 * Throws unless soffice starts, which the runner checks before it replays anything: without
 * it, the version reads as unknown.
 */
export function checkPrerequisites (): void {
  if (spawnSync(SOFFICE, [ '--version' ]).error) throw new Error('soffice not found (brew install --cask libreoffice, or set SOFFICE)');
}

let probedVersion: string | undefined;

export const engine = {
  name: 'libreoffice',
  // Probed when first read, which only the replay's own process does: the job process of each
  // fixture imports this module, and starting soffice there would slow every fixture.
  get version () {
    return probedVersion ??= version();
  },
  warmUp: false,
  inProcess: false,
  buildFile: (fixture: string) => convertAndRead(recalc, fixture),
} satisfies Engine;

if (import.meta.main) {
  checkPrerequisites();
  if (process.argv.includes('--verify-recalc')) await verifyRecalc(recalc, 'soffice (forced-recalc profile)');
  else {
    profile();
    await replay(engine, import.meta.url);
  }
}
