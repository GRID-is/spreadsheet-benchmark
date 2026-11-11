// Score Google Sheets against the Excel oracle from Drive round-trip exports:
// each fixture is uploaded to Drive with convert-to-Google-Sheet (conversion
// is the recalculation), exported back as xlsx (Google's computed values ride
// along as cached values), and this engine reads those exports. The round
// trip itself is gsheets_drive.py, over the Drive API; GSHEETS_EXPORT_DIR names
// the directory of exports, holding one subdirectory per suite laid out like
// that suite's excel dir. No timings:
// wall time here is network and conversion queue, not engine compute.
//
// Each export has a capture identity beside it (<export>.json, written by
// gsheets_drive.py): the SHA-256 of the fixture uploaded and of the export
// received, and the UTC date of the export. An export without one, whose
// fixture or bytes have changed since, or made before the version,
// GSHEETS_SNAPSHOT or else today, fails its workbook rather than scoring a
// different workbook or Google Sheets of an earlier date.
//
// Where Google failed to export the whole workbook, gsheets_drive.py exports it without the
// formulas of the sheets whose calculation fails, and names those sheets in the identity
// (`droppedSheets`), and the sheets whose formulas may read those (`dependentSheets`), which would
// then read Excel's values. The cases of both fail as load errors here, and the manifest's engine
// options list both kinds of sheet for every export.
import { createHash } from 'node:crypto';
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { cachedValues, replay, type CaseCell, type Engine } from './common.ts';

const EXPORT_DIR = resolve(process.env.GSHEETS_EXPORT_DIR ?? '.scratch/gsheets-exports');
// Set in the environment, so that the fixture workers, forked with it, check against this process's
// date rather than one their own clocks give past midnight.
process.env.GSHEETS_SNAPSHOT ??= new Date().toISOString().slice(0, 10);
const SNAPSHOT = process.env.GSHEETS_SNAPSHOT;

function sha256 (path: string): string {
  return createHash('sha256').update(readFileSync(path)).digest('hex');
}

type Identity = { sourceSha256?: string, exportSha256?: string, exportedOn?: string, droppedSheets?: string[], dependentSheets?: string[] };

/** The sheets of that kind of each export in EXPORT_DIR made on or after SNAPSHOT that has any, by
 * its suite-relative path. An earlier export is not scored, so its sheets are not listed. */
function sheetsOfExports (kind: 'droppedSheets' | 'dependentSheets'): Record<string, string[]> {
  const out: Record<string, string[]> = {};
  if (!existsSync(EXPORT_DIR)) return out;
  for (const entry of readdirSync(EXPORT_DIR, { recursive: true, encoding: 'utf8' }).sort()) {
    if (!entry.endsWith('.xlsx.json')) continue;
    let identity: Identity;
    try {
      identity = JSON.parse(readFileSync(join(EXPORT_DIR, entry), 'utf8'));
    }
    catch {
      continue; // buildFile fails that workbook
    }
    const sheets = identity[kind];
    if ((identity.exportedOn ?? '') >= SNAPSHOT && sheets?.length) out[entry.slice(0, -'.json'.length)] = sheets;
  }
  return out;
}

const options: Record<string, string> = { exportDir: EXPORT_DIR };

export const engine = {
  name: 'google-sheets',
  version: SNAPSHOT,
  warmUp: false,
  inProcess: false,
  options,
  noTimings: true,
  async buildFile (fixture: string) {
    const m = /\/suites\/([^/]+)\/excel\/(.+)$/.exec(fixture);
    if (!m) {
      throw new Error(`cannot derive suite-relative path from ${fixture}`);
    }
    const exported = join(EXPORT_DIR, m[1], m[2]);
    if (!existsSync(exported)) {
      throw new Error(`no Google Sheets export for ${m[1]}/${m[2]}`);
    }
    if (!existsSync(`${exported}.json`)) {
      throw new Error(`Google Sheets export for ${m[1]}/${m[2]} has no capture identity; re-export it`);
    }
    const identity: Identity = JSON.parse(readFileSync(`${exported}.json`, 'utf8'));
    if (identity.sourceSha256 !== sha256(fixture) || identity.exportSha256 !== sha256(exported)) {
      throw new Error(`Google Sheets export for ${m[1]}/${m[2]} does not match its capture identity; re-export it`);
    }
    // On or after, as `run --since` has it: a round trip that runs past midnight exports the rest
    // of the fixtures on the next day.
    if ((identity.exportedOn ?? '') < SNAPSHOT) {
      throw new Error(`Google Sheets export for ${m[1]}/${m[2]} was made on ${identity.exportedOn ?? 'an unrecorded date'}, before ${SNAPSHOT}; re-export it with gsheets_drive.py run --since ${SNAPSHOT}`);
    }
    const values = await cachedValues(exported);
    const dropped = new Set(identity.droppedSheets ?? []);
    const dependent = new Set(identity.dependentSheets ?? []);
    if (dropped.size === 0 && dependent.size === 0) return values;
    return Object.assign(values, {
      loadError: (c: CaseCell) => dropped.has(c.sheet)
        ? `Google Sheets failed to export sheet ${c.sheet}`
        : dependent.has(c.sheet) ? `sheet ${c.sheet} may read a sheet that Google Sheets failed to export` : undefined,
    });
  },
} satisfies Engine;

if (import.meta.main) {
  // Only the replay's own process writes the manifest, so only it reads every identity.
  for (const kind of [ 'droppedSheets', 'dependentSheets' ] as const) {
    const sheets = sheetsOfExports(kind);
    if (Object.keys(sheets).length) options[kind] = JSON.stringify(sheets);
  }
  await replay(engine, import.meta.url);
}
