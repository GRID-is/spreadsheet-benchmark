// Whether each upstream workbook of a suite has a fixture made from its current content, told by
// the content hash that its excel-pass.json entry records; and the digest of a suite's fixtures,
// which a report records so that the baseline check can tell whether they changed since.
import { createHash } from 'node:crypto';
import { existsSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { ZipArchive } from '@borgar/zip';
import { fixturePath } from './fixture.ts';

/** The fields of an excel-pass.json entry that tell whether its fixture is current. */
export type PassState = { fixture?: string, error?: string, upstreamHash?: string };
type PassEntries = Record<string, PassState>;

function readJson<T> (path: string, fallback: T): T {
  return existsSync(path) ? JSON.parse(readFileSync(path, 'utf8')) : fallback;
}

/**
 * The SHA-256 of a workbook's content. For a zip (xlsx), that is of its entries' names and
 * contents, not of the archive, whose entry times change when an import writes the same workbook
 * again. `normalize`, if given, maps each entry's content before it is hashed.
 */
export async function workbookHash (path: string, normalize?: (entry: string, content: Uint8Array) => Uint8Array): Promise<string> {
  const bytes = readFileSync(path);
  const hash = createHash('sha256');
  if (bytes.subarray(0, 4).toString('latin1') !== 'PK\x03\x04') {
    return hash.update(bytes).digest('hex');
  }
  const zip = new ZipArchive(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
  for (const entry of [ ...zip.files ].sort()) {
    const content = await zip.read(entry);
    if (content != null) {
      const data = normalize?.(entry, new Uint8Array(content)) ?? new Uint8Array(content);
      hash.update(`${entry}\0${data.byteLength}\0`).update(data);
    }
  }
  return hash.digest('hex');
}

/**
 * Whether an excel-pass.json entry holds a fixture of the upstream workbook whose content hash is
 * `hash`. An entry without a hash, written before the Excel pass recorded one, counts as current,
 * so that the committed suites stay valid without Excel; recordUpstreamHashes gives it the hash
 * before an import can change its workbook. The cost: a change to its workbook made outside an
 * import, or left by an import that failed before hashes were recorded, goes unnoticed until the
 * workbook's next Excel pass.
 */
export function isPassCurrent (entry: PassState | undefined, hash: string): boolean {
  return entry != null && entry.error == null && (entry.upstreamHash == null || entry.upstreamHash === hash);
}

/**
 * Records, in each excel-pass.json entry that has none, the content hash of its upstream workbook
 * as it stands. The import calls it before the importer, which first deletes the workbooks, so
 * that a change the import makes, even one it leaves half done by failing, differs from the hash.
 */
export async function recordUpstreamHashes (suiteDir: string): Promise<void> {
  const passFile = join(suiteDir, 'excel-pass.json');
  const entries = readJson<PassEntries>(passFile, {});
  let recorded = false;
  for (const [ file, entry ] of Object.entries(entries)) {
    const path = join(suiteDir, 'workbooks', file);
    if (entry.upstreamHash == null && existsSync(path)) {
      entry.upstreamHash = await workbookHash(path);
      recorded = true;
    }
  }
  if (recorded) {
    writeFileSync(passFile, JSON.stringify(entries, null, 2) + '\n');
  }
}

/**
 * Drops the excel-pass.json entries of the workbooks that are no longer in the suite's manifest and
 * deletes their fixtures, except one that a remaining workbook shares, and returns those workbooks.
 */
export function forgetRemovedWorkbooks (suiteDir: string): string[] {
  const passFile = join(suiteDir, 'excel-pass.json');
  const entries = readJson<PassEntries>(passFile, {});
  const manifest = readJson<{ files: Array<{ file: string }> }>(join(suiteDir, 'manifest.json'), { files: [] });
  const kept = new Set(manifest.files.map(({ file }) => file));
  const removed = Object.keys(entries).filter(file => !kept.has(file)).sort();
  const fixtureOf = (file: string) => entries[file].fixture ?? fixturePath(file);
  const keptFixtures = new Set(Object.keys(entries).filter(file => kept.has(file)).map(fixtureOf));
  for (const file of removed) {
    if (!keptFixtures.has(fixtureOf(file))) {
      rmSync(join(suiteDir, 'excel', fixtureOf(file)), { force: true });
    }
    delete entries[file];
  }
  if (removed.length) {
    writeFileSync(passFile, JSON.stringify(entries, null, 2) + '\n');
  }
  return removed;
}

/**
 * The workbooks of a suite's manifest that have no fixture to measure: no excel-pass.json entry,
 * an entry for a failed Excel pass, a fixture file missing, or a fixture of other content than the
 * upstream workbook's (see isPassCurrent).
 */
export async function workbooksWithoutFixture (suiteDir: string): Promise<string[]> {
  const manifest = readJson<{ files: Array<{ file: string }> }>(join(suiteDir, 'manifest.json'), { files: [] });
  const entries = readJson<PassEntries>(join(suiteDir, 'excel-pass.json'), {});
  const missing: string[] = [];
  for (const { file } of manifest.files) {
    const entry = entries[file];
    const upstream = join(suiteDir, 'workbooks', file);
    // isPassCurrent passes an entry without a hash whatever the hash, so its workbook is unread.
    const hash = entry?.upstreamHash != null && existsSync(upstream) ? await workbookHash(upstream) : '';
    if (!isPassCurrent(entry, hash) || !existsSync(join(suiteDir, 'excel', entry?.fixture ?? fixturePath(file)))) {
      missing.push(file);
    }
  }
  return missing;
}

/**
 * The SHA-256 of a suite's fixtures: of the path and bytes of every file in its excel/ folder,
 * leaving out dotfiles. The Excel pass puts back the bytes of a fixture whose content it left
 * unchanged, so the digest changes only with a fixture's content, and a report that records it
 * tells whether it was measured on the fixtures as they stand.
 */
export function fixturesDigest (suiteDir: string): string {
  const excelDir = join(suiteDir, 'excel');
  const hash = createHash('sha256');
  if (!existsSync(excelDir)) {
    return hash.digest('hex');
  }
  const files = readdirSync(excelDir, { recursive: true, withFileTypes: true })
    .filter(entry => entry.isFile() && !entry.name.startsWith('.'))
    .map(entry => join(entry.parentPath, entry.name).slice(excelDir.length + 1))
    .sort();
  for (const file of files) {
    const bytes = readFileSync(join(excelDir, file));
    hash.update(`${file}\0${bytes.byteLength}\0`).update(bytes);
  }
  return hash.digest('hex');
}
