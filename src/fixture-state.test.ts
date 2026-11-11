import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, readFileSync, renameSync, rmSync, writeFileSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { ZipArchive } from '@borgar/zip';
import { fixturesDigest, forgetRemovedWorkbooks, recordUpstreamHashes, workbookHash, workbooksWithoutFixture } from './fixture-state.ts';

function write (path: string, content: string): void {
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, content);
}

function suite (t: { after: (fn: () => void) => void }): string {
  const dir = mkdtempSync(join(tmpdir(), 'fixture-state-'));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  for (const file of [ 'a.xlsx', 'b/c.xls', 'gone.xlsx' ]) {
    write(join(dir, 'workbooks', file), `upstream ${file}`);
    write(join(dir, 'excel', file.replace(/\.xls$/, '.xlsx')), `fixture ${file}`);
  }
  write(join(dir, 'excel-pass.json'), JSON.stringify({
    'a.xlsx': { fixture: 'a.xlsx' },
    'b/c.xls': { fixture: 'b/c.xlsx' },
    'gone.xlsx': { fixture: 'gone.xlsx' },
  }));
  write(join(dir, 'manifest.json'), JSON.stringify({ files: [ { file: 'a.xlsx' }, { file: 'b/c.xls' }, { file: 'gone.xlsx' } ] }));
  return dir;
}

const readPass = (dir: string) => JSON.parse(readFileSync(join(dir, 'excel-pass.json'), 'utf8'));

test('a workbook whose content differs from the hash its entry records awaits the Excel pass', async t => {
  // Arrange: entries that record the hashes of their workbooks, as the Excel pass writes them, then
  // a change with no snapshot from before it, as a failed import leaves one.
  const dir = suite(t);
  const pass = readPass(dir);
  for (const file of Object.keys(pass)) {
    pass[file].upstreamHash = await workbookHash(join(dir, 'workbooks', file));
  }
  write(join(dir, 'excel-pass.json'), JSON.stringify(pass));
  write(join(dir, 'workbooks', 'b/c.xls'), 'changed upstream');

  // Assert
  assert.deepEqual(await workbooksWithoutFixture(dir), [ 'b/c.xls' ], 'only the changed workbook awaits the Excel pass');
});

test('an entry without a hash counts as current', async t => {
  const dir = suite(t);

  assert.deepEqual(await workbooksWithoutFixture(dir), [], 'no workbook awaits the Excel pass');
});

test('an import records the hashes that entries lack before it changes their workbooks', async t => {
  // Arrange
  const dir = suite(t);
  const original = await workbookHash(join(dir, 'workbooks', 'b/c.xls'));

  // Act: the import records the hashes, then changes, adds and removes workbooks.
  await recordUpstreamHashes(dir);
  write(join(dir, 'workbooks', 'b/c.xls'), 'changed upstream');
  write(join(dir, 'workbooks', 'new.xlsx'), 'new upstream');
  rmSync(join(dir, 'workbooks', 'gone.xlsx'));
  write(join(dir, 'manifest.json'), JSON.stringify({ files: [ { file: 'a.xlsx' }, { file: 'b/c.xls' }, { file: 'new.xlsx' } ] }));
  const removed = forgetRemovedWorkbooks(dir);

  // Assert
  assert.equal(readPass(dir)['b/c.xls'].upstreamHash, original, 'the entry records the hash of the workbook before the import');
  assert.deepEqual(removed, [ 'gone.xlsx' ], 'the workbook no longer in the manifest is removed');
  assert.deepEqual(Object.keys(readPass(dir)), [ 'a.xlsx', 'b/c.xls' ], 'and its entry dropped');
  assert.equal(existsSync(join(dir, 'excel', 'gone.xlsx')), false, 'and its fixture deleted');
  assert.deepEqual(await workbooksWithoutFixture(dir), [ 'b/c.xls', 'new.xlsx' ], 'the changed and added workbooks await the Excel pass');
});

test('forgetRemovedWorkbooks keeps a fixture that a remaining workbook shares', t => {
  const dir = suite(t);
  write(join(dir, 'workbooks', 'b/c.xlsx'), 'upstream b/c.xlsx');
  write(join(dir, 'excel-pass.json'), JSON.stringify({ ...readPass(dir), 'b/c.xlsx': { fixture: 'b/c.xlsx' } }));
  write(join(dir, 'manifest.json'), JSON.stringify({ files: [ { file: 'a.xlsx' }, { file: 'b/c.xlsx' }, { file: 'gone.xlsx' } ] }));
  assert.deepEqual(forgetRemovedWorkbooks(dir), [ 'b/c.xls' ], 'the .xls workbook no longer in the manifest is removed');
  assert.equal(existsSync(join(dir, 'excel', 'b/c.xlsx')), true, 'the fixture it shares with b/c.xlsx is kept');
});

test('workbooksWithoutFixture counts a failed Excel pass and a missing fixture file', async t => {
  const dir = suite(t);
  assert.deepEqual(await workbooksWithoutFixture(dir), [], 'every workbook has its fixture');
  write(join(dir, 'excel-pass.json'), JSON.stringify({ 'a.xlsx': { fixture: 'a.xlsx', error: 'Excel failed' }, 'b/c.xls': { fixture: 'b/c.xlsx' }, 'gone.xlsx': { fixture: 'gone.xlsx' } }));
  rmSync(join(dir, 'excel', 'b/c.xlsx'));
  assert.deepEqual(await workbooksWithoutFixture(dir), [ 'a.xlsx', 'b/c.xls' ], 'a failed pass and a missing fixture leave a workbook without one');
});

test('a workbook written again with the same content keeps its hash, though its zip entry times differ', async t => {
  // Arrange
  const dir = mkdtempSync(join(tmpdir(), 'workbook-hash-'));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  const zip = new ZipArchive();
  await zip.write('xl/workbook.xml', '<workbook/>');
  const bytes = Buffer.from(zip.toArrayBuffer());
  const later = Buffer.from(bytes);
  // The modification time of an entry: offset 10 of a local file header, 12 of a central
  // directory header.
  for (const [ signature, offset ] of [ [ 'PK\x03\x04', 10 ], [ 'PK\x01\x02', 12 ] ] as const) {
    const at = later.indexOf(signature, 0, 'latin1');
    later.writeUInt16LE(later.readUInt16LE(at + offset) ^ 0x0421, at + offset);
  }
  writeFileSync(join(dir, 'a.xlsx'), bytes);
  writeFileSync(join(dir, 'b.xlsx'), later);
  const changed = new ZipArchive();
  await changed.write('xl/workbook.xml', '<workbook><sheets/></workbook>');
  writeFileSync(join(dir, 'c.xlsx'), Buffer.from(changed.toArrayBuffer()));

  // Assert
  assert.notDeepEqual(bytes, later, 'the archives differ in their entry times');
  assert.equal(await workbookHash(join(dir, 'b.xlsx')), await workbookHash(join(dir, 'a.xlsx')), 'the content hashes are equal');
  assert.notEqual(await workbookHash(join(dir, 'c.xlsx')), await workbookHash(join(dir, 'a.xlsx')), 'changed content changes the hash');
});

test('fixturesDigest changes with a fixture\'s bytes or path, not with a dotfile or upstream workbook', t => {
  const dir = suite(t);
  const digest = fixturesDigest(dir);

  write(join(dir, 'excel', '.DS_Store'), 'finder');
  write(join(dir, 'workbooks', 'a.xlsx'), 'upstream changed');
  assert.equal(fixturesDigest(dir), digest, 'unchanged by a dotfile or an upstream workbook');
  write(join(dir, 'excel', 'b', 'c.xlsx'), 'fixture saved again');
  assert.notEqual(fixturesDigest(dir), digest, 'changed by a fixture\'s bytes');
  write(join(dir, 'excel', 'b', 'c.xlsx'), 'fixture b/c.xls');
  assert.equal(fixturesDigest(dir), digest, 'back with the bytes');
  renameSync(join(dir, 'excel', 'a.xlsx'), join(dir, 'excel', 'a2.xlsx'));
  assert.notEqual(fixturesDigest(dir), digest, 'changed by a fixture\'s path');
  renameSync(join(dir, 'excel', 'a2.xlsx'), join(dir, 'excel', 'a.xlsx'));
  rmSync(join(dir, 'excel', 'gone.xlsx'));
  assert.notEqual(fixturesDigest(dir), digest, 'changed by a fixture removed');
});
