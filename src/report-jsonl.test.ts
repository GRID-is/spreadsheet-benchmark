import assert from 'node:assert/strict';
import { existsSync, mkdirSync, mkdtempSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test, type TestContext } from 'node:test';
import { setTimeout } from 'node:timers/promises';
import { caseGroups, readCases } from './report.ts';
import { compressJsonl, copyJsonl, jsonlLines, jsonlName, jsonlNames, readJsonl, readJsonlAsync, removeJsonl, storedPath, writeJsonl, writePlainJsonl } from './report-jsonl.ts';

const TEXT = '{"file":"a.xlsx","id":"a1"}\n{"file":"a.xlsx","id":"a2"}\n{"file":"b.xlsx","id":"b1"}\n';

function tempDir (t: TestContext): string {
  const dir = mkdtempSync(join(tmpdir(), 'report-jsonl-'));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  return dir;
}

async function lines (path: string): Promise<string[]> {
  const out: string[] = [];
  for await (const line of jsonlLines(path)) out.push(line);
  return out;
}

test('a compressed file reads as its plain text, through either path and every reader', async t => {
  // Arrange
  const path = join(tempDir(t), 'suite.jsonl');
  writeFileSync(path, TEXT);

  // Act
  compressJsonl(path);

  // Assert
  assert.equal(existsSync(path), false, 'compressing removes the plain form');
  assert.equal(storedPath(path), path + '.bz2', 'the stored form is the .bz2');
  assert.equal(readJsonl(path), TEXT, 'readJsonl decompresses');
  assert.equal(readJsonl(path + '.bz2'), TEXT, 'a path naming the .bz2 reads it too');
  assert.equal(await readJsonlAsync(path), TEXT, 'readJsonlAsync decompresses');
  assert.deepEqual(await lines(path), TEXT.trimEnd().split('\n'), 'jsonlLines streams the decompressed lines');
  assert.deepEqual(readCases(path).map(c => c.id), [ 'a1', 'a2', 'b1' ], 'readCases reads the compressed form');
  const groups: string[] = [];
  for await (const [ file ] of caseGroups(path)) groups.push(file);
  assert.deepEqual(groups, [ 'a.xlsx', 'b.xlsx' ], 'caseGroups reads the compressed form');
});

for (const [ count, compressed ] of [ [ 4096, false ], [ 1026, true ] ] as const) {
  for (const grouped of [ false, true ]) {
    test(`a buffered JSONL tail survives a busy consumer (${compressed ? 'compressed' : 'plain'}, ${grouped ? 'groups' : 'lines'})`, async t => {
      const path = join(tempDir(t), 'suite.jsonl');
      const expected = Array.from({ length: count }, (_, id) => JSON.stringify({ file: id === 0 ? 'a.xlsx' : 'b.xlsx', id: String(id) }));
      writeFileSync(path, expected.join('\n') + '\n');
      if (compressed) compressJsonl(path);
      const actual: string[] = [];
      if (grouped) {
        for await (const [ file, cases ] of caseGroups(path)) {
          actual.push(...cases.map(row => JSON.stringify(row)));
          if (file === 'a.xlsx') await setTimeout(50);
        }
      }
      else {
        for await (const line of jsonlLines(path)) {
          actual.push(line);
          if (actual.length === 1) await setTimeout(50);
        }
      }
      assert.deepEqual(actual, expected);
    });
  }
}

test('the plain form, which a run appends to, is read where both forms exist', async t => {
  // Arrange
  const path = join(tempDir(t), 'suite.jsonl');
  writeJsonl(path, 'compressed\n');

  // Act
  writeFileSync(path, 'plain\n');

  // Assert
  assert.equal(storedPath(path), path, 'the plain form is the stored one');
  assert.equal(readJsonl(path), 'plain\n', 'readJsonl reads the plain form');
  assert.deepEqual(await lines(path), [ 'plain' ], 'jsonlLines reads the plain form');
});

test('writing the plain form removes the compressed one', t => {
  // Arrange
  const path = join(tempDir(t), 'timings.jsonl');
  writeJsonl(path, 'first\n');

  // Act
  writePlainJsonl(path, 'second\n');

  // Assert
  assert.equal(existsSync(path + '.bz2'), false, 'the stale .bz2 is removed');
  assert.equal(readJsonl(path), 'second\n', 'the plain form holds the new text');
});

test('writeJsonl replaces the plain form with the compressed one', t => {
  // Arrange
  const path = join(tempDir(t), 'timings.jsonl');
  writePlainJsonl(path, 'first\n');

  // Act
  writeJsonl(path, 'second\n');

  // Assert
  assert.equal(existsSync(path), false, 'writeJsonl leaves no plain form');
  assert.equal(readJsonl(path), 'second\n', 'the .bz2 holds the new text');
});

test('removeJsonl removes both forms, after which reading fails as a missing file', t => {
  // Arrange
  const path = join(tempDir(t), 'timings.jsonl');
  writeJsonl(path, 'compressed\n');
  writeFileSync(path, 'plain\n');

  // Act
  removeJsonl(path);

  // Assert
  assert.equal(storedPath(path), null, 'neither form remains');
  assert.throws(() => readJsonl(path), { code: 'ENOENT' }, 'reading a missing file fails as a missing file');
});

test('copyJsonl copies the stored form and removes the other one in the target', t => {
  // Arrange
  const dir = tempDir(t);
  const from = join(dir, 'a', 'suite.jsonl');
  const to = join(dir, 'b');
  mkdirSync(join(dir, 'a'));
  mkdirSync(to);
  writeJsonl(from, TEXT);
  writeFileSync(join(to, 'suite.jsonl'), 'stale\n');

  // Act
  const copy = copyJsonl(from, to);

  // Assert
  assert.equal(copy, join(to, 'suite.jsonl.bz2'), 'the copy keeps the .bz2 suffix');
  assert.deepEqual(readdirSync(to), [ 'suite.jsonl.bz2' ], 'the stale plain form is gone');
  assert.equal(readJsonl(join(to, 'suite.jsonl')), TEXT, 'the copy reads as the original');
});

test('a corrupt compressed file fails the stream instead of ending it early', async t => {
  // Arrange
  const path = join(tempDir(t), 'suite.jsonl');
  writeFileSync(path + '.bz2', 'not bzip2');

  // Act & Assert
  await assert.rejects(lines(path), /bzip2 -dc/, 'the stream names the failed decompression');
});

test('jsonlName names a report JSONL file in either form, and nothing else', () => {
  // Act & Assert
  assert.equal(jsonlName('poi.jsonl'), 'poi', 'the plain form');
  assert.equal(jsonlName('poi.jsonl.bz2'), 'poi', 'the compressed form');
  assert.equal(jsonlName('poi.md'), null, 'a summary');
  assert.equal(jsonlName('manifest.json'), null, 'a manifest');
});

test('jsonlNames names each report JSONL file in a folder once, whichever forms it has', t => {
  // Arrange
  const dir = tempDir(t);
  for (const file of [ 'poi.jsonl', 'poi.jsonl.bz2', 'ironcalc.jsonl.bz2', 'timings.jsonl', 'poi.md', 'manifest.json' ]) {
    writeFileSync(join(dir, file), '');
  }

  // Act
  const names = jsonlNames(dir);

  // Assert
  assert.deepEqual(names, [ 'ironcalc', 'poi', 'timings' ]);
});
