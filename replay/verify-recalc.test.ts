import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { test } from 'node:test';
import { ZipArchive } from '@borgar/zip';
import { cachedValues } from './common.ts';
import { CANARY } from './verify-recalc.ts';

test('the recalc canary caches 99999 for A3 =A1+A2 over 1 and 2', async () => {
  // Arrange
  const expected = { A1: 1, A2: 2, A3: 99999 };

  // Act
  const value = await cachedValues(CANARY);

  // Assert
  assert.deepEqual(Object.fromEntries(Object.keys(expected).map(ref => [ ref, value({ sheet: 'Sheet1', ref }) ])), expected);
});

test('the recalc canary calculation chain references only existing formula cells', async () => {
  const bytes = readFileSync(CANARY);
  const workbook = new ZipArchive(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
  const worksheet = await workbook.readText('xl/worksheets/sheet1.xml');
  assert.ok(worksheet);
  const formulas = new Set(Array.from(worksheet.matchAll(/<c\b[^>]*\br="([^"]+)"[^>]*>([\s\S]*?)<\/c>/g))
    .filter(([, , content ]) => /<f[\s>]/.test(content))
    .map(([, ref ]) => ref));
  const chain = await workbook.readText('xl/calcChain.xml');
  for (const [, ref ] of chain?.matchAll(/<c\b[^>]*\br="([^"]+)"/g) ?? []) {
    assert.ok(formulas.has(ref), `calculation chain references ${ref}, which has no formula`);
  }
});
