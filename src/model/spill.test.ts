import { test } from 'node:test';
import assert from 'node:assert/strict';
import { areaRange, areaRefs, parseArea, spillArea } from './spill.ts';

test('parseArea reads a range or a cell, and nothing else', () => {
  assert.deepEqual(parseArea('C1:D3'), { top: 0, left: 2, bottom: 2, right: 3 });
  assert.deepEqual(parseArea('AA10'), { top: 9, left: 26, bottom: 9, right: 26 });
  assert.deepEqual(parseArea('$B$2:$A$1'), { top: 0, left: 0, bottom: 1, right: 1 }, 'absolute and reversed corners');
  assert.equal(parseArea('Sheet1!A1:B2'), undefined, 'a sheet-qualified range');
  assert.equal(parseArea('A:A'), undefined, 'a whole column');
});

test('areaRange writes a single cell without a colon', () => {
  assert.equal(areaRange({ top: 0, left: 2, bottom: 2, right: 3 }), 'C1:D3');
  assert.equal(areaRange({ top: 4, left: 0, bottom: 4, right: 0 }), 'A5');
  assert.equal(areaRange(spillArea('B2', 'B2:B2')), 'B2', 'a one-cell range is the cell');
});

test('areaRefs lists cells by row', () => {
  assert.deepEqual(areaRefs({ top: 0, left: 25, bottom: 1, right: 26 }), [ [ 'Z1', 'AA1' ], [ 'Z2', 'AA2' ] ]);
});

test('spillArea is the anchor alone where the range is absent, unreadable or anchored elsewhere', () => {
  assert.equal(areaRange(spillArea('C1', 'C1:D3')), 'C1:D3');
  assert.equal(areaRange(spillArea('C1', undefined)), 'C1');
  assert.equal(areaRange(spillArea('C1', null)), 'C1');
  assert.equal(areaRange(spillArea('C1', 'A:A')), 'C1');
  assert.equal(areaRange(spillArea('C2', 'C1:D3')), 'C2');
});
