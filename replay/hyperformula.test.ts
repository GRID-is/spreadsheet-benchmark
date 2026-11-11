import assert from 'node:assert/strict';
import { test } from 'node:test';
import { harnessTiming, type InputSheet } from './common.ts';

// HyperFormula is installed ad hoc (replay/README.md), so these tests run only where it is.
const installed = await import('hyperformula').then(() => true, () => false);

test('HyperFormula evaluates an array formula as an array and an ordinary formula with implicit intersection', { skip: !installed && 'hyperformula is not installed' }, async () => {
  // Arrange
  const { HyperFormula } = await import('hyperformula');
  const { hyperFormulaSheets } = await import('./hyperformula.ts');
  const input: InputSheet[] = [ { name: 'Sheet1', cells: {
    A1: { v: 1 }, B1: { v: 2 }, C1: { v: 3 },
    A3: { f: 'A1:C1*2', v: 2, F: 'A3:C3' },
    A5: { f: 'SUM(A1:C1*A1:C1)', v: 14, F: 'A5' },
    A7: { f: 'TRANSPOSE(A1:C1)', v: 1, F: 'A7:A9' },
    B10: { f: 'A1:C1*2', v: 4 },
    D1: { f: 'SUM(A3:C3)', v: 12 },
  } } ];

  // Act
  const hf = HyperFormula.buildFromSheets(hyperFormulaSheets(input), { licenseKey: 'gpl-v3' });
  const value = (ref: string) => {
    const [ , col, row ] = /^([A-Z])(\d+)$/.exec(ref) ?? [];
    return hf.getCellValue({ sheet: 0, row: Number(row) - 1, col: col.charCodeAt(0) - 65 });
  };

  // Assert
  assert.deepEqual([ value('A3'), value('B3'), value('C3') ], [ 2, 4, 6 ], 'a dynamic array formula spills its array');
  assert.equal(value('D1'), 12, 'a formula over the spill range reads the spilled values');
  assert.equal(value('A5'), 14, 'a CSE array formula evaluates its expression as an array');
  assert.deepEqual([ value('A7'), value('A8'), value('A9') ], [ 1, 2, 3 ], 'a function that returns an array spills the same array in ARRAYFORMULA');
  assert.equal(value('B10'), 4, 'an ordinary formula intersects a range in a scalar argument with its own column');
});

test('buildHyperFormula gives an array formula that HyperFormula throws on as an ordinary formula, and keeps the others in ARRAYFORMULA', { skip: !installed && 'hyperformula is not installed' }, async () => {
  // Arrange: HyperFormula 3.5.0 throws while building XLOOKUP of an array in ARRAYFORMULA
  const { buildHyperFormula } = await import('./hyperformula.ts');
  const input: InputSheet[] = [ { name: 'Sheet1', cells: {
    A1: { v: 1 }, A2: { v: 2 }, B1: { v: 10 }, B2: { v: 20 },
    C1: { f: 'XLOOKUP(A1:A2,A1:A2,B1:B2)', v: 10, F: 'C1:C2' },
    D1: { f: 'A1:A2*2', v: 2, F: 'D1:D2' },
  } } ];

  // Act
  const hf = buildHyperFormula(input);
  const value = (row: number, col: number) => hf.getCellValue({ sheet: 0, row, col });

  // Assert
  assert.equal(value(0, 2), 10, 'the formula that throws in ARRAYFORMULA is evaluated as an ordinary formula');
  assert.deepEqual([ value(0, 3), value(1, 3) ], [ 2, 4 ], 'another array formula of the workbook still spills its array');
});

test('buildHyperFormula counts the builds before its last one as harness time', { skip: !installed && 'hyperformula is not installed' }, async t => {
  // Arrange: a clock that advances 10 ms at each reading, and a formula that makes the first build throw
  const { buildHyperFormula } = await import('./hyperformula.ts');
  let now = 0;
  t.mock.method(Date, 'now', () => (now += 10));
  const input: InputSheet[] = [ { name: 'Sheet1', cells: {
    A1: { v: 1 }, A2: { v: 2 }, B1: { v: 10 }, B2: { v: 20 },
    C1: { f: 'XLOOKUP(A1:A2,A1:A2,B1:B2)', v: 10, F: 'C1:C2' },
  } } ];
  harnessTiming.ms = 0;

  // Act
  buildHyperFormula(input);

  // Assert
  assert.equal(harnessTiming.ms, 10, 'the time from the first build to the last one is the harness\'s');
});
