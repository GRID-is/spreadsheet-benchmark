import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DEFAULT_EPS, functionNames, isVolatile, numbersAreClose, observeSpill, valuesMatch } from './compare.ts';

test('isVolatile flags nondeterministic functions only', () => {
  // Arrange
  const volatile = [ 'RAND()', 'ROUND(RANDBETWEEN(1,9),0)', 'RANDARRAY(2,2)', 'NOW()-1', 'TODAY()' ];
  const stable = [ 'OFFSET(A1,1,1)', 'INDIRECT("A1")', 'SUM(A1:A3)', 'A1*B1' ];

  // Act & Assert
  for (const f of volatile) {
    assert.equal(isVolatile(f), true, f);
  }
  for (const f of stable) {
    assert.equal(isVolatile(f), false, f);
  }
});

test('numbersAreClose follows IronCalc semantics', () => {
  // Arrange
  const eps = 5e-8;

  // Act & Assert
  assert.equal(numbersAreClose(0, 0, eps), true);
  assert.equal(numbersAreClose(1, 1 + 1e-9, eps), true);
  assert.equal(numbersAreClose(1e12, 1e12 + 1e3, eps), true);
  assert.equal(numbersAreClose(1, 1.001, eps), false);
  assert.equal(numbersAreClose(0, 1e-9, eps), true);
});

test('valuesMatch compares by type', () => {
  // Arrange
  const eps = 5e-8;

  // Act & Assert
  assert.equal(valuesMatch('#DIV/0!', '#DIV/0!', eps), true);
  assert.equal(valuesMatch('#DIV/0!', '#VALUE!', eps), false);
  assert.equal(valuesMatch(true, true, eps), true);
  assert.equal(valuesMatch(1, '1', eps), false);
  assert.equal(valuesMatch(null, undefined, eps), true);
  assert.equal(valuesMatch('', null, eps), true);
  assert.equal(valuesMatch(null, 0, eps), false);
});

test('functionNames extracts every distinct function in a formula', () => {
  // Arrange
  const formula = 'SUM(A1:A3)+_xlfn.XLOOKUP(1,B:B,C:C,IF(x,1,2))*Sheet1!D1';
  const expected = [ 'SUM', 'XLOOKUP', 'IF' ];

  // Act
  const names = functionNames(formula);

  // Assert
  assert.deepEqual(names, expected);
  assert.deepEqual(functionNames('A1*B1'), []);
});

test('observeSpill records both spill ranges and the values over Excel\'s', () => {
  const excel = { C1: 2, C2: 4, C3: 6 };
  const engine = { C1: 2, C2: 4, C3: 7, C4: 8 };

  const spill = observeSpill('C1', { range: 'C1:C3', value: ref => excel[ref as keyof typeof excel] }, { range: 'C1:C4', value: ref => engine[ref as keyof typeof engine] }, DEFAULT_EPS);

  assert.deepEqual(spill, {
    excelRange: 'C1:C3',
    excelValues: [ [ 2 ], [ 4 ], [ 6 ] ],
    apiaryRange: 'C1:C4',
    apiaryValues: [ [ 2 ], [ 4 ], [ 7 ] ],
    mismatches: 1,
  });
});

test('observeSpill counts mismatches over the spilled cells, not the anchor', () => {
  const spill = observeSpill('A1', { range: 'A1:B1', value: ref => ref === 'A1' ? 1 : 2 }, { value: () => 9 }, DEFAULT_EPS);

  assert.equal(spill?.mismatches, 1, 'B1 differs; A1, the anchor, is the case\'s own value');
  assert.equal(spill?.apiaryRange, undefined, 'no engine range where the engine reports none');
});

test('observeSpill is undefined where neither answer spills', () => {
  assert.equal(observeSpill('A1', { value: () => 1 }, { range: 'A1', value: () => 1 }, DEFAULT_EPS), undefined, 'both single cells');
  assert.equal(observeSpill('A1', { range: 'A1:A1', value: () => 1 }, { value: () => 1 }, DEFAULT_EPS), undefined, 'a one-cell array formula');
});

test('observeSpill records a spill of the engine alone, over the anchor', () => {
  const spill = observeSpill('A1', { value: () => 1 }, { range: 'A1:A2', value: () => 1 }, DEFAULT_EPS);

  assert.deepEqual(spill, { excelRange: 'A1', excelValues: [ [ 1 ] ], apiaryRange: 'A1:A2', apiaryValues: [ [ 1 ] ], mismatches: 0 });
});

test('observeSpill compares empty and absent values as valuesMatch does', () => {
  const spill = observeSpill('A1', { range: 'A1:A3', value: ref => ref === 'A2' ? '' : 1 }, { value: ref => ref === 'A2' ? undefined : 1 + 1e-12 }, DEFAULT_EPS);

  assert.equal(spill?.mismatches, 0);
  assert.deepEqual(spill?.apiaryValues, [ [ 1 + 1e-12 ], [ null ], [ 1 + 1e-12 ] ], 'an absent value is recorded as null');
});
