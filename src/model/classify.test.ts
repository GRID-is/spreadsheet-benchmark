import { test } from 'node:test';
import assert from 'node:assert/strict';
import { type Case } from './case.ts';
import { type Catalogue } from './catalogue.ts';
import { classify } from './classify.ts';

const catalogue: Catalogue = {
  source: { apiaryVersion: 'test', generatedAt: '' },
  functions: {
    SUM: { name: 'SUM', inExcel: true, apiary: 'supported', legacy: false, nondeterministic: false },
    RAND: { name: 'RAND', inExcel: true, apiary: 'supported', legacy: false, nondeterministic: true },
    CUBEVALUE: { name: 'CUBEVALUE', inExcel: true, apiary: 'never', legacy: false, nondeterministic: false },
    B: { name: 'B', inExcel: false, apiary: 'unsupported', legacy: false, nondeterministic: false },
  },
};

function aCase (over: Partial<Case>): Case {
  return { suite: 's', category: 'c', file: 'f', sheet: 'S', ref: 'A1', id: 'x', formula: 'SUM(1)', functions: [ 'SUM' ], excelValue: 1, apiaryValue: 1, match: true, eps: 0, excelDroppedFormula: false, oracleApp: '', apiaryVersion: '', ...over };
}

test('classify derives status, parity and reasons from observations and the catalogue', () => {
  // Arrange
  const rows: Array<[ Case, string ]> = [
    [ aCase({}), 'pass' ],
    [ aCase({ match: false, apiaryValue: 2 }), 'fail' ],
    [ aCase({ match: false, apiaryValue: '#NAME?' }), 'unsupported' ],
    [ aCase({ loadError: 'timeout' }), 'error' ],
    [ aCase({ excelDroppedFormula: true }), 'skip:excel-dropped-formula' ],
    [ aCase({ fileFormula: '_xludf.SUM(1)', match: false }), 'skip:excel-read-as-user-function' ],
    [ aCase({ formula: '_xludf.SUM(1)', functions: [ '_XLUDF.SUM' ], match: false }), 'skip:excel-read-as-user-function' ],
    [ aCase({ fileFormula: '_xludf.B(1)' }), 'pass' ],
    [ aCase({ fileFormula: "'_xludf.SUM(SINGLE(A1))'!B2" }), 'pass' ],
    [ aCase({ functions: [ 'B' ], match: false }), 'skip:excel-lacks-function' ],
    [ aCase({ functions: [ 'SUM', 'B' ], match: false, apiaryValue: '#NAME?' }), 'unsupported' ],
    [ aCase({ functions: [ 'ROUND', 'RAND' ] }), 'skip:excel-lacks-function' ],
    [ aCase({ functions: [ 'SUM', 'RAND' ] }), 'skip:volatile' ],
    [ aCase({ formula: 'TEXT(A1,"yyyy")', volatilePrecedent: 'S!A1', match: false }), 'skip:volatile-precedent' ],
    [ aCase({ volatilePrecedent: '' }), 'skip:volatile-precedent' ],
    [ aCase({ functions: [ 'SUM', 'RAND' ], volatilePrecedent: 'S!A1' }), 'skip:volatile' ],
    [ aCase({ formula: 'CELL("filename")', match: false }), 'skip:environment' ],
    [ aCase({ formula: 'IF(1, info("OSVersion"), 2)', match: false }), 'skip:environment' ],
    [ aCase({ formula: 'INFO("recalc")' }), 'pass' ],
    [ aCase({ formula: 'INFO("system"&"atic")', match: false, apiaryValue: 2 }), 'fail' ],
    [ aCase({ formula: 'CELL("filename"&"s", A1)', match: false, apiaryValue: 2 }), 'fail' ],
    [ aCase({ formula: 'CELL( "filename" , A1)', match: false }), 'skip:environment' ],
    [ aCase({ formula: 'INFO("system")', match: false }), 'skip:environment' ],
    [ aCase({ formula: 'MYINFO("system")', match: false, apiaryValue: 2 }), 'fail' ],
    [ aCase({ formula: '\'INFO("system")\'!A1', match: false, apiaryValue: 2 }), 'fail' ],
    [ aCase({ formula: 'SUM(Table1[INFO("system")])', match: false, apiaryValue: 2 }), 'fail' ],
    [ aCase({ formula: '"x"&\'CELL("filename")\'!B2', match: false, apiaryValue: 2 }), 'fail' ],
    [ aCase({ formula: 'CELL("format", A1)' }), 'pass' ],
    [ aCase({ formula: 'WEEKNUM(1,-1)', excelValue: '#NUM!', match: false }), 'skip:excel-unstable' ],
    [ aCase({ formula: 'IF(A1, WEEKNUM(E9, 999999), 0)' }), 'skip:excel-unstable' ],
    [ aCase({ formula: 'ROWS(SEQUENCE(1048576,16384))', excelValue: 0, apiaryValue: '#VALUE!', match: false }), 'skip:excel-abandoned-array' ],
    [ aCase({ formula: 'ROWS(SEQUENCE(1048576,16384))', excelValue: 1048576, apiaryValue: 1048576 }), 'pass' ],
    [ aCase({ formula: 'IFERROR(ROWS(MUNIT(-10001)),0)', excelValue: 0, apiaryValue: 1, match: false }), 'fail' ],
    [ aCase({ formula: 'IFERROR(ROWS(SEQUENCE(-20000,-20000)),0)', excelValue: 0, apiaryValue: 1, match: false }), 'fail' ],
    [ aCase({ formula: 'IF(FALSE,ROWS(SEQUENCE(16384,8191)),0)', excelValue: 0, apiaryValue: 1, match: false }), 'fail' ],
    [ aCase({ formula: 'IF(TRUE,0,ROWS(MUNIT(20000)))', excelValue: 0, apiaryValue: 1, match: false }), 'fail' ],
    [ aCase({ formula: 'IF(TRUE,0,ROWS(MUNIT(20000)))', excelValue: 0, apiaryValue: 0 }), 'pass' ],
    [ aCase({ formula: 'WEEKNUM(1,0)' }), 'pass' ],
    [ aCase({ formula: 'WEEKNUM(1,21)' }), 'pass' ],
    [ aCase({ formula: 'WEEKNUM(1,A2)' }), 'pass' ],
    [ aCase({ functions: [] }), 'pass' ],
  ];

  // Act
  const got = rows.map(([ c ]) => {
    const k = classify(c, catalogue);
    return k.skipReason ? `${k.status}:${k.skipReason}` : k.status;
  });

  // Assert
  assert.deepEqual(got, rows.map(([ , expected ]) => expected));
});

test('classify names the kind of disagreement of a failed case', () => {
  // Arrange
  const rows: Array<[ Case, string | undefined ]> = [
    [ aCase({ match: false, apiaryValue: 2 }), 'wrong-value' ],
    [ aCase({ match: false, excelValue: '#NUM!', apiaryValue: '#VALUE!' }), 'wrong-error' ],
    [ aCase({ match: false, excelValue: 1, apiaryValue: '#VALUE!' }), 'error-for-value' ],
    [ aCase({ match: false, excelValue: '#N/A', apiaryValue: 'sd' }), 'value-for-error' ],
    [ aCase({ match: false, excelValue: '#N/A', apiaryValue: null }), 'value-for-error' ],
    // Text that starts with a #: LEFT(A8,3) over a #DIV/0! cell, and A1&B1 over #N/A.
    [ aCase({ match: false, excelValue: '#DI', apiaryValue: '#DIV/0!' }), 'error-for-value' ],
    [ aCase({ match: false, excelValue: '#N/A', apiaryValue: '#N/A0' }), 'value-for-error' ],
    // Errors an engine spells its own way: LibreOffice's TEXT() overflow, IronCalc's
    // not-implemented answer, and Gnumeric naming an error it has no constant for: an
    // Excel-only code, or a formula it could not parse and replaced with ERROR().
    [ aCase({ match: false, excelValue: '#VALUE!', apiaryValue: '#FMT' }), 'wrong-error' ],
    [ aCase({ match: false, excelValue: '#NUM!', apiaryValue: '#N/IMPL!' }), 'wrong-error' ],
    [ aCase({ match: false, excelValue: '#NUM!', apiaryValue: '#"#CALC!"' }), 'wrong-error' ],
    [ aCase({ match: false, excelValue: 0.5, apiaryValue: '#"CORREL({1,#CALC!,#N/A}, {4,5,6})"' }), 'error-for-value' ],
    [ aCase({}), undefined ],
    [ aCase({ match: false, apiaryValue: '#NAME?' }), undefined ],
  ];

  // Act
  const got = rows.map(([ c ]) => classify(c, catalogue).failKind);

  // Assert
  assert.deepEqual(got, rows.map(([ , expected ]) => expected));
});

test('classify marks parity, never-supported gaps and propagated errors', () => {
  // Arrange
  const gap = aCase({ functions: [ 'CUBEVALUE' ], match: false, apiaryValue: '#NAME?' });
  const propagated = aCase({ match: false, apiaryValue: '#VALUE!', apiaryErrorOrigin: 'S!B7' });
  const own = aCase({ match: false, apiaryValue: '#VALUE!', apiaryErrorOrigin: 'S!A1' });
  const skipped = aCase({ excelDroppedFormula: true });

  // Act & Assert
  assert.deepEqual(classify(gap, catalogue), { status: 'unsupported', parity: true, propagatedError: false, neverSupported: true });
  assert.equal(classify(propagated, catalogue).propagatedError, true);
  assert.equal(classify(own, catalogue).propagatedError, false);
  assert.equal(classify(skipped, catalogue).parity, false);
});

test('classify passes a spilling case only when its spill agrees with Excel\'s', () => {
  // Arrange
  const excel = { excelRange: 'A1:A2', excelValues: [ [ 1 ], [ 2 ] ] };
  const rows: Array<[ Case, string ]> = [
    [ aCase({ spill: { ...excel, apiaryRange: 'A1:A2', apiaryValues: [ [ 1 ], [ 2 ] ], mismatches: 0 } }), 'pass' ],
    [ aCase({ spill: { ...excel, apiaryValues: [ [ 1 ], [ 2 ] ], mismatches: 0 } }), 'pass' ],
    [ aCase({ spill: { ...excel, apiaryRange: 'A1:A2', apiaryValues: [ [ 1 ], [ 3 ] ], mismatches: 1 } }), 'fail:wrong-spilled-value' ],
    [ aCase({ spill: { ...excel, apiaryValues: [ [ 1 ], [ null ] ], mismatches: 1 } }), 'fail:wrong-spilled-value' ],
    [ aCase({ spill: { ...excel, apiaryRange: 'A1:A3', apiaryValues: [ [ 1 ], [ 2 ] ], mismatches: 0 } }), 'fail:wrong-spill-size' ],
    [ aCase({ spill: { ...excel, apiaryRange: 'A1', apiaryValues: [ [ 1 ], [ null ] ], mismatches: 1 } }), 'fail:wrong-spill-size' ],
    [ aCase({ match: false, apiaryValue: '#VALUE!', spill: { ...excel, apiaryRange: 'A1', apiaryValues: [ [ '#VALUE!' ], [ null ] ], mismatches: 1 } }), 'fail:error-for-value' ],
    [ aCase({ match: false, apiaryValue: '#NAME?', spill: { ...excel, apiaryRange: 'A1', apiaryValues: [ [ '#NAME?' ], [ null ] ], mismatches: 1 } }), 'unsupported' ],
    [ aCase({ match: false, apiaryValue: 2, spill: excel }), 'fail:wrong-value' ],
    [ aCase({ loadError: 'timeout', spill: excel }), 'error' ],
  ];

  // Act
  const got = rows.map(([ c ]) => {
    const k = classify(c, catalogue);
    return [ k.status, k.failKind ?? k.skipReason ].filter(Boolean).join(':');
  });

  // Assert
  assert.deepEqual(got, rows.map(([ , expected ]) => expected));
});
