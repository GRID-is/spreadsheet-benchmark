import { test } from 'node:test';
import assert from 'node:assert/strict';
import { blankNames, callArguments, storedFormulaDiffers, userFunctionNames } from './formula-text.ts';

test('storedFormulaDiffers ignores file-form conventions and keeps real misreadings', () => {
  // Arrange: stored text from corpus files, and Apiary's reading of each
  const rows: Array<[ string, string, boolean ]> = [
    [ '_xlfn._xlws.FILTER(A1:D16, A1:D1<>"Apple")', 'FILTER(A1:D16, A1:D1<>"Apple")', false ],
    [ '_xlfn.FORMULATEXT(_xlfn.SINGLE(A:A))', 'FORMULATEXT(@A:A)', false ],
    [ '_xlfn.SINGLE("abc")', '@"abc"', false ],
    [ 'SUM(_xlfn.ANCHORARRAY(C2))', 'SUM(C2#)', false ],
    [ 'DAVERAGE(_xlnm.database,"Weight",_xlnm.criteria)', 'DAVERAGE(database,"Weight",criteria)', false ],
    [ 'SUMPRODUCT(C9:C9,3,B10,C9:C9)', 'SUMPRODUCT(C9,3,B10,C9)', false ],
    [ 'QUERY(A1:H7, "group\r\nby G")', 'QUERY(A1:H7, "group\nby G")', false ],
    [ 'IFERROR(__xludf.DUMMYFUNCTION("SUM(A1:A1)"), 0)', 'IFERROR(__xludf.DUMMYFUNCTION("SUM(A1:A1)"), 0)', false ],
    [ '_xlfn.SINGLE(-1.2)', '@-1.2', false ],
    [ 'ROWS(_xlfn.SINGLE(OFFSET(B1, 0, 0, 4, 1)))', 'ROWS(@OFFSET(B1, 0, 0, 4, 1))', false ],
    [ 'SUM(_xlfn.ANCHORARRAY(OFFSET(A1,1,1)))', 'SUM(OFFSET(A1,1,1)#)', false ],
    [ 'LOG10!LOG10*2', "'LOG10'!LOG10*2", false ],
    [ "'SINGLE(A1)'!B2", "'SINGLE(A1)'!B2", false ],
    [ "'_xlfn.foo'!A1+'A1:A1'!B2", "'_xlfn.foo'!A1+'A1:A1'!B2", false ],
    [ "SUM('It''s SINGLE(1)'!A1:A1)", "SUM('It''s SINGLE(1)'!A1)", false ],
    [ "_xlfn.SINGLE('Q(1'!A1)", "@'Q(1'!A1", false ],
    [ 'IF(A1="it\'s",_xlfn.SINGLE(B1),0)', 'IF(A1="it\'s",@B1,0)', false ],
    [ "SUM(Table1[Item '#])+_xlfn.CONCAT(A1)", "SUM(Table1[Item '#])+CONCAT(A1)", false ],
    [ "T[[#This Row],[Col'[1']]]+_xlfn.SINGLE(A1)", "T[[#This Row],[Col'[1']]]+@A1", false ],
    [ '_xlfn.SINGLE(_xlfn.SEQUENCE(2,2) * 42)', '@SEQUENCE(2,2) * 42', true ],
    [ 'SUM(Sheet2:Sheet5!A11)', "SUM(Sheet2:'Sheet5'!A11)", true ],
    [ '_xlfn._xlws.FILTER([1]NoSuchSheet!A1, {TRUE})', 'FILTER([NoSuchSheet]NoSuchSheet!A1, {TRUE})', true ],
    [ '_xlfn.LET(_xlpm.f,PLUS2, _xlpm.n,1, _xlpm.f(_xlpm.n))', 'LET(f,PLUS2, n,1, _xlpm.f(n))', true ],
    [ '_xludf.MINIFS($C$6:$C$12,$B$6:$B$12,">=2")', 'MINIFS($C$6:$C$12,$B$6:$B$12,">=2")', true ],
  ];

  // Act
  const got = rows.map(([ stored, apiary ]) => storedFormulaDiffers(stored, apiary));

  // Assert
  assert.deepEqual(got, rows.map(r => r[2]));
});

test('userFunctionNames finds _xludf. calls outside string literals, quoted sheet names and structured references only', () => {
  // Arrange
  const rows: Array<[ string, string[] ]> = [
    [ '_xludf.MINIFS($C$6:$C$12,$B$6:$B$12,">=2")', [ 'MINIFS' ] ],
    [ 'IF(_xludf.maxifs(A1:A2,B1:B2,1)>0,1,0)', [ 'MAXIFS' ] ],
    [ 'IFERROR(__xludf.DUMMYFUNCTION("MATCH(3, A1:A2)"), 0)', [] ],
    [ '"_xludf.MINIFS(1)"', [] ],
    [ "'_xludf.MINIFS(SINGLE(A1))'!B2", [] ],
    [ "'It''s _xludf.X(1)'!A1+_xludf.MINIFS(1)", [ 'MINIFS' ] ],
    [ "Table1[Item '#]+_xludf.MINIFS(A1:A2,B1:B2,1)", [ 'MINIFS' ] ],
    [ 'T[_xludf.X(1)]', [] ],
  ];

  // Act
  const got = rows.map(([ stored ]) => userFunctionNames(stored));

  // Assert
  assert.deepEqual(got, rows.map(r => r[1]));
});

test('callArguments splits each call of a function at its own commas, nested calls included', () => {
  // Arrange
  const rows: Array<[ string, string[][] ]> = [
    [ 'WEEKNUM(1,-1)', [ [ '1', '-1' ] ] ],
    [ 'weeknum(DATE(2013, 1, 1), 2)', [ [ 'DATE(2013, 1, 1)', ' 2' ] ] ],
    [ 'WEEKNUM(SUM({1,2}),"a,b")+WEEKNUM(1)', [ [ 'SUM({1,2})', '"a,b"' ], [ '1' ] ] ],
    [ 'WEEKNUM(WEEKNUM(1,2),3)', [ [ 'WEEKNUM(1,2)', '3' ], [ '1', '2' ] ] ],
    [ 'WEEKNUM(1,)', [ [ '1', '' ] ] ],
    [ '"Straße"&WEEKNUM(1,-1)', [ [ '1', '-1' ] ] ],
    [ 'ISOWEEKNUM(1)+_xlfn.WEEKNUM(1,2)+\'WEEKNUM(1,2)\'!A1+"WEEKNUM(1,2)"', [] ],
    [ 'WEEKNUM(1,2', [] ],
  ];

  // Act
  const got = rows.map(([ formula ]) => callArguments(formula, 'WEEKNUM'));

  // Assert
  assert.deepEqual(got, rows.map(r => r[1]));
});

test('blankNames blanks quoted sheet names and bracketed references, not string literals', () => {
  // Arrange
  const rows: Array<[ string, string ]> = [
    [ '\'INFO("system")\'!A1', '                !A1' ],
    [ 'SUM(Table1[INFO("system")])', 'SUM(Table1                )' ],
    [ '"x"&\'It\'\'s\'!B2+INFO("system")', '"x"&       !B2+INFO("system")' ],
    [ "T[[#This Row],[Col'[1']]]&\"[a]\"", "T                        &\"[a]\"" ],
  ];

  // Act
  const got = rows.map(([ formula ]) => blankNames(formula));

  // Assert
  assert.deepEqual(got, rows.map(r => r[1]));
});
