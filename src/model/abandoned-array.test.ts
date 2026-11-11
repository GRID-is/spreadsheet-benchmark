import { test } from 'node:test';
import assert from 'node:assert/strict';
import { buildsAbandonedArray } from './abandoned-array.ts';

test('buildsAbandonedArray finds array functions sized past what Excel computes, as the whole formula or the whole argument of ROWS', () => {
  // Arrange
  const rows: Array<[ string, boolean ]> = [
    [ 'ROWS(SEQUENCE(1048576,16384))', true ],
    [ 'ROWS(MUNIT(1E6))', true ],
    [ 'ROWS(RANDARRAY(1E6,1E6))', true ],
    [ 'ROWS(MAKEARRAY(1E6,1E6,LAMBDA(r,c,1)))', true ],
    [ 'SEQUENCE(16384,8191)', true ],
    [ 'SEQUENCE(134209536)', false ],
    [ 'SEQUENCE(10001,10000)', false ],
    [ 'SEQUENCE(10000,10000)', false ],
    [ 'MUNIT(10000)', false ],
    [ 'RANDARRAY(,1E8)', false ],
    [ 'RANDARRAY(,1E9)', false ],
    [ 'RANDARRAY(1048576,1000)', true ],
    [ 'SEQUENCE(1E10)', false ],
    [ 'ROWS(SEQUENCE(1,134217728))', false ],
    [ 'ROWS(SEQUENCE(1048577,1000))', false ],
    [ 'RANDARRAY()', false ],
    [ 'SEQUENCE(A1,1E9)', false ],
    [ 'MUNIT(B2)', false ],
    [ 'MUNIT(-10001)', false ],
    [ 'SEQUENCE(-20000,-20000)', false ],
    [ 'MAKEARRAY(-20000,-20000,LAMBDA(r,c,1))', false ],
    [ 'RANDARRAY(-1E6,-1E6)', false ],
    [ 'MAKEARRAY(0.5,1E9,LAMBDA(r,c,1))', false ],
    [ ' ROWS( SEQUENCE(16384, 8191) ) ', true ],
    [ 'IF(FALSE,ROWS(SEQUENCE(16384,8191)),0)', false ],
    [ 'IF(TRUE,0,ROWS(MUNIT(20000)))', false ],
    [ 'ROWS(SEQUENCE(16384,8191))+0', false ],
    [ 'SUM(SEQUENCE(16384,8191))', false ],
    [ 'COLUMNS(SEQUENCE(16384,8191))', false ],
    [ 'ROWS(SEQUENCE(16384,8191),1)', false ],
    [ 'ROWS(SEQUENCE(1,1))*ROWS(SEQUENCE(16384,8191))', false ],
  ];

  // Act
  const got = rows.map(([ formula ]) => buildsAbandonedArray(formula));

  // Assert
  assert.deepEqual(got, rows.map(r => r[1]));
});
