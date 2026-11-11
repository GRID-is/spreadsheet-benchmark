import { test } from 'node:test';
import assert from 'node:assert/strict';
import { callsUnstableWeeknum } from './unstable-weeknum.ts';

test('callsUnstableWeeknum finds WEEKNUM calls with a numeric constant return type Excel answers unstably', () => {
  // Arrange
  const rows: Array<[ string, boolean ]> = [
    [ 'WEEKNUM(1,-1)', true ],
    [ 'WEEKNUM(E9, -2)', true ],
    [ 'WEEKNUM(datum,999999)', true ],
    [ 'WEEKNUM($K$9,150)', true ],
    [ 'WEEKNUM(1,(- 1))', true ],
    [ 'WEEKNUM(1,-1.0)', true ],
    [ 'IF(A1,WEEKNUM(1,2),WEEKNUM(1,99))', true ],
    [ 'WEEKNUM(46023,34903)', true ],
    [ 'WEEKNUM(1,34903.5)', true ],
    [ 'WEEKNUM(1,-0.5)', true ],
    [ 'WEEKNUM(1,(-.5))', true ],
    [ 'WEEKNUM(1,22)', true ],
    [ 'WEEKNUM(1,32767)', true ],
    [ 'WEEKNUM(1,0)', false ],
    [ 'WEEKNUM(1,0.5)', false ],
    [ 'WEEKNUM(46023,DATE(1995,7,23))', true ],
    [ 'WEEKNUM(46023, date( 95 , 7, 23 ))', true ],
    [ 'WEEKNUM(46023,DATE(1995,6,53))', true ],
    [ 'WEEKNUM(46023,DATE(1995,7,24))', false ],
    [ 'WEEKNUM(46023,DATE(1900,1,1))', false ],
    [ 'WEEKNUM(46023,DATE(1995,7,A1))', false ],
    [ 'WEEKNUM(44561,3)', false ],
    [ 'WEEKNUM(2, 9)', false ],
    [ 'WEEKNUM(1,-3)', false ],
    [ 'WEEKNUM(1,1.5)', false ],
    [ 'WEEKNUM(1,3.2)', false ],
    [ 'WEEKNUM(1,1)+WEEKNUM(1,2)+WEEKNUM(1,11)+WEEKNUM(1,17)+WEEKNUM(1,21)', false ],
    [ 'WEEKNUM(1,+21)', false ],
    [ 'WEEKNUM(1)', false ],
    [ 'WEEKNUM(1,)', false ],
    [ 'WEEKNUM(1,A2)', false ],
    [ 'WEEKNUM(1,1-2)', false ],
    [ 'WEEKNUM(1,"-1")', false ],
    [ 'WEEKNUM(1,FALSE)', false ],
    [ 'ISOWEEKNUM(1,-1)', false ],
  ];

  // Act
  const got = rows.map(([ formula ]) => callsUnstableWeeknum(formula));

  // Assert
  assert.deepEqual(got, rows.map(r => r[1]));
});
