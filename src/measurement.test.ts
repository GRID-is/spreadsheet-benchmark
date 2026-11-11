import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CURRENT_MEASUREMENT, measurementDifferences } from './measurement.ts';

test('measurementDifferences names each field that differs from the current measurement', () => {
  assert.deepEqual(measurementDifferences(CURRENT_MEASUREMENT), [], 'the current measurement has no differences');
  assert.deepEqual(measurementDifferences(undefined), [ 'no measurement recorded' ], 'a report from before measurements were recorded');
  assert.deepEqual(
    measurementDifferences({ ...CURRENT_MEASUREMENT, timeoutMs: 1000 }),
    [ `timeoutMs 1000, current ${CURRENT_MEASUREMENT.timeoutMs}` ],
    'a changed timeout',
  );
});
