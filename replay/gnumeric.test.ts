import assert from 'node:assert/strict';
import { test } from 'node:test';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

test('gnumeric refuses to start without ssconvert', t => {
  // Arrange
  const dir = mkdtempSync(join(tmpdir(), 'gnumeric-'));
  t.after(() => rmSync(dir, { recursive: true, force: true }));

  // Act
  const res = spawnSync(process.execPath, [ 'replay/gnumeric.ts', '--source', dir ], {
    env: { ...process.env, SSCONVERT: join(dir, 'ssconvert') }, encoding: 'utf8',
  });

  // Assert
  assert.notEqual(res.status, 0, 'the runner fails');
  assert.match(res.stderr, /ssconvert not found/, 'the runner says that ssconvert is missing');
});
