import assert from 'node:assert/strict';
import { test } from 'node:test';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

test('gridcraft refuses to start without gridcraft-cli', t => {
  // Arrange
  const dir = mkdtempSync(join(tmpdir(), 'gridcraft-'));
  t.after(() => rmSync(dir, { recursive: true, force: true }));

  // Act
  const res = spawnSync(process.execPath, [ 'replay/gridcraft.ts', '--source', dir ], {
    env: { ...process.env, GRIDCRAFT_CLI: join(dir, 'gridcraft-cli') }, encoding: 'utf8',
  });

  // Assert
  assert.notEqual(res.status, 0, 'the runner fails');
  assert.match(res.stderr, /gridcraft-cli not found/, 'the runner says that gridcraft-cli is missing');
});
