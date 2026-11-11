import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { hostname } from 'node:os';
import { test } from 'node:test';
import { executionEnvironment } from './execution-environment.ts';

test('records the explicit machine identity even in an engine container', () => {
  process.env.ENGINE_COMPATIBILITY_ENGINE = 'apiary';
  process.env.ENGINE_COMPATIBILITY_MACHINE = 'remote-host';
  assert.equal(executionEnvironment().machine, 'remote-host');
});

test('uses the hostname on a native host', { skip: existsSync('/.dockerenv') || existsSync('/run/.containerenv') }, () => {
  delete process.env.ENGINE_COMPATIBILITY_MACHINE;
  delete process.env.ENGINE_COMPATIBILITY_ENGINE;
  assert.equal(executionEnvironment().machine, hostname());
});

test('does not identify an engine container by its hostname', () => {
  delete process.env.ENGINE_COMPATIBILITY_MACHINE;
  process.env.ENGINE_COMPATIBILITY_ENGINE = 'apiary';
  assert.equal(executionEnvironment().machine, undefined);
});
