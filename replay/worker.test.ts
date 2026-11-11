import assert from 'node:assert/strict';
import { test } from 'node:test';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { forkJobProcess, JobWorker } from '../src/job-worker.ts';
import { type Case } from '../src/model/case.ts';
import { type WorkerData, type WorkerRequest, type WorkerResult } from './common.ts';

test('the job process of an engine evaluated outside it opens the command reader before the fixture', async t => {
  // The engine points FOOTPRINT_PYTHON at a missing interpreter before it runs its command, so a
  // reader started for the command would fail with EMEMREAD: the command is measured only if the job
  // process opened the reader before the fixture.
  const dir = mkdtempSync(join(tmpdir(), 'worker-test-'));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  const engineModule = join(dir, 'engine.mjs');
  writeFileSync(engineModule, `
    import { runLimited } from ${JSON.stringify(new URL('./child-process.ts', import.meta.url).href)};
    export const engine = {
      name: 'command', version: '0', warmUp: false, inProcess: false,
      async buildFile () {
        process.env.FOOTPRINT_PYTHON = '/nonexistent/python3';
        const res = await runLimited(process.execPath, [ '-e', 'setTimeout(() => {}, 300)' ], { timeoutMs: 10000 });
        if (res.error) throw res.error;
        return () => res.peakFootprintMiB;
      },
    };
  `);
  const runner = new JobWorker<WorkerRequest, WorkerResult>(() => forkJobProcess(new URL('./worker.ts', import.meta.url), { engineModule: pathToFileURL(engineModule).href, versionTag: '0' } satisfies WorkerData), 20000);
  const cell = { suite: 's', category: 'c', file: 'f.xlsx', sheet: 'Sheet1', ref: 'A1', id: 'f.xlsx#Sheet1#A1', formula: '1', functions: [], excelValue: 1, apiaryValue: null, match: false, eps: 0, excelDroppedFormula: false, oracleApp: 'Excel', apiaryVersion: '' } satisfies Case;

  const outcome = await runner.run({ fixture: join(dir, 'f.xlsx'), cases: [ cell ] });
  await runner.stop();

  assert.ok('result' in outcome && 'rows' in outcome.result, `the command is measured: ${JSON.stringify(outcome)}`);
  const peak = outcome.result.rows[0].apiaryValue;
  assert.ok(typeof peak === 'number' && peak > 0, `the command's peak footprint is read (${peak} MiB)`);
});

test('the job process reads and compares the spilled cells of Excel\'s spill range', async t => {
  const dir = mkdtempSync(join(tmpdir(), 'worker-spill-'));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  // The engine answers from a fixed table and records the cells it was asked for.
  const engineModule = (name: string) => {
    const path = join(dir, `${name}.mjs`);
    writeFileSync(path, `
      const values = { A1: 1, A2: 2, A3: 7, C1: 5, C2: 6 };
      export const engine = {
        name: ${JSON.stringify(name)}, version: '0', warmUp: false,
        buildFile (fixture, cells) {
          const asked = cells.map(c => c.ref).join(',');
          return Object.assign(c => c.ref === 'Z1' ? asked : values[c.ref] ?? null, { spillRange: c => ({ A1: 'A1:A3', C1: 'C1:C2' })[c.ref] });
        },
      };
    `);
    return pathToFileURL(path).href;
  };
  const base = { suite: 's', category: 'c', file: 'f.xlsx', sheet: 'Sheet1', formula: 'SEQUENCE(3)', functions: [ 'SEQUENCE' ], apiaryValue: null, match: false, eps: 5e-8, excelDroppedFormula: false, oracleApp: 'Excel', apiaryVersion: '' };
  const cases: Case[] = [
    // Excel spilled into A1:A3; the engine's A3 differs.
    { ...base, ref: 'A1', id: 'a', excelValue: 1, spill: { excelRange: 'A1:A3', excelValues: [ [ 1 ], [ 2 ], [ 3 ] ], apiaryRange: 'A1:A3', apiaryValues: [ [ 1 ], [ 2 ], [ 3 ] ], mismatches: 0 } },
    // Excel did not spill; the engine did.
    { ...base, ref: 'C1', id: 'c', excelValue: 5 },
    // The cells the engine was asked for.
    { ...base, ref: 'Z1', id: 'z', excelValue: 'A1,C1,Z1,A2,A3' },
  ];
  const rowsOf = async (module: string) => {
    const runner = new JobWorker<WorkerRequest, WorkerResult>(() => forkJobProcess(new URL('./worker.ts', import.meta.url), { engineModule: module, versionTag: 'e@0' } satisfies WorkerData), 20000);
    t.after(() => runner.stop());
    const outcome = await runner.run({ fixture: join(dir, 'f.xlsx'), cases });
    assert.ok('result' in outcome && 'rows' in outcome.result, JSON.stringify(outcome));
    return outcome.result.rows;
  };

  const [ a, c, z ] = await rowsOf(engineModule('spilling'));
  assert.deepEqual(a.spill, { excelRange: 'A1:A3', excelValues: [ [ 1 ], [ 2 ], [ 3 ] ], apiaryRange: 'A1:A3', apiaryValues: [ [ 1 ], [ 2 ], [ 7 ] ], mismatches: 1 }, 'the engine\'s spilled values replace the source engine\'s');
  assert.deepEqual(c.spill, { excelRange: 'C1', excelValues: [ [ 5 ] ], apiaryRange: 'C1:C2', apiaryValues: [ [ 5 ] ], mismatches: 0 }, 'a spill of the engine alone');
  assert.equal(z.apiaryValue, 'A1,C1,Z1,A2,A3', 'the engine is asked for the cases, then the spilled cells');
});

test('the job process records the phase times an engine reports with its values', async t => {
  const dir = mkdtempSync(join(tmpdir(), 'worker-phases-'));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  const engineModule = join(dir, 'engine.mjs');
  writeFileSync(engineModule, `
    export const engine = {
      name: 'timed', version: '0', warmUp: false,
      buildFile: () => Object.assign(() => 1, { phasesMs: { readMs: 1.5, recalcMs: 2.25 } }),
    };
  `);
  const runner = new JobWorker<WorkerRequest, WorkerResult>(() => forkJobProcess(new URL('./worker.ts', import.meta.url), { engineModule: pathToFileURL(engineModule).href, versionTag: 'e@0' } satisfies WorkerData), 20000);
  t.after(() => runner.stop());
  const cell = { suite: 's', category: 'c', file: 'f.xlsx', sheet: 'Sheet1', ref: 'A1', id: 'f.xlsx#Sheet1#A1', formula: '1', functions: [], excelValue: 1, apiaryValue: null, match: false, eps: 0, excelDroppedFormula: false, oracleApp: 'Excel', apiaryVersion: '' } satisfies Case;

  const outcome = await runner.run({ fixture: join(dir, 'f.xlsx'), cases: [ cell ] });

  assert.ok('result' in outcome && 'timing' in outcome.result, JSON.stringify(outcome));
  assert.deepEqual(outcome.result.timing.enginePhasesMs, { readMs: 1.5, recalcMs: 2.25 }, 'the engine\'s own phase times are in the file\'s timing');
});
