import assert from 'node:assert/strict';
import { test } from 'node:test';
import { execFileSync } from 'node:child_process';
import { commandFootprint, runLimited } from './child-process.ts';

const node = process.execPath;

const alive = (pid: number) => {
  try {
    process.kill(pid, 0);
    return true;
  }
  catch {
    return false;
  }
};

const zombie = (pid: number) => {
  try {
    return execFileSync('ps', [ '-o', 'stat=', '-p', String(pid) ], { encoding: 'utf8' }).trim().startsWith('Z');
  }
  catch {
    // ps fails for a process that is gone
    return false;
  }
};

test('runLimited returns the output and status of a command that finishes', async () => {
  const res = await runLimited(node, [ '-e', 'console.log("hi"); console.error("there"); process.exit(3)' ], { timeoutMs: 10000 });
  assert.equal(res.error, null);
  assert.equal(res.status, 3, 'the exit status is passed through');
  assert.equal(res.stdout, 'hi\n', 'stdout is collected');
  assert.equal(res.stderr, 'there\n', 'stderr is collected');
});

test('runLimited kills a command whose physical footprint exceeds the limit', async () => {
  const hog = 'const k = []; setInterval(() => { for (let i = 0; i < 8; i++) k.push(Buffer.alloc(16 << 20, 1)); }, 20);';
  const started = Date.now();
  const res = await runLimited(node, [ '-e', hog ], { timeoutMs: 60000, limitMiB: 200 });
  assert.equal(res.error?.code, 'EMEMLIMIT', String(res.error));
  assert.equal(res.error.message, 'out of memory: exceeded 200 MiB');
  assert.ok(Date.now() - started < 30000, 'killed well before the timeout');
});

test('runLimited kills a command that outlives its time budget, with its descendants', async () => {
  const parent = 'const c = require("node:child_process").spawn("sleep", ["30"]); console.log(c.pid); setInterval(() => {}, 1000);';
  const res = await runLimited(node, [ '-e', parent ], { timeoutMs: 1000 });
  assert.equal(res.error?.code, 'ETIMEDOUT');
  assert.equal(res.error.message, 'timeout after 1000 ms');
  const grandchild = Number(res.stdout.trim());
  // runLimited waits for the command to exit, not for its descendants. A SIGKILL lands
  // asynchronously, and the killed grandchild, orphaned, stays a zombie until init reaps it, which
  // `kill(pid, 0)` still finds.
  const deadline = Date.now() + 5000;
  while (alive(grandchild) && !zombie(grandchild) && Date.now() < deadline) {
    await new Promise(resolve => setTimeout(resolve, 50));
  }
  assert.ok(!alive(grandchild) || zombie(grandchild), 'the grandchild is killed too');
});

test('runLimited reports a command that cannot be started', async () => {
  const res = await runLimited('/nonexistent/engine', [], { timeoutMs: 1000 });
  assert.equal(res.error?.code, 'ENOENT');
});

test('runLimited keeps the command in the harness\'s process group, so a Ctrl-C reaches it', async () => {
  const pgid = (pid: number) => runLimited('ps', [ '-o', 'pgid=', '-p', String(pid) ], { timeoutMs: 10000 }).then(r => r.stdout.trim());
  const res = await runLimited('sh', [ '-c', 'ps -o pgid= -p $$' ], { timeoutMs: 10000 });
  assert.equal(res.stdout.trim(), await pgid(process.pid));
});

test('runLimited bounds each output stream by maxBuffer separately, as spawnSync does', async () => {
  const both = 'process.stdout.write("a".repeat(600)); process.stderr.write("b".repeat(600));';
  const res = await runLimited(node, [ '-e', both ], { timeoutMs: 10000, maxBuffer: 1000 });
  assert.equal(res.error, null, 'streams within maxBuffer each are not cut off');
  assert.equal(res.stdout.length, 600);
  const over = await runLimited(node, [ '-e', 'process.stdout.write("a".repeat(2000))' ], { timeoutMs: 10000, maxBuffer: 1000 });
  assert.equal(over.error?.code, 'ENOBUFS', 'a stream over maxBuffer is an error');
});

test('runLimited kills a command whose memory cannot be read, rather than leave it unlimited', async () => {
  const script = `
    import { runLimited } from ${JSON.stringify(new URL('./child-process.ts', import.meta.url).href)};
    const res = await runLimited(process.execPath, [ '-e', 'setTimeout(() => {}, 30000)' ], { timeoutMs: 60000 });
    console.log(JSON.stringify({ code: res.error?.code, message: res.error?.message }));
  `;
  const started = Date.now();
  const out = execFileSync(node, [ '--input-type=module', '-e', script ], { encoding: 'utf8', env: { ...process.env, FOOTPRINT_PYTHON: '/nonexistent/python3' } });
  const res = JSON.parse(out.trim().split('\n').pop() ?? '');
  assert.equal(res.code, 'EMEMREAD', res.message);
  assert.match(res.message, /^memory not measured: /);
  assert.ok(Date.now() - started < 20000, 'killed at once, not at the end of the command or its time budget');
});

test('runLimited samples through the open command reader rather than start a reader per command', () => {
  // After the command reader is open, FOOTPRINT_PYTHON is set to a missing interpreter, so a reader
  // started for the command would fail with EMEMREAD: the command is measured only if runLimited
  // samples through the open reader.
  const script = `
    import { FootprintReader } from ${JSON.stringify(new URL('../src/memory.ts', import.meta.url).href)};
    import { commandReader, runLimited } from ${JSON.stringify(new URL('./child-process.ts', import.meta.url).href)};
    commandReader.reader = new FootprintReader();
    await commandReader.reader.pressure();
    process.env.FOOTPRINT_PYTHON = '/nonexistent/python3';
    const res = await runLimited(process.execPath, [ '-e', 'setTimeout(() => {}, 300)' ], { timeoutMs: 10000 });
    console.log(JSON.stringify({ code: res.error?.code, peak: res.peakFootprintMiB, open: !commandReader.reader.failed }));
    commandReader.reader.close();
  `;
  const res = JSON.parse(execFileSync(node, [ '--input-type=module', '-e', script ], { encoding: 'utf8' }).trim().split('\n').pop() ?? '');
  assert.equal(res.code, undefined, 'the command is measured');
  assert.ok(res.peak > 0, `the command's peak footprint is read (${res.peak} MiB)`);
  assert.ok(res.open, 'the command reader is left open for the next command');
});

test('runLimited records the highest peak footprint of the commands it has run', async () => {
  commandFootprint.peakMiB = null;
  // About 200 MiB, held long enough to be sampled.
  const big = await runLimited(node, [ '-e', 'const b = Buffer.alloc(200 << 20, 1); setTimeout(() => b, 600);' ], { timeoutMs: 10000 });
  // Long enough for a sample under the load of the whole test run, where 300 ms sometimes got none.
  const small = await runLimited(node, [ '-e', 'setTimeout(() => {}, 1000)' ], { timeoutMs: 10000 });

  assert.ok(big.peakFootprintMiB != null && big.peakFootprintMiB > 200, `peak ${big.peakFootprintMiB} MiB`);
  assert.ok(small.peakFootprintMiB != null && small.peakFootprintMiB < big.peakFootprintMiB, `peaks ${small.peakFootprintMiB} and ${big.peakFootprintMiB} MiB`);
  assert.equal(commandFootprint.peakMiB, big.peakFootprintMiB, 'the larger of the two peaks is recorded');
});
