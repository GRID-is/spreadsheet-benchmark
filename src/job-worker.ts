import { fork, type ChildProcess, type Serializable } from 'node:child_process';
import { GCProfiler, getHeapStatistics, setFlagsFromString } from 'node:v8';
import { once } from 'node:events';
import { flushCompileCache } from 'node:module';
import { fileURLToPath } from 'node:url';
import { runInNewContext } from 'node:vm';
import { FootprintReader, FootprintWatch, HEAP_LIMIT_MIB, MEMORY_LIMIT_MIB, treeFootprintMiB, treePids } from './memory.ts';
import { PressureWatch, type MemoryPressure } from './memory-pressure.ts';

/**
 * The physical footprint (CONTEXT.md) of a job's process tree: after the warm-up, before the job
 * was sent, and the highest sampled while it ran. Peak minus baseline is the job's own memory.
 * Either is null where it could not be read.
 */
export type JobMemory = { baselineFootprintMiB: number | null, peakFootprintMiB: number | null };

/**
 * A job process's own account of its memory when it answered, in MiB, from `process.memoryUsage()`
 * and V8's heap statistics: the V8 heap committed and in use, memory held outside it (`external`,
 * which includes `arrayBuffers` and wasm memory), and the most `malloc` V8 has held. Wasm memory
 * never shrinks, so its size when the job answers is its peak.
 */
export type ProcessMemory = { heapTotalMiB: number, heapUsedMiB: number, externalMiB: number, arrayBuffersMiB: number, peakMallocedMiB: number };

/**
 * The retained heap and peak live heap (docs/measurements.md) of a garbage-collected engine, in MiB: the
 * heap in use after a full collection while the evaluated workbook is still reachable, with the
 * memory then held outside the heap (wasm memory and buffers in V8), and the highest heap in use
 * after any full collection of the job up to that measurement. Unlike the physical footprint,
 * neither counts garbage that the collector has yet to reclaim. A runtime that does not account for
 * its memory outside the heap (.NET) has no `retainedExternalMiB`.
 */
export type HeapMemory = { retainedHeapMiB: number, retainedExternalMiB?: number, peakLiveHeapMiB: number };

/**
 * What a job's handler is given to measure its engine's heap. `measure` collects garbage in the job
 * process and reads its heap; `keep` holds the evaluated workbook's model reachable meanwhile, and
 * the call ends the window of the peak live heap. An engine evaluated in another runtime reports
 * that runtime's figures with `report` instead. A job that calls neither has no heap figures.
 */
export type HeapProbe = { measure (...keep: unknown[]): void, report (heap: HeapMemory): void };

/**
 * A job's answer or failure, with its memory, and the memory pressure while it ran, if it was
 * sent.
 */
export type JobOutcome<R> = ({ result: R, processMemory: ProcessMemory, heap: HeapMemory | null } | { failure: string }) & JobMemory & { memoryPressure?: MemoryPressure };

/**
 * Formats a job's memory for a progress line, e.g.
 * `peak 412 MiB, baseline 96 MiB, heap 300/320 MiB, external 40 MiB, peak live heap 180 MiB,
 * retained 120 MiB heap + 14 MiB external`: after the footprint, the process's own account where
 * it answered, and the engine's heap figures where measured.
 */
export function formatJobMemory ({ baselineFootprintMiB, peakFootprintMiB, processMemory, heap }: JobMemory & { processMemory?: ProcessMemory, heap?: HeapMemory | null }): string {
  const mib = (reading: number | null) => reading == null ? 'unread' : `${Math.round(reading)} MiB`;
  const own = processMemory ? `, heap ${processMemory.heapUsedMiB}/${processMemory.heapTotalMiB} MiB, external ${processMemory.externalMiB} MiB` : '';
  const external = heap?.retainedExternalMiB == null ? '' : ` + ${heap.retainedExternalMiB} MiB external`;
  const live = heap ? `, peak live heap ${heap.peakLiveHeapMiB} MiB, retained ${heap.retainedHeapMiB} MiB heap${external}` : '';
  return `peak ${mib(peakFootprintMiB)}, baseline ${mib(baselineFootprintMiB)}${own}${live}`;
}

/** Messages from a job process to its parent. */
type ChildMessage<R> = { ready: true } | { result: R, processMemory: ProcessMemory, heap: HeapMemory | null } | { crashed: string };

function processMemory (): ProcessMemory {
  const usage = process.memoryUsage();
  return {
    heapTotalMiB: mib(usage.heapTotal),
    heapUsedMiB: mib(usage.heapUsed),
    externalMiB: mib(usage.external),
    arrayBuffersMiB: mib(usage.arrayBuffers),
    peakMallocedMiB: mib(getHeapStatistics().peak_malloced_memory),
  };
}

const mib = (bytes: number) => Math.round(bytes / 2 ** 20);

let exposedGc: NodeJS.GCFunction | undefined;

/**
 * Returns V8's `gc`, exposed on first use rather than by starting the process with `--expose-gc`.
 * Under that flag, V8 rejects the compiled code that Node caches for its built-in modules, which
 * then compile afresh: about 25 ms of a job process's start. Exposing it at run time has the same
 * effect on built-in modules loaded afterwards, so it is first used at the heap measurement, once
 * the engine has evaluated the fixture.
 */
function collectGarbage (): NodeJS.GCFunction {
  if (globalThis.gc != null) {
    return globalThis.gc;
  }
  if (exposedGc == null) {
    setFlagsFromString('--expose-gc');
    exposedGc = runInNewContext('gc') as NodeJS.GCFunction;
  }
  return exposedGc;
}

/**
 * Makes a job's `HeapProbe`, and `finish`, which returns the job's heap figures: those reported,
 * else those measured, else null. It profiles the job's collections from its creation until
 * `measure`, or `finish` if `measure` was not called.
 */
function heapProbe (): { probe: HeapProbe, finish: () => HeapMemory | null } {
  const profiler = new GCProfiler();
  profiler.start();
  let profiling = true;
  const stop = () => {
    if (!profiling) return [];
    profiling = false;
    return profiler.stop().statistics;
  };
  let measured: HeapMemory | null = null;
  let reported: HeapMemory | null = null;
  const probe: HeapProbe = {
    measure (...keep) {
      // A second collection frees a little more: about 0.1 MiB after Apiary evaluates a fixture.
      for (let i = 0; i < 2; i++) {
        collectGarbage()({ type: 'major', execution: 'sync' });
      }
      const usage = process.memoryUsage();
      // The forced collections are among these, so a job too short for any of its own has one.
      const afterFullCollections = stop().filter(gc => gc.gcType === 'MarkSweepCompact').map(gc => gc.afterGC.heapStatistics.usedHeapSize);
      measured = {
        retainedHeapMiB: mib(usage.heapUsed),
        retainedExternalMiB: mib(usage.external),
        peakLiveHeapMiB: mib(Math.max(usage.heapUsed, ...afterFullCollections)),
      };
      // Used after the reading, so that what it holds stays reachable through it.
      keep.length = 0;
    },
    report (heap) {
      reported = heap;
    },
  };
  return { probe, finish: () => {
    stop();
    return reported ?? measured;
  } };
}

/** The fixture a job process evaluates, unmeasured, before it reports ready: one sheet, A3 = A1+A2. */
export const WARM_UP_FIXTURE = fileURLToPath(new URL('../replay/recalc-canary.xlsx', import.meta.url));

const SAMPLE_INTERVAL_MS = 200;

/**
 * The compile cache of job processes, in the repository's ignored `.runs/`. Not Node's default of
 * `node-compile-cache` in the temporary directory, which another user may own: GitHub's Ubuntu
 * runner images ship one that their CI user cannot write, and Node then runs without the cache.
 */
const COMPILE_CACHE_DIR = fileURLToPath(new URL('../.runs/compile-cache', import.meta.url));

const messageOf = (err: unknown) => err instanceof Error ? err.message : String(err);

/**
 * Forks a job process running `module`, which serves jobs with `serveJobs` and reads `data` with
 * `jobData`. V8's heap limit is `HEAP_LIMIT_MIB`, below the memory limit on the process tree's
 * footprint, so that V8 collects garbage before the footprint reaches the memory limit. The process
 * reads its modules' compiled code from Node's compile cache, as each process would otherwise
 * compile them afresh.
 */
export function forkJobProcess (module: URL, data: unknown, execArgv: string[] = []): ChildProcess {
  return fork(fileURLToPath(module), [ JSON.stringify(data) ], {
    serialization: 'advanced',
    // stderr is piped for JobWorker, which passes it through and reads V8's report of an exhausted
    // heap from it.
    stdio: [ 'inherit', 'inherit', 'pipe', 'ipc' ],
    env: { ...process.env, NODE_COMPILE_CACHE: process.env.NODE_COMPILE_CACHE ?? COMPILE_CACHE_DIR },
    // Apiary's wasm import would print an ExperimentalWarning in every process. V8's default of 10
    // stack frames cuts off the traces that engines log.
    execArgv: [ ...process.execArgv, '--no-warnings=ExperimentalWarning', '--stack-trace-limit=50', `--max-old-space-size=${HEAP_LIMIT_MIB}`, ...execArgv ],
  });
}

/** In a job process: the data its parent passed to `forkJobProcess`. */
export function jobData<T> (): T {
  return JSON.parse(process.argv[2]);
}

/**
 * In a job process: runs `warmUp`, then reports ready and answers each job with `handle`, which
 * may measure its engine's heap with the probe it is given. An uncaught error is reported to the
 * parent as a crash, with its message, and ends the process.
 */
export async function serveJobs<J extends Serializable, R extends Serializable> (handle: (job: J, heap: HeapProbe) => Promise<R>, warmUp?: () => Promise<unknown>): Promise<void> {
  const send = (message: ChildMessage<R>, then?: () => void) => {
    if (!process.send) {
      throw new Error('a job process must be forked with an IPC channel');
    }
    process.send(message, () => then?.());
  };
  const crash = (err: unknown) => send({ crashed: messageOf(err) }, () => process.exit(1));
  process.on('uncaughtException', crash);
  process.on('unhandledRejection', crash);
  // An orphaned job process, whose parent has died, ends itself.
  process.on('disconnect', () => process.exit(0));
  try {
    await warmUp?.();
  }
  catch (err) {
    crash(new Error(`warm-up failed: ${messageOf(err)}`));
    return;
  }
  // The process is killed after its job, which would leave the cache unwritten.
  flushCompileCache();
  process.on('message', async (job: J) => {
    const { probe, finish } = heapProbe();
    const result = await handle(job, probe);
    send({ result, processMemory: processMemory(), heap: finish() });
  });
  send({ ready: true });
}

async function killTree (child: ChildProcess, reader: FootprintReader): Promise<void> {
  if (child.exitCode != null || child.signalCode != null) {
    return;
  }
  const exited = once(child, 'exit');
  // An `error` while the tree's pids are read rejects `exited` before it is awaited.
  exited.catch(() => {});
  const pid = child.pid;
  const pids = pid == null ? [] : await treePids(pid, reader).catch(() => [ pid ]);
  for (const p of pids) {
    try {
      process.kill(p, 'SIGKILL');
    }
    catch {
      // already gone
    }
  }
  if (pid != null) {
    await exited;
  }
}

function exitFailure (code: number | null, signal: NodeJS.Signals | null): string {
  return signal ? `worker exited with signal ${signal}` : `worker exited with code ${code}`;
}

/** The line V8 writes to stderr before it aborts on exhausting its heap. */
const HEAP_EXHAUSTED = /^FATAL ERROR: .*JavaScript heap out of memory/;
/** How long a dead process's stderr may stay open, e.g. held by a descendant, before it is read. */
const STDERR_CLOSE_MS = 1000;
/**
 * How long a process whose job could not be sent may take to exit, which names the cause, before
 * the job fails with the send error.
 */
const SEND_FAILED_EXIT_MS = 1000;

/**
 * Passes a job process's stderr, if piped, through to the harness's, and returns a function that
 * resolves, once the stream has closed or STDERR_CLOSE_MS has passed, whether the process reported
 * exhausting V8's heap there. It keeps only the current line, up to 1 KiB, not a tail, which the
 * native stack trace after V8's report would overrun.
 */
function watchHeapExhaustion (child: ChildProcess): () => Promise<boolean> {
  const stderr = child.stderr;
  if (stderr == null) {
    return async () => false;
  }
  let exhausted = false;
  let line = '';
  stderr.on('data', (chunk: Buffer) => {
    process.stderr.write(chunk);
    const lines = (line + chunk.toString('utf8')).split('\n');
    line = (lines.pop() as string).slice(0, 1024);
    exhausted ||= lines.some(l => HEAP_EXHAUSTED.test(l));
  });
  const closed = once(stderr, 'close');
  return async () => {
    let timer: NodeJS.Timeout | undefined;
    await Promise.race([ closed, new Promise(resolve => {
      timer = setTimeout(resolve, STDERR_CLOSE_MS);
    }) ]);
    clearTimeout(timer);
    return exhausted || HEAP_EXHAUSTED.test(line);
  };
}

/**
 * The heap limit, in MiB, of a job process: the last `--max-old-space-size` it was given, or null
 * if none.
 */
function heapLimitOf (child: ChildProcess): string | null {
  const flags = child.spawnargs.map(arg => /^--max-old-space-size=(\d+)$/.exec(arg)?.[1]).filter(mib => mib != null);
  return flags.at(-1) ?? null;
}

/**
 * Watches a job process's stderr and returns the failure of the process if it exits before it
 * answers. V8 aborts on exhausting its heap, so that cause is read from stderr, not from the signal.
 */
function exitCause (child: ChildProcess): (code: number | null, signal: NodeJS.Signals | null) => Promise<string> {
  const heapExhausted = watchHeapExhaustion(child);
  return async (code, signal) => {
    if (!await heapExhausted()) {
      return exitFailure(code, signal);
    }
    const limit = heapLimitOf(child);
    return `out of memory: exceeded ${limit == null ? 'the V8 heap limit' : `the ${limit} MiB heap limit`}`;
  };
}

/** Kills a job's process tree, then stops sampling its footprint. */
async function endProcess (child: ChildProcess, reader: FootprintReader, watch: FootprintWatch): Promise<void> {
  try {
    await killTree(child, reader);
  }
  finally {
    await watch.stop().catch(() => {
      // a failed reading has already settled the job
    });
  }
}

/**
 * How long a job process may take to start and run its warm-up before it reports ready. It is
 * separate from the job timeout, which a run may set far shorter than an engine takes to load.
 */
export const START_TIMEOUT_MS = 60000;

type Settled<R> = { result: R, processMemory: ProcessMemory, heap: HeapMemory | null } | { failure: string };

/**
 * Runs each job in a fresh process, so that no job inherits another's memory or state, with a
 * per-job timeout and the memory limit on the process tree's physical footprint. The process
 * evaluates a warm-up fixture before it reports ready; its footprint then is the job's baseline.
 * A job settles at once with a failure naming the cause when its process crashes (including a
 * failed warm-up), exits, goes over the memory limit or has its memory fail to read, before or
 * after it is ready, or when it is not ready within the start-up timeout or does not answer within
 * the job timeout. Once a job has settled, its process tree is killed in the background; the next
 * job's process, and `stop`, wait until that process has exited.
 */
export class JobWorker<J extends Serializable, R extends Serializable> {
  private readonly spawn: () => ChildProcess;
  private readonly timeoutMs: number;
  private readonly limitMiB: number;
  private readonly startTimeoutMs: number;
  private reader: FootprintReader | null = null;
  /**
   * The pending ends of answered jobs, each of which kills the job's process tree. The next job and
   * `stop` wait for them.
   */
  private readonly ending = new Set<Promise<void>>();

  constructor (spawn: () => ChildProcess, timeoutMs: number, { limitMiB = MEMORY_LIMIT_MIB, startTimeoutMs = START_TIMEOUT_MS }: { limitMiB?: number, startTimeoutMs?: number } = {}) {
    this.spawn = spawn;
    this.timeoutMs = timeoutMs;
    this.limitMiB = limitMiB;
    this.startTimeoutMs = startTimeoutMs;
  }

  /** Resolves to null once the process reports ready, or to the failure that came first. */
  private ready (child: ChildProcess, aborted: Promise<{ failure: string }>, died: (code: number | null, signal: NodeJS.Signals | null) => Promise<string>): Promise<string | null> {
    return new Promise(resolve => {
      let settled = false;
      const settle = (failure: string | null | Promise<string>) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        child.off('message', onMessage);
        child.off('error', onError);
        child.off('exit', onExit);
        resolve(failure);
      };
      const onMessage = (message: ChildMessage<R>) => settle('crashed' in message ? `worker crashed: ${message.crashed}` : null);
      const onError = (err: Error) => settle(`worker crashed: ${err.message}`);
      const onExit = (code: number | null, signal: NodeJS.Signals | null) => settle(died(code, signal));
      const timer = setTimeout(() => settle(`not ready after ${this.startTimeoutMs} ms`), this.startTimeoutMs);
      child.on('message', onMessage);
      child.on('error', onError);
      child.on('exit', onExit);
      aborted.then(({ failure }) => settle(failure));
    });
  }

  /**
   * Sends `job` and resolves to the process's answer, or to the failure that came first. A job that
   * could not be sent fails with the send error if the process has not exited within
   * SEND_FAILED_EXIT_MS.
   */
  private answer (child: ChildProcess, job: J, aborted: Promise<{ failure: string }>, died: (code: number | null, signal: NodeJS.Signals | null) => Promise<string>): Promise<Settled<R>> {
    return new Promise<Settled<R>>(resolve => {
      let settled = false;
      let sendFailedTimer: NodeJS.Timeout | undefined;
      const finish = (outcome: Settled<R> | Promise<Settled<R>>) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        clearTimeout(sendFailedTimer);
        child.off('message', onMessage);
        child.off('error', onError);
        child.off('exit', onExit);
        resolve(outcome);
      };
      const onMessage = (message: ChildMessage<R>) => {
        if ('result' in message) finish({ result: message.result, processMemory: message.processMemory, heap: message.heap });
        else if ('crashed' in message) finish({ failure: `worker crashed: ${message.crashed}` });
      };
      const onError = (err: Error) => finish({ failure: `worker crashed: ${err.message}` });
      const onExit = (code: number | null, signal: NodeJS.Signals | null) => finish(died(code, signal).then(failure => ({ failure })));
      const timer = setTimeout(() => finish({ failure: `timeout after ${this.timeoutMs} ms` }), this.timeoutMs);
      child.on('message', onMessage);
      child.on('error', onError);
      child.on('exit', onExit);
      aborted.then(finish);
      // Without a callback, a failed send emits `error`, which no listener catches once the job has
      // settled. A process that closed its channel is usually dying, and onExit reports its exit.
      child.send(job, err => {
        if (err != null) {
          sendFailedTimer = setTimeout(() => finish({ failure: `job not sent: ${err.message}` }), SEND_FAILED_EXIT_MS);
        }
      });
    });
  }

  /** The footprint reader, replaced first if its helper has died, as it would fail every later job. */
  private liveReader (): FootprintReader {
    if (this.reader?.failed) {
      this.reader.close();
      this.reader = null;
    }
    return this.reader ??= new FootprintReader();
  }

  /**
   * Samples the footprint of the process tree of `pid()`. `aborted` settles with the memory limit
   * or a failed reading, whichever ends the process first, from its start to its answer: the
   * warm-up runs the engine too. `jobPeakMiB` is the highest sample since `jobSent` was called.
   */
  private watchFootprint (pid: () => number | null, reader: FootprintReader): { watch: FootprintWatch, aborted: Promise<{ failure: string }>, jobSent: () => void, jobPeakMiB: () => number | null } {
    let abort = (_failure: string) => {};
    const aborted = new Promise<{ failure: string }>(resolve => {
      abort = failure => resolve({ failure });
    });
    let jobSent = false;
    let jobPeakMiB: number | null = null;
    const watch = new FootprintWatch(pid, SAMPLE_INTERVAL_MS, mib => {
      if (jobSent) jobPeakMiB = Math.max(jobPeakMiB ?? 0, mib);
      if (mib > this.limitMiB) abort(`out of memory: exceeded ${this.limitMiB} MiB`);
    }, err => abort(`memory not measured: ${err.message}`), reader);
    return { watch, aborted, jobSent: () => {
      jobSent = true;
    }, jobPeakMiB: () => jobPeakMiB };
  }

  /** Ends a settled job's process in the background, for the next job and `stop` to wait for. */
  private endLater (end: () => Promise<void>, pid: number | undefined): void {
    // The job has settled, so a failure to end its process (e.g. an `error` from the process while
    // its exit is awaited) is reported here, not as an unhandled rejection.
    const ended = end().catch(err => {
      process.stderr.write(`job process ${pid} not ended cleanly: ${messageOf(err)}\n`);
    });
    this.ending.add(ended);
    ended.finally(() => this.ending.delete(ended));
  }

  async run (job: J): Promise<JobOutcome<R>> {
    await Promise.all(this.ending);
    const reader = this.liveReader();
    const child = this.spawn();
    const pid = child.pid;
    const died = exitCause(child);
    const alive = () => child.exitCode == null && child.signalCode == null;
    const { watch, aborted, jobSent, jobPeakMiB } = this.watchFootprint(() => alive() ? pid ?? null : null, reader);
    const end = () => endProcess(child, reader, watch);

    const startFailure = await this.ready(child, aborted, died);
    if (startFailure != null) {
      await end();
      return { failure: `start failed: ${startFailure}`, baselineFootprintMiB: null, peakFootprintMiB: watch.peakMiB };
    }
    let baselineFootprintMiB: number | null = null;
    try {
      baselineFootprintMiB = await treeFootprintMiB(pid as number, reader);
    }
    catch (err) {
      await end();
      return { failure: `memory not measured: ${messageOf(err)}`, baselineFootprintMiB, peakFootprintMiB: null };
    }
    // A process that died while its baseline was read has already emitted its exit.
    if (!alive()) {
      await end();
      return { failure: await died(child.exitCode, child.signalCode), baselineFootprintMiB, peakFootprintMiB: baselineFootprintMiB };
    }
    jobSent();
    const pressure = await PressureWatch.start(reader, SAMPLE_INTERVAL_MS);
    const outcome = await this.answer(child, job, aborted, died);
    const memoryPressure = await pressure.stop();
    this.endLater(end, pid);
    const peaks = [ baselineFootprintMiB, jobPeakMiB() ].filter(mib => mib != null);
    return { ...outcome, baselineFootprintMiB, peakFootprintMiB: peaks.length ? Math.max(...peaks) : null, memoryPressure };
  }

  async stop (): Promise<void> {
    await Promise.all(this.ending);
    this.reader?.close();
    this.reader = null;
  }
}
