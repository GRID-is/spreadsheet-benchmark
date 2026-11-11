import { spawn } from 'node:child_process';
import { FootprintReader, FootprintWatch, MEMORY_LIMIT_MIB, treePids } from '../src/memory.ts';

// How often runLimited samples a command's footprint, in ms. FootprintWatch takes its first sample
// at once, when the command has only just been spawned, so a command that lives a few hundred ms,
// as ssconvert does on most small fixtures, is measured by the samples after that one.
const SAMPLE_INTERVAL_MS = 50;

// An error of runLimited, with the code that names its cause.
export type LimitedError = Error & { code?: string };

export type LimitedResult = {
  status: number | null,
  signal: NodeJS.Signals | null,
  stdout: string,
  stderr: string,
  error: LimitedError | null,
  peakFootprintMiB: number | null,
};

export type LimitedOptions = { timeoutMs?: number, maxBuffer?: number, limitMiB?: number };

// The reader that runLimited samples commands through while it is open and has not failed;
// otherwise each command starts its own. Starting one starts a Python process, which takes from
// ~15 ms to over 100 ms (through a pyenv shim), so the job process of an engine run outside it opens
// this one before the fixture is timed (replay/worker.ts).
export const commandReader: { reader: FootprintReader | null } = { reader: null };

// The highest peak footprint of the commands run since `peakMiB` was last reset, which the job
// process does before each fixture.
export const commandFootprint: { peakMiB: number | null } = { peakMiB: null };

const failure = (message: string, code: string): LimitedError => Object.assign(new Error(message), { code });

// Kills the process pid, if any, and its descendants.
async function killTree (pid: number | undefined): Promise<void> {
  const pids = pid == null ? [] : await treePids(pid).catch(() => [ pid ]);
  for (const p of pids) {
    try {
      process.kill(p, 'SIGKILL');
    }
    catch {
      // already gone
    }
  }
}

// Runs command with args and, once it has exited and its output has closed,
// resolves to { status, signal, stdout, stderr, error }, like spawnSync, plus
// the peak physical footprint sampled, peakFootprintMiB. On a limit, error is
// an Error with code 'ETIMEDOUT' or 'EMEMLIMIT'; maxBuffer bounds each stream,
// as in spawnSync, and over it error has code 'ENOBUFS'. A failure to sample
// the memory kills the command too, with code 'EMEMREAD', since its memory
// would otherwise go unlimited.
export function runLimited (command: string, args: string[], { timeoutMs, maxBuffer = 64 * 1024 * 1024, limitMiB = MEMORY_LIMIT_MIB }: LimitedOptions = {}): Promise<LimitedResult> {
  return new Promise(resolve => {
    const child = spawn(command, args, { stdio: [ 'ignore', 'pipe', 'pipe' ] });
    const out: Record<'stdout' | 'stderr', Buffer[]> = { stdout: [], stderr: [] };
    const bytes = { stdout: 0, stderr: 0 };
    let error: LimitedError | null = null;
    let killing = false;
    // Once over a limit, every later chunk and sample would kill again, each with its own `ps`.
    const kill = async (err: LimitedError) => {
      error ??= err;
      if (killing) return;
      killing = true;
      await killTree(child.pid);
    };
    const collect = (name: 'stdout' | 'stderr') => (chunk: Buffer) => {
      bytes[name] += chunk.length;
      if (bytes[name] > maxBuffer) kill(failure(`${name} exceeded maxBuffer`, 'ENOBUFS'));
      else out[name].push(chunk);
    };
    child.stdout.on('data', collect('stdout'));
    child.stderr.on('data', collect('stderr'));
    const timer = timeoutMs ? setTimeout(() => kill(failure(`timeout after ${timeoutMs} ms`, 'ETIMEDOUT')), timeoutMs) : null;
    const watch = new FootprintWatch(() => child.exitCode == null && child.signalCode == null ? child.pid ?? null : null, SAMPLE_INTERVAL_MS, mib => {
      if (mib > limitMiB) kill(failure(`out of memory: exceeded ${limitMiB} MiB`, 'EMEMLIMIT'));
    }, err => kill(failure(`memory not measured: ${err.message}`, 'EMEMREAD')), commandReader.reader?.failed === false ? commandReader.reader : undefined);
    child.on('error', err => {
      error ??= err;
    });
    child.on('close', async (status, signal) => {
      if (timer) clearTimeout(timer);
      await watch.stop().catch((err: Error) => {
        error ??= err;
      });
      if (watch.peakMiB != null) {
        commandFootprint.peakMiB = Math.max(commandFootprint.peakMiB ?? 0, watch.peakMiB);
      }
      resolve({
        status,
        signal,
        stdout: Buffer.concat(out.stdout).toString('utf8'),
        stderr: Buffer.concat(out.stderr).toString('utf8'),
        error,
        peakFootprintMiB: watch.peakMiB,
      });
    });
  });
}
