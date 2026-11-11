import { execFile, spawn, type ChildProcess } from 'node:child_process';
import { join } from 'node:path';
import { createInterface } from 'node:readline';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);

/**
 * The memory one engine, or Excel, may use to evaluate one workbook. An engine's processes are
 * held to it, and Excel's growth measured against it, by sampling physical footprint, which counts
 * everything charged to the process and can overshoot between samples. Excel is not stopped over
 * it: the Excel pass marks the file.
 */
export const MEMORY_LIMIT_MIB = 11264;

/**
 * The heap limit of a garbage-collected engine runtime (V8, the JVM): the memory limit less 1 GiB
 * for what such a process holds outside its heap. The heap limit also sets how hard the collector
 * works; one far above the memory limit lets garbage accumulate until the footprint passes it.
 */
export const HEAP_LIMIT_MIB = MEMORY_LIMIT_MIB - 1024;

type ProcessTable = { alive: Set<number>, children: Map<number, number[]> };

async function processTable (): Promise<ProcessTable> {
  const { stdout } = await execFileAsync('ps', [ '-A', '-o', 'pid=,ppid=' ]);
  const alive = new Set<number>();
  const children = new Map<number, number[]>();
  for (const line of stdout.split('\n')) {
    const [ p, pp ] = line.trim().split(/\s+/).map(Number);
    if (!Number.isFinite(p)) {
      continue;
    }
    alive.add(p);
    const siblings = children.get(pp);
    if (siblings) {
      siblings.push(p);
    }
    else {
      children.set(pp, [ p ]);
    }
  }
  return { alive, children };
}

/** A process and all its descendants, the process first, or an empty list if it has exited. */
function treeOf ({ alive, children }: ProcessTable, pid: number): number[] {
  if (!alive.has(pid)) {
    return [];
  }
  const tree: number[] = [];
  const pending = [ pid ];
  while (pending.length) {
    const p = pending.pop() as number;
    tree.push(p);
    pending.push(...(children.get(p) ?? []));
  }
  return tree;
}

/**
 * The IDs of a process and all its descendants, or an empty list if it has exited. It asks
 * `reader` if one is passed and has not failed; otherwise it runs `ps`, which suits a one-off
 * lookup, not a loop.
 */
export async function treePids (pid: number, reader?: FootprintReader): Promise<number[]> {
  if (reader != null && !reader.failed) {
    try {
      return (await reader.tree(pid)).map(member => member.pid);
    }
    catch {
      // the reader failed meanwhile; ps still answers
    }
  }
  return treeOf(await processTable(), pid);
}

/**
 * The system's memory pressure: on macOS the kernel's level, on Linux one derived from its pressure
 * stall information (`src/phys-footprint.py` has the thresholds), and unknown where neither exists.
 */
export type PressureLevel = 'normal' | 'warn' | 'critical' | 'unknown';

const PRESSURE_LEVELS: readonly string[] = [ 'normal', 'warn', 'critical', 'unknown' ] satisfies PressureLevel[];

function footprint (answer: string): number | null {
  return answer === '-' ? null : Number(answer);
}

/**
 * Reads the physical footprint of processes (CONTEXT.md) from the kernel's ledger, or on Linux
 * from `/proc`, through a Python helper, since Node cannot call `proc_pid_rusage` without a native
 * addon. The helper also finds a process's descendants and reads the system's memory pressure and
 * available memory, so that a sample starts no process. It runs until the reader is closed.
 */
export class FootprintReader {
  private readonly helper: ChildProcess;
  private readonly answers: Array<{ resolve: (line: string) => void, reject: (err: Error) => void }> = [];
  private failure: Error | null = null;

  /** Rejects every pending and later read with the first failure. */
  private fail (err: Error): void {
    this.failure ??= err;
    for (const { reject } of this.answers.splice(0)) {
      reject(this.failure);
    }
  }

  constructor () {
    this.helper = spawn(process.env.FOOTPRINT_PYTHON ?? 'python3', [ join(import.meta.dirname, 'phys-footprint.py') ], { stdio: [ 'pipe', 'pipe', 'pipe' ] });
    let stderr = '';
    this.helper.stderr?.on('data', chunk => {
      stderr += chunk;
    });
    createInterface({ input: this.helper.stdout as NodeJS.ReadableStream }).on('line', line => this.answers.shift()?.resolve(line));
    const fail = (err: Error) => this.fail(err);
    this.helper.on('error', fail);
    this.helper.stdin?.on('error', fail);
    this.helper.on('exit', (code, signal) => fail(new Error(`phys-footprint.py exited with ${signal ?? code}${stderr ? `: ${stderr.trim()}` : ''}`)));
  }

  /** Whether the helper has failed, after which every read throws. */
  get failed (): boolean {
    return this.failure != null;
  }

  private async request (line: string): Promise<string> {
    if (this.failure) {
      throw this.failure;
    }
    return new Promise<string>((resolve, reject) => {
      this.answers.push({ resolve, reject });
      this.helper.stdin?.write(line + '\n');
    });
  }

  /**
   * A process and all its descendants, the process first, each with its footprint in bytes, or null
   * for one that cannot be read, such as one that has exited. An empty list if the process does not
   * exist.
   */
  async tree (pid: number): Promise<Array<{ pid: number, bytes: number | null }>> {
    const line = await this.request(`tree ${pid}`);
    return line ? line.split(' ').map(entry => {
      const [ p, bytes ] = entry.split('=');
      return { pid: Number(p), bytes: footprint(bytes) };
    }) : [];
  }

  /** The system's memory pressure level. */
  async pressure (): Promise<PressureLevel> {
    const line = await this.request('pressure');
    return PRESSURE_LEVELS.includes(line) ? line as PressureLevel : 'unknown';
  }

  /** The memory available to processes in bytes, or null where the system does not say. */
  async available (): Promise<number | null> {
    const line = await this.request('available');
    return line === '-' ? null : Number(line);
  }

  close (): void {
    // A watch that shares this reader may still be reading; it fails rather than waits forever.
    this.fail(new Error('footprint reader closed'));
    this.helper.removeAllListeners('exit');
    this.helper.kill();
  }
}

/**
 * Physical footprint of a process and all its descendants, in MiB, or null if it has exited.
 * A reader to use may be passed; otherwise one is started for this reading.
 */
export async function treeFootprintMiB (pid: number, reader?: FootprintReader): Promise<number | null> {
  const own = reader ?? new FootprintReader();
  try {
    const tree = await own.tree(pid);
    if (tree[0]?.bytes == null) {
      return null;
    }
    return tree.reduce((total, { bytes }) => total + (bytes ?? 0), 0) / 1024 / 1024;
  }
  finally {
    if (!reader) {
      own.close();
    }
  }
}

/**
 * The output of `ps -Ao rss=,comm=`, each process's resident memory in KiB and its command, or an
 * empty string where ps fails.
 */
export async function residentMemory (): Promise<string> {
  try {
    return (await execFileAsync('ps', [ '-Ao', 'rss=,comm=' ])).stdout;
  }
  catch {
    return '';
  }
}

/** The process ID of the running application with this exact process name, or null. */
export async function pidOf (name: string): Promise<number | null> {
  try {
    const { stdout } = await execFileAsync('pgrep', [ '-x', name ]);
    const pid = Number(stdout.trim().split('\n')[0]);
    return Number.isFinite(pid) ? pid : null;
  }
  catch {
    return null;
  }
}

/**
 * Returns a function that resolves to a process ID, which it finds with `find` and remembers until
 * that process has exited. While no such process runs, every call asks `find`.
 */
export function cachedPid (find: () => Promise<number | null>): () => Promise<number | null> {
  let pid: number | null = null;
  return async () => {
    if (pid == null || !exists(pid)) {
      pid = await find();
    }
    return pid;
  };
}

function exists (pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  }
  catch (err) {
    // EPERM: the process exists but belongs to another user.
    return (err as NodeJS.ErrnoException).code === 'EPERM';
  }
}

/**
 * Samples the physical footprint of a process tree at once and then every `intervalMs` until
 * stopped, keeping the first and the highest reading. `pid` is asked again before each sample, so
 * the watched process may be relaunched. On a failed reading it stops sampling and calls
 * `onFailure`, and `stop` throws, so that a watch that could not measure is not taken for one that
 * saw no growth. It reads through `reader` if one is passed, and leaves it open; otherwise it
 * starts its own and closes it on stop.
 */
export class FootprintWatch {
  firstMiB: number | null = null;
  peakMiB: number | null = null;
  private stopped = false;
  private failure: Error | null = null;
  private readonly reader: FootprintReader;
  private readonly ownReader: boolean;
  private readonly loop: Promise<void>;
  /** Ends the wait between samples early, so that stop need not wait out the interval. */
  private wake = () => {};

  constructor (pid: () => Promise<number | null> | number | null, intervalMs: number, onSample?: (mib: number) => void, onFailure?: (err: Error) => void, reader?: FootprintReader) {
    const own = reader ?? new FootprintReader();
    this.reader = own;
    this.ownReader = reader == null;
    this.loop = (async () => {
      while (!this.stopped && !this.failure) {
        const p = await pid();
        const mib = p == null ? null : await treeFootprintMiB(p, own).catch((err: Error) => {
          this.failure = err;
          onFailure?.(err);
          return null;
        });
        if (mib != null) {
          this.firstMiB ??= mib;
          this.peakMiB = Math.max(this.peakMiB ?? 0, mib);
          onSample?.(mib);
        }
        // A stop during the reading found no wait to end, so it is checked here.
        if (this.stopped) {
          break;
        }
        await new Promise<void>(resolve => {
          const timer = setTimeout(resolve, intervalMs);
          this.wake = () => {
            clearTimeout(timer);
            resolve();
          };
        });
      }
    })();
  }

  async stop (): Promise<void> {
    this.stopped = true;
    this.wake();
    await this.loop;
    if (this.ownReader) {
      this.reader.close();
    }
    if (this.failure) {
      throw this.failure;
    }
  }
}
