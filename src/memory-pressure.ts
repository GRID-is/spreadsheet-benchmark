// The memory pressure (CONTEXT.md) of a run: the system's level at its start, the highest level
// sampled during it, and the level at its end. Memory pressure leaves the footprint readings alone
// but can slow every process, so timings taken under it may be inflated.
import { FootprintReader, type PressureLevel } from './memory.ts';

/**
 * The memory pressure of a run. `perFile` marks a report whose timings rows, if any, record the
 * memory pressure of their file where it was elevated; reports made before they did lack it.
 */
export type MemoryPressure = { start: PressureLevel, peak: PressureLevel, end?: PressureLevel, perFile?: true };

/** The memory pressure while one file was timed: the level at its start and the highest level. */
export type FilePressure = { start: PressureLevel, peak: PressureLevel };

/** A timed file under elevated memory pressure, named `<suite>/<file>`. */
export type PressedFile = { file: string } & FilePressure;

export type PressureSource = { pressure (): Promise<PressureLevel> };

export const PRESSURE_SAMPLE_MS = 2000;

const RANK: Record<PressureLevel, number> = { unknown: 0, normal: 1, warn: 2, critical: 3 };

export function higherPressure (a: PressureLevel, b: PressureLevel): PressureLevel {
  return RANK[b] > RANK[a] ? b : a;
}

export function isElevated (level: PressureLevel | undefined): boolean {
  return level === 'warn' || level === 'critical';
}

/** A report's memory pressure, marked as recorded per file on its timings rows. */
export function perFile (pressure: MemoryPressure): MemoryPressure {
  return { ...pressure, perFile: true };
}

/**
 * The memory pressure of a run that continued another: a report and its timings, or an append. It
 * is per file only where both are.
 */
export function combinePressure (first: MemoryPressure, then: MemoryPressure): MemoryPressure {
  return {
    start: first.start,
    peak: higherPressure(first.peak, then.peak),
    ...(then.end == null ? {} : { end: then.end }),
    ...(first.perFile && then.perFile ? { perFile: true as const } : {}),
  };
}

/**
 * The memory pressure of a run, `current`, combined with that recorded in the manifest of the
 * report it appends to, if any. A report without a valid record is from before memory pressure was
 * recorded: its pressure is unknown, and its timings rows record none.
 */
export function continuedPressure (appendedTo: { memoryPressure?: unknown } | null, current: MemoryPressure): MemoryPressure {
  if (appendedTo == null) {
    return current;
  }
  return combinePressure(recordedPressure(appendedTo.memoryPressure) ?? { start: 'unknown', peak: 'unknown' }, current);
}

/** The memory pressure recorded in a manifest, or null where it has none or a malformed one. */
export function recordedPressure (value: unknown): MemoryPressure | null {
  if (value == null || typeof value !== 'object') {
    return null;
  }
  const { start, peak, end, perFile } = value as Record<string, unknown>;
  if (!isLevel(start) || !isLevel(peak) || (end != null && !isLevel(end))) {
    return null;
  }
  return { start, peak, ...(end == null ? {} : { end }), ...(perFile === true ? { perFile } : {}) };
}

const isLevel = (v: unknown): v is PressureLevel => typeof v === 'string' && Object.hasOwn(RANK, v);

/** The fields a timings row gets for its file's memory pressure: none unless it was elevated. */
export function rowPressure (pressure: FilePressure | undefined): { memoryPressure?: FilePressure } {
  return pressure != null && isElevated(pressure.peak) ? { memoryPressure: { start: pressure.start, peak: pressure.peak } } : {};
}

/** The timed files that timings rows record as under elevated memory pressure, in row order. */
export function pressedFiles (rows: Array<{ suite?: unknown, file?: unknown, memoryPressure?: unknown }>): PressedFile[] {
  return rows.flatMap(({ suite, file, memoryPressure }) => {
    const pressure = recordedPressure(memoryPressure);
    return pressure != null && isElevated(pressure.peak) ? [ { file: `${String(suite)}/${String(file)}`, start: pressure.start, peak: pressure.peak } ] : [];
  });
}

const GIB = 1024 ** 3;

/**
 * The multiple of the memory limit that the available memory should reach before a report, or
 * memory pressure will likely rise while an engine nears the limit and inflate the timings.
 * Calibrated on one observation only: with 19 GiB available, reports reached the warn level.
 */
export const MEMORY_MARGIN = 2;

/** Whether the available memory, where known, is below MEMORY_MARGIN times the memory limit. */
export function isMemoryLow (availableBytes: number | null, limitMiB: number): boolean {
  return availableBytes != null && availableBytes < MEMORY_MARGIN * limitMiB * 1024 * 1024;
}

export type AppMemory = { app: string, bytes: number };

/**
 * The app a process belongs to, named by the outermost `.app` bundle in its executable path, so
 * that a helper inside an app counts as the app; else the process's command name.
 */
export function appOf (command: string): string {
  const bundle = /(?:^|\/)([^/]+)\.app(?:\/|$)/.exec(command);
  return bundle?.[1] ?? command.slice(command.lastIndexOf('/') + 1);
}

/**
 * The apps using the most memory, biggest first: at most `count` of them, each over `minBytes`,
 * from the output of `ps -Ao rss=,comm=` (each process's resident memory in KiB and its command).
 */
export function memoryHungryApps (ps: string, count = 5, minBytes = GIB): AppMemory[] {
  const totals = new Map<string, number>();
  for (const line of ps.split('\n')) {
    const match = /^\s*(\d+)\s+(.+)$/.exec(line);
    if (match != null) {
      const app = appOf(match[2].trim());
      totals.set(app, (totals.get(app) ?? 0) + Number(match[1]) * 1024);
    }
  }
  return [ ...totals ]
    .map(([ app, bytes ]) => ({ app, bytes }))
    .filter(({ bytes }) => bytes > minBytes)
    .sort((a, b) => b.bytes - a.bytes)
    .slice(0, count);
}

/** Bytes in GiB, to `digits` decimals, without trailing zeros. */
export function gib (bytes: number, digits = 1): string {
  return String(Number((bytes / GIB).toFixed(digits)));
}

/**
 * What a report's memory pressure says about its timings, or null if nothing. A report that records
 * it per file is judged by its `pressed` timed files alone, another by its own levels.
 */
export function pressureProblem (pressure: MemoryPressure | null, pressed: PressedFile[] = []): string | null {
  if (pressure?.perFile) {
    return pressedFilesProblem(pressed);
  }
  if (pressure == null || !isElevated(pressure.peak)) {
    return null;
  }
  return `memory pressure reached the ${pressure.peak} level (${pressure.start} at the start, ${pressure.end ?? 'unrecorded'} at the end), so the timings may be inflated`;
}

/**
 * What the timed files under elevated memory pressure say about a report's timings, or null if
 * there are none. A file that started under it may have been slowed by memory used before it, by
 * other processes or an earlier file; a file during which it rose likely raised it with its own
 * engine.
 */
export function pressedFilesProblem (pressed: PressedFile[]): string | null {
  const before = pressed.filter(f => isElevated(f.start));
  const rose = pressed.filter(f => !isElevated(f.start));
  const clauses: string[] = [];
  if (before.length) {
    clauses.push(`${timedFiles(before)} started under memory pressure (${fileList(before)}), so ${before.length === 1 ? 'its' : 'their'} timings may be inflated`);
  }
  if (rose.length) {
    const peak = rose.reduce<PressureLevel>((level, f) => higherPressure(level, f.peak), 'unknown');
    clauses.push(`memory pressure rose to the ${peak} level during ${timedFiles(rose)} (${fileList(rose)}), likely from ${rose.length === 1 ? 'its' : 'their'} own engine`);
    if (!before.length) {
      clauses.push('no other timed file was under it');
    }
  }
  return clauses.length ? clauses.join('; ') : null;
}

const timedFiles = (files: unknown[]) => `${files.length} timed file${files.length === 1 ? '' : 's'}`;

/** The files' names, at most `max` of them and a count of the rest. */
export function fileList (files: Array<{ file: string }>, max = 3): string {
  const names = files.map(f => f.file);
  if (names.length > max) {
    return `${names.slice(0, max).join(', ')} and ${names.length - max} more`;
  }
  return names.length > 1 ? `${names.slice(0, -1).join(', ')} and ${names.at(-1)}` : names.join('');
}

/**
 * Samples the system's memory pressure until stopped: once at the start, every `intervalMs`, and
 * once at the stop. A failed reading counts as unknown and ends the sampling, since the pressure
 * qualifies the timings and is no reason to fail a run. It reads through `source` if one is passed;
 * otherwise it starts its own footprint reader and closes it on stop.
 */
export class PressureWatch {
  private start: PressureLevel = 'unknown';
  private peak: PressureLevel = 'unknown';
  private stopped = false;
  private failed = false;
  private loop: Promise<void> = Promise.resolve();
  private wake = () => {};
  private readonly source: PressureSource;
  private readonly close: () => void;

  private constructor (source: PressureSource, close: () => void) {
    this.source = source;
    this.close = close;
  }

  static async start (source?: PressureSource, intervalMs = PRESSURE_SAMPLE_MS): Promise<PressureWatch> {
    const watch = source == null ? PressureWatch.withOwnReader() : new PressureWatch(source, () => {});
    watch.start = await watch.read();
    watch.peak = watch.start;
    watch.loop = watch.sample(intervalMs);
    return watch;
  }

  private static withOwnReader (): PressureWatch {
    const reader = new FootprintReader();
    return new PressureWatch(reader, () => reader.close());
  }

  private async read (): Promise<PressureLevel> {
    if (this.failed) {
      return 'unknown';
    }
    try {
      return await this.source.pressure();
    }
    catch {
      this.failed = true;
      return 'unknown';
    }
  }

  private async sample (intervalMs: number): Promise<void> {
    while (!this.stopped && !this.failed) {
      await new Promise<void>(resolve => {
        const timer = setTimeout(resolve, intervalMs);
        this.wake = () => {
          clearTimeout(timer);
          resolve();
        };
      });
      if (this.stopped) {
        break;
      }
      this.peak = higherPressure(this.peak, await this.read());
    }
  }

  /** The memory pressure so far, without an end. */
  record (): MemoryPressure {
    return { start: this.start, peak: this.peak };
  }

  async stop (): Promise<MemoryPressure> {
    this.stopped = true;
    this.wake();
    await this.loop;
    const end = await this.read();
    this.close();
    this.peak = higherPressure(this.peak, end);
    return { start: this.start, peak: this.peak, end };
  }
}
