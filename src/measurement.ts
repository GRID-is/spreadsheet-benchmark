// The measurement (CONTEXT.md) that every report records in its manifest, so that two baselines
// measured differently are told apart rather than compared.
import { HEAP_LIMIT_MIB, MEMORY_LIMIT_MIB } from './memory.ts';

/** The time one engine may take to evaluate one workbook, unless a run says otherwise. */
export const TIMEOUT_MS = 60000;

/**
 * The measurement revision, recorded in every report's measurement. Bump it when a change to the
 * harnesses changes what a report measures: the value comparison (`src/compare.ts`), what the
 * timings or memory figures count, or how a fixture is evaluated. Every published baseline measured
 * at an older revision is then stale.
 *
 * 1: a fresh Node process per fixture after a warm-up, the memory limit, the retained heap and
 * peak live heap, and the read-back of a convert-and-read engine's output counted as harness time.
 * 2: the footprint reader that a command is sampled through started before the fixture is timed, and
 * so counted in the baseline footprint, for an engine evaluated outside the job process.
 * 3: a case's spilled cells (CONTEXT.md) compared as well as its anchor, so the engine under test is
 * asked for them too; a fixture's array formulas evaluated as arrays in the replays; and Apiary's
 * values read from its cells, so the build time of the Apiary replay no longer counts writing its
 * JSF; and one convert-and-read engine's values read with cachedValues, which keeps a string of
 * only spaces that readCellInput read as empty.
 */
export const MEASUREMENT_REVISION = 3;

export type Measurement = { revision: number, timeoutMs: number, memoryLimitMiB: number, heapLimitMiB: number };

export function measurement (timeoutMs: number): Measurement {
  return { revision: MEASUREMENT_REVISION, timeoutMs, memoryLimitMiB: MEMORY_LIMIT_MIB, heapLimitMiB: HEAP_LIMIT_MIB };
}

/** The measurement of a report with no options changed: what a baseline must have been made with. */
export const CURRENT_MEASUREMENT: Measurement = measurement(TIMEOUT_MS);

/** How a report's recorded measurement differs from the current one, or an empty list. */
export function measurementDifferences (recorded: unknown): string[] {
  if (recorded == null || typeof recorded !== 'object') {
    return [ 'no measurement recorded' ];
  }
  const fields = recorded as Record<string, unknown>;
  return Object.entries(CURRENT_MEASUREMENT)
    .filter(([ key, value ]) => fields[key] !== value)
    .map(([ key, value ]) => `${key} ${String(fields[key] ?? 'unrecorded')}, current ${value}`);
}
