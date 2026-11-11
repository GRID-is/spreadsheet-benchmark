// What the replays of JVM engines (such as replay/poi.ts) share: the java they run,
// its options, the output format of their shims (such as replay/PoiEvaluate.java)
// and the error of a failed process.
import { existsSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { parseHeapLine, type EngineValue } from './common.ts';
import { type HeapMemory } from '../src/job-worker.ts';
import { HEAP_LIMIT_MIB } from '../src/memory.ts';

export const JAVA = existsSync('/opt/homebrew/opt/openjdk/bin/java') ? '/opt/homebrew/opt/openjdk/bin/java' : 'java';

/**
 * JVM options that keep the JVM a background process on macOS: without them, the first use of
 * AWT (fonts, which POI touches) gives every fixture's JVM a Dock icon that takes focus
 * from the application in front, many times a minute over a run.
 */
export const HEADLESS = [ '-Djava.awt.headless=true', '-Dapple.awt.UIElement=true' ];

/**
 * Throws unless java runs, which the runner checks before it replays anything: the jars and the
 * compiled classes are made on first use. macOS has a java stub that exits 1 without a JDK.
 */
export function checkPrerequisites (): void {
  const res = spawnSync(JAVA, [ '-version' ], { encoding: 'utf8' });
  if (res.error != null || res.status !== 0) {
    throw new Error(`${JAVA} does not run (brew install openjdk)`);
  }
}

function unescape (s: string): string {
  return s.replace(/\\r/g, '\r').replace(/\\n/g, '\n').replace(/\\t/g, '\t').replace(/\\\\/g, '\\');
}

// Reads a shim's output: the value of each target cell, and the JVM's heap figures.
export function parseOutput (stdout: string): { values: Map<string, EngineValue>, heap: HeapMemory | null } {
  const values = new Map<string, EngineValue>();
  let heap: HeapMemory | null = null;
  for (const line of stdout.split('\n')) {
    const [ key, kind, payload ] = line.split('\t');
    if (key === undefined || kind === undefined) continue;
    if (key === '#memory') heap = parseHeapLine(line);
    else if (kind === 'N') values.set(key, Number(payload));
    else if (kind === 'S') values.set(key, unescape(payload ?? ''));
    else if (kind === 'B') values.set(key, payload === 'true');
    else if (kind === 'E') values.set(key, payload);
    else if (kind === 'U') values.set(key, '#NAME?');
    else if (kind === 'X') values.set(key, '#VALUE!');
    else values.set(key, null);
  }
  return { values, heap };
}

/**
 * The error of a JVM process of `engineName` that exited with `status`. The JVM, unlike V8, exits
 * with an ordinary status when its heap is exhausted, so that cause is read from stderr and
 * reported as V8's is.
 */
export function exitError (engineName: string, status: number | null, stderr: string): Error {
  if (/java\.lang\.OutOfMemoryError: Java heap space/.test(stderr)) {
    return new Error(`out of memory: exceeded the ${HEAP_LIMIT_MIB} MiB heap limit`);
  }
  return new Error(`${engineName} exit ${status}: ${stderr.trim().slice(0, 300)}`);
}
