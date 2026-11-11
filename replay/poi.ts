// Replay a report's cases through Apache POI's FormulaEvaluator. Pattern C:
// replay/PoiEvaluate.java (compiled once into .cache/poi/classes) opens each
// fixture with POI's own xlsx reader and evaluates the case cells, so nothing
// is masked by Apiary's parser; per-cell evaluation means an unimplemented
// function (reported as #NAME?) poisons one case, not the file. buildMs
// includes JVM startup (~200ms) and POI's file parsing - see the Performance
// tab's caveat panel. Jars are pinned in .cache/poi/lib (Maven Central);
// java comes from Homebrew openjdk.
import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { replay, type CaseCell, type Engine, type EngineValue } from './common.ts';
import { runLimited } from './child-process.ts';
import { exitError, HEADLESS, JAVA, parseOutput } from './jvm.ts';
import { type HeapMemory, type HeapProbe } from '../src/job-worker.ts';
import { HEAP_LIMIT_MIB } from '../src/memory.ts';

const POI_VERSION = '6.0.0';
const LIB = resolve('.cache', 'poi', 'lib');
const CLASSES = resolve('.cache', 'poi', 'classes');
const CLASSPATH = `${CLASSES}:${LIB}/*`;

const MAVEN = 'https://repo1.maven.org/maven2';
const JARS = [
  `org/apache/poi/poi/${POI_VERSION}/poi-${POI_VERSION}.jar`,
  `org/apache/poi/poi-ooxml/${POI_VERSION}/poi-ooxml-${POI_VERSION}.jar`,
  `org/apache/poi/poi-ooxml-full/${POI_VERSION}/poi-ooxml-full-${POI_VERSION}.jar`,
  'org/apache/xmlbeans/xmlbeans/5.4.1/xmlbeans-5.4.1.jar',
  'org/apache/commons/commons-collections4/4.6.0/commons-collections4-4.6.0.jar',
  'org/apache/commons/commons-compress/1.28.0/commons-compress-1.28.0.jar',
  'commons-io/commons-io/2.22.0/commons-io-2.22.0.jar',
  'org/apache/logging/log4j/log4j-api/2.26.1/log4j-api-2.26.1.jar',
  'com/zaxxer/SparseBitSet/1.3/SparseBitSet-1.3.jar',
  'com/github/virtuald/curvesapi/1.08/curvesapi-1.08.jar',
  'org/apache/commons/commons-math3/3.6.1/commons-math3-3.6.1.jar',
  'commons-codec/commons-codec/1.23.0/commons-codec-1.23.0.jar',
  'org/apache/commons/commons-lang3/3.18.0/commons-lang3-3.18.0.jar',
];

function ensureJars () {
  // The classpath takes every jar in LIB, so a jar from an earlier POI_VERSION would shadow or
  // clash with the current one.
  const wanted = new Set(JARS.map(path => path.split('/').pop()));
  if (existsSync(LIB)) {
    for (const file of readdirSync(LIB)) {
      if (file.endsWith('.jar') && !wanted.has(file)) rmSync(join(LIB, file));
    }
  }
  const missing = JARS.filter(path => !existsSync(join(LIB, path.split('/').pop() ?? '')));
  if (missing.length === 0) {
    return;
  }
  mkdirSync(LIB, { recursive: true });
  for (const path of missing) {
    const res = spawnSync('curl', [ '-sfL', '-o', join(LIB, path.split('/').pop() ?? ''), `${MAVEN}/${path}` ], { encoding: 'utf8' });
    if (res.status !== 0) {
      throw new Error(`failed to download ${path}`);
    }
  }
}

export function ensureCompiled () {
  ensureJars();
  const compiled = join(CLASSES, 'PoiEvaluate.class');
  // The jars that ensureJars downloads for a new POI_VERSION are newer than a class compiled
  // against the old ones.
  const sources = [ resolve('replay', 'PoiEvaluate.java'), resolve('replay', 'JvmHeap.java') ];
  const inputs = [ ...sources, ...JARS.map(path => join(LIB, path.split('/').pop() ?? '')) ];
  if (existsSync(compiled) && statSync(compiled).mtimeMs >= Math.max(...inputs.map(file => statSync(file).mtimeMs))) {
    return;
  }
  const javac = JAVA.replace(/java$/, 'javac');
  const res = spawnSync(javac, [ '-cp', `${LIB}/*`, '-d', CLASSES, ...sources ], { encoding: 'utf8' });
  if (res.status !== 0) {
    throw new Error(`javac failed: ${(res.stderr || '').slice(0, 400)}`);
  }
}

export { checkPrerequisites } from './jvm.ts';

export const engine = {
  name: 'poi',
  version: POI_VERSION,
  warmUp: false,
  inProcess: false,
  async buildFile (fixture: string, cases: CaseCell[], heap?: HeapProbe) {
    ensureCompiled();
    const dir = mkdtempSync(join(tmpdir(), 'poi-'));
    let values: Map<string, EngineValue>;
    try {
      const targets = join(dir, 'targets.tsv');
      writeFileSync(targets, cases.map(c => `${c.sheet}\t${c.ref}`).join('\n') + '\n');
      const budget = Math.max(5000, Number(process.env.REPLAY_TIMEOUT_MS ?? 60000) - 5000);
      const res = await runLimited(JAVA, [ `-Xmx${HEAP_LIMIT_MIB}m`, ...HEADLESS, '-cp', CLASSPATH, 'PoiEvaluate', fixture, targets ], { timeoutMs: budget });
      if (res.error) throw res.error;
      if (res.status !== 0) {
        throw exitError('poi', res.status, res.stderr || '');
      }
      const output = parseOutput(res.stdout);
      values = output.values;
      if (output.heap) heap?.report(output.heap);
    }
    finally {
      rmSync(dir, { recursive: true, force: true });
    }
    return (c: CaseCell) => values.get(`${c.sheet}!${c.ref}`) ?? null;
  },
} satisfies Engine;

if (import.meta.main) {
  ensureCompiled();
  await replay(engine, import.meta.url);
}
