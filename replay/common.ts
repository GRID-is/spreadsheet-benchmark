// Shared loop for replaying an existing report's cases through another engine.
// The source report supplies the case list (file, sheet, ref, formula, Excel
// value, eps); the fixture supplies the cells; the engine under test supplies
// recalculated values.
import { readFileSync, mkdirSync, mkdtempSync, writeFileSync, appendFileSync, existsSync, rmSync, realpathSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { tmpdir } from 'node:os';
import { join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { type Workbook } from '@borgar/xlsx-convert';
import { ZipArchive } from '@borgar/zip';
import { type CellValue } from '../src/compare.ts';
import { forkJobProcess, formatJobMemory, JobWorker, type HeapMemory, type HeapProbe, type JobMemory, type JobOutcome, type ProcessMemory } from '../src/job-worker.ts';
import { type Case } from '../src/model/case.ts';
import { fixturesDigest } from '../src/fixture-state.ts';
import { executionEnvironment, type ExecutionEnvironment } from '../src/execution-environment.ts';
import { MEASUREMENT_REVISION, measurement, TIMEOUT_MS } from '../src/measurement.ts';
import { continuedPressure, perFile, PressureWatch, rowPressure, type FilePressure, type MemoryPressure } from '../src/memory-pressure.ts';
import { caseGroups, developmentReportDir, reportIfSmall } from '../src/report.ts';
import { compressJsonl, readJsonl, removeJsonl, storedPath, writePlainJsonl } from '../src/report-jsonl.ts';
import { isPrivateSuite } from '../src/suites.ts';
export { numbersAreClose, valuesMatch } from '../src/compare.ts';

// The cell of a case that an engine is asked for.
export type CaseCell = Pick<Case, 'sheet' | 'ref'>;

// What an engine answers for a case cell: a scalar, an error string, or null (or undefined) for
// an empty cell.
export type EngineValue = CellValue;

// Reads the engine's value of a cell. `spillRange`, where the engine reports one, returns the spill
// range (`C1:D3`, or the cell alone) of the formula at a case's cell. `loadError`, where the engine
// answered for only part of the file, returns why it has no answer for a case's cell; the case then
// fails as the cases of a file that failed to load do. `phasesMs`, where the engine times its own
// work, holds those times in milliseconds by phase, which the file's timing records as reported.
export type ValueReader<C extends CaseCell = Case> = ((c: C) => EngineValue) & {
  spillRange?: (c: CaseCell) => string | undefined,
  loadError?: (c: CaseCell) => string | undefined,
  phasesMs?: Record<string, number>,
};

// An engine under replay. Each runner exports one as `engine`, checked with `satisfies Engine`.
// buildFile(fixturePath, cells, heap) returns the reader of the engine's value for each of `cells`:
// the cases, then the spilled cells of Excel's spill ranges (CONTEXT.md). It may throw to fail the
// whole file. An engine evaluated outside the job process may report its own heap figures with
// `heap.report`.
export type Engine = {
  name: string,
  version: string,
  options?: Record<string, string | number | boolean>,
  // Skip timings.jsonl, for an engine whose wall time is not engine compute.
  noTimings?: boolean,
  // Whether a file's build error is one that every later file would repeat, such as an exhausted
  // API quota. The replay then stops with an error instead of recording each remaining file as a
  // load error.
  stopsReplay? (buildError: string): boolean,
  // Set to false to skip the warm-up fixture, which the job process otherwise evaluates before it
  // reports ready: for an engine whose measured work starts afresh on every file anyway (a child
  // process, a browser), or that cannot evaluate a fixture outside the suites.
  warmUp?: boolean,
  // Set to false for an engine evaluated outside the job process (a command line, a browser, a web
  // service): the job process's heap is then the harness's own, so the worker does not measure it.
  inProcess?: false,
  // Set for an engine that sends the fixture off the machine (a web service): it is never given the
  // fixtures of a private suite.
  sendsWorkbooks?: true,
  buildFile (fixturePath: string, cells: CaseCell[], heap?: HeapProbe): ValueReader<CaseCell> | Promise<ValueReader<CaseCell>>,
};

/** Excel's side of a source case's spill, for a row of a workbook that the engine failed to load. */
export function excelSpill (c: Case): Pick<Case, 'spill'> {
  return c.spill === undefined ? {} : { spill: { excelRange: c.spill.excelRange, excelValues: c.spill.excelValues } };
}

// A report row a replay writes: the source case with the engine's answer.
export type ReplayRow = Case & { evalError?: string };

export type FileTiming = { harnessMs: number, buildMs: number, readMs: number, enginePhasesMs?: Record<string, number> };

// Messages between replay() and the worker (replay/worker.ts).
export type WorkerData = { engineModule: string, versionTag: string, timeoutMs?: number };
export type WorkerRequest = { fixture: string, cases: Case[] };
export type WorkerResult = ({ rows: ReplayRow[], timing: FileTiming } | { buildError: string }) & { commandPeakFootprintMiB?: number };

// A row of a report's timings.jsonl: a file's timing or its failure, its memory as far as measured,
// and the memory pressure while the file was evaluated where it was elevated.
type TimingRow = { engine: string, suite: string, file: string, cases: number, memoryPressure?: FilePressure }
  & (FileTiming | { failure: string }) & Partial<JobMemory> & Partial<HeapMemory>;

function readRows<T> (path: string): T[] {
  return storedPath(path) != null ? readJsonl(path).split('\n').filter(Boolean).map(line => JSON.parse(line)) : [];
}

function writeRows (path: string, rows: unknown[]): void {
  writePlainJsonl(path, rows.map(row => JSON.stringify(row)).join('\n') + (rows.length ? '\n' : ''));
}

export function refToCoords (ref: string): { col: number, row: number } {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m) {
    throw new Error(`not a cell reference: ${ref}`);
  }
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col: col - 1, row: Number(m[2]) - 1 };
}

// Wall-clock the harness spends during the current buildFile call (reading
// the fixture in readCellInput, reading an engine's output in cachedValues,
// browser setup in spread-sheets.ts); the worker resets and reads it so that
// time is excluded from an engine's own build time in timings.jsonl.
export const harnessTiming = { ms: 0 };

// Imports @borgar/xlsx-convert on first use, not when this module loads: some engines never use
// it, and a top-level import adds about 13 ms to the start of every job process.
const xlsxConvert = () => import('@borgar/xlsx-convert');

type JsfCell = NonNullable<Workbook['sheets'][number]['cells']>[string];

// A cell as readCellInput hands it to an engine: a formula with its cached value, or a literal
// with its type. An array formula, dynamic or CSE, also has its range (F), and a CSE one is marked
// cse.
export type InputCell = { f?: JsfCell['f'], v?: JsfCell['v'], t?: JsfCell['t'], F?: JsfCell['F'], cse?: JsfCell['cse'] };

export type InputSheet = { name: string, cells: Record<string, InputCell> };

// Loads a fixture into [{ name, cells: { ref: { f?, v? } } }] for the engines
// with cell input. The reader is @borgar/xlsx-convert, an independent
// open-source xlsx reader, not Apiary: the engine under test is then handed
// the file's own cells and formula text, and Apiary is nowhere in its input
// path. With cellFormulas the reader expands shared formulas into one
// A1-style formula per cell. Checked against Apiary's reading over the
// whole corpus on 2026-10-01: 185,981 formula cells, 24 differ: 20 single-cell
// ranges (`A5:A5`, which Apiary renders `A5`), 3 3-D references that Apiary
// 17.0.0 misreads, and 1 sheet name that Apiary quotes (`'LOG10'!LOG10`).
export async function readCellInput (fixturePath: string): Promise<InputSheet[]> {
  const t0 = Date.now();
  try {
    return await readCellInputInner(fixturePath);
  }
  finally {
    harnessTiming.ms += Date.now() - t0;
  }
}

async function readCellInputInner (fixturePath: string): Promise<InputSheet[]> {
  const jsf = await (await xlsxConvert()).convert(fixturePath, { cellFormulas: true });
  return jsf.sheets.map(sheet => {
    const cells: Record<string, InputCell> = {};
    for (const [ref, cell] of Object.entries(sheet.cells ?? {})) {
      // A formula-less cell inside another cell's array or spill range (F) holds
      // only that array's cached value; writing it as a literal would collide
      // with the anchor's spill in the engine under test.
      if (cell.f == null && cell.F) continue;
      if (cell.f == null && cell.v == null) continue;
      if (cell.f == null) cells[ref] = { v: cell.v, t: cell.t };
      else cells[ref] = { f: cell.f, v: cell.v ?? null, ...(cell.F != null ? { F: cell.F } : {}), ...(cell.cse ? { cse: true } : {}) };
    }
    return { name: sheet.name, cells };
  });
}

// read the values saved in a workbook (uses @borgar/xlsx-convert so there's no recalculation)
export async function cachedValues (workbookPath: string): Promise<ValueReader<CaseCell>> {
  const t0 = Date.now();
  try {
    return await cachedValuesInner(workbookPath);
  }
  finally {
    harnessTiming.ms += Date.now() - t0;
  }
}

// Excel holds at most 32,767 characters in a cell, so a longer string is an answer that Excel cannot
// give. An engine can still write one, of hundreds of megabytes, which the job could not send back.
const EXCEL_MAX_TEXT = 32767;
const XML_ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

/**
 * Shorten each XML text node that reads as more characters than a cell can hold to the part that
 * reads as its first EXCEL_MAX_TEXT + 1 UTF-16 code units, and the rest of a character they split.
 * The result is still too long for a cell, so it never matches Excel's value.
 *
 * It reads a node as @borgar/xlsx-convert does: XML entities, then each OOXML escape _xHHHH_ as the
 * one code unit it stands for, from left to right, so `_x005F_x000D_` reads as `_x000D_`. The cut
 * keeps every escape whole.
 */
export function shortenOverlongText (xml: string): string {
  const parts: string[] = [];
  let kept = 0;
  for (let start = xml.indexOf('>'); start !== -1; start = xml.indexOf('>', start + 1)) {
    const end = xml.indexOf('<', start + 1);
    if (end === -1) break;
    if (end - start - 1 > EXCEL_MAX_TEXT) {
      // A code unit is at most an escape of 7 characters, each at most 10 characters of XML, so
      // this prefix holds the code units to keep
      const text = unescapeXml(xml.slice(start + 1, Math.min(end, start + 1 + 70 * (EXCEL_MAX_TEXT + 2))));
      // One code unit as the reader reads it: an escape where one starts, else one character
      const unit = /_x([\da-f]{4})_|[\s\S]/giy;
      let cut = 0;
      let split = false;
      let match: RegExpExecArray | null;
      for (let units = 0; units <= EXCEL_MAX_TEXT + Number(split) && (match = unit.exec(text)) != null; units++) {
        cut = unit.lastIndex;
        split = units === EXCEL_MAX_TEXT && /[\uD800-\uDBFF]/.test(match[1] ? String.fromCharCode(parseInt(match[1], 16)) : match[0]);
      }
      if (cut < text.length) {
        parts.push(xml.slice(kept, start + 1), text.slice(0, cut).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'));
        kept = end;
      }
    }
    start = end;
  }
  return parts.length === 0 ? xml : parts.join('') + xml.slice(kept);
}

function unescapeXml (escaped: string): string {
  return escaped.replace(/&(?:#x([0-9a-f]+)|#(\d+)|(\w+));/gi, (entity, hex, dec, name) =>
    hex ? String.fromCodePoint(parseInt(hex, 16)) : dec ? String.fromCodePoint(Number(dec)) : XML_ENTITIES[name] ?? entity);
}

async function cachedValuesInner (workbookPath: string): Promise<ValueReader<CaseCell>> {
  const bytes = readFileSync(workbookPath);
  const zip = new ZipArchive(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
  const preserve = (tag: string) => tag.replace(/\s+xml:space=(?:"[^"]*"|'[^']*')/, '').replace(/(\/?>)$/, ' xml:space="preserve"$1');
  // simple-xml drops whitespace-only text without xml:space. Annotate the
  // in-memory string nodes, not the saved export or its cached values.
  for (const file of zip.files.filter(f => /^xl\/(?:worksheets\/[^/]+|sharedStrings)\.xml$/.test(f))) {
    const xml = await zip.readText(file);
    if (xml === undefined) {
      throw new Error(`cannot read ${file} in ${workbookPath}`);
    }
    const preserved = shortenOverlongText(xml).replace(/<t(?:\s[^>]*)?>/g, preserve)
      .replace(/<c\b([^>]*?)(?<!\/)>[\s\S]*?<\/c>/g, (cell, attrs) =>
        /\bt\s*=\s*(?:"str"|'str')/.test(attrs) ? cell.replace(/<v(?:\s[^>]*)?>/g, preserve) : cell);
    if (preserved !== xml) await zip.write(file, preserved);
  }
  const { convertBinary } = await xlsxConvert();
  const sheets = new Map((await convertBinary(zip.toArrayBuffer(), workbookPath, { skipMerged: false })).sheets.map(s => [ s.name, s.cells ]));
  // The range of the array formula the workbook stores at a cell. A writer may store a formula that
  // spilled without a range, so a cell without one reports nothing, not a one-cell spill.
  const spillRange = (c: CaseCell) => sheets.get(c.sheet)?.[c.ref]?.F ?? undefined;
  return Object.assign((c: CaseCell) => sheets.get(c.sheet)?.[c.ref]?.v ?? null, { spillRange });
}

/**
 * Check that the given path exists and return it.
 *
 * Convert-and-read engines (Gnumeric, GridCraft, LibreOffice) answer with an
 * xlsx whose cached values are their results. Their converters can exit 0
 * after a failed import, so a missing output file is the failure signal.
 *
 * @returns path if it exists
 * @throws {Error} if the path does not exist, saying the tool wrote nothing, and why where its
 *   exit says so (the signal that killed it, or a nonzero status), then the beginning of
 *   res.stderr or res.stdout
 */
export function guardConvertedFileExists (path: string, tool: string, res: { stderr?: string, stdout?: string, status?: number | null, signal?: NodeJS.Signals | null }): string {
  if (!existsSync(path)) {
    const exit = res.signal != null ? ` (killed by ${res.signal})` : res.status ? ` (exit status ${res.status})` : '';
    throw new Error(`${tool} wrote nothing${exit}: ${(res.stderr || res.stdout || '').trim().slice(0, 300)}`);
  }
  return path;
}

// A converter writes a recalculated copy of the fixture into dir and returns its path.
export type Converter = (fixture: string, dir: string) => string | Promise<string>;

export async function convertAndRead (convert: Converter, fixture: string): Promise<ValueReader<CaseCell>> {
  const dir = mkdtempSync(join(tmpdir(), 'convert-and-read-'));
  try {
    return await cachedValues(await convert(fixture, dir));
  }
  finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

// Manifest paths are relative to the repository root, so a report records the same
// provenance wherever it was captured.
const repoRoot = fileURLToPath(new URL('..', import.meta.url));
const repoPath = (path: string) => relative(repoRoot, path);

// A report manifest. A replay copies its source's and replaces the engine, replay, memory pressure
// and suites entries; a suite's entry records the digest of the fixtures that the replay
// evaluated.
type ReplayManifest = {
  createdAt?: string,
  environment?: ExecutionEnvironment,
  engine?: unknown,
  replay?: unknown,
  measurement?: unknown,
  memoryPressure?: unknown,
  apiary?: unknown,
  suites: Record<string, { fixtures?: string, files?: number, cases?: number }>,
};

// What a file's evaluation yields: the worker's result, or the failure that ended it.
export type FileResult = { rows?: ReplayRow[], timing?: FileTiming, buildError?: string, memory?: JobMemory & { processMemory?: ProcessMemory, heap?: HeapMemory | null }, memoryPressure?: MemoryPressure };

/**
 * The heap figures in a line that the shim of an engine in another runtime prints, in MiB:
 *   #memory \t M \t <retained heap> <retained external memory> <peak live heap>
 * with `-` for a retained external memory that the runtime does not measure. Null for any other
 * line.
 */
export function parseHeapLine (line: string): HeapMemory | null {
  const [ key, , payload ] = line.split('\t');
  if (key !== '#memory') return null;
  const [ retained, external, peak ] = (payload ?? '').split(' ');
  return { retainedHeapMiB: Number(retained), ...(external === '-' ? {} : { retainedExternalMiB: Number(external) }), peakLiveHeapMiB: Number(peak) };
}

/**
 * A file's footprint readings, with the peak raised, where higher, to the baseline plus the peak of
 * the commands the engine ran. The job's watch samples the job's process tree, command included,
 * on its own interval, and can miss a command that lives less than one; runLimited samples the
 * command alone from its spawn. The sum assumes that the job process stays at its baseline
 * meanwhile.
 */
export function fileMemory ({ baselineFootprintMiB, peakFootprintMiB }: JobMemory, commandPeakFootprintMiB?: number): JobMemory {
  if (commandPeakFootprintMiB == null || baselineFootprintMiB == null) {
    return { baselineFootprintMiB, peakFootprintMiB };
  }
  const withCommand = baselineFootprintMiB + commandPeakFootprintMiB;
  return { baselineFootprintMiB, peakFootprintMiB: Math.max(peakFootprintMiB ?? 0, withCommand) };
}

/** What a file's evaluation yields, from the outcome of its job. */
export function fileResult (outcome: JobOutcome<WorkerResult>): FileResult {
  if (!('result' in outcome)) {
    const { baselineFootprintMiB, peakFootprintMiB, memoryPressure } = outcome;
    return { buildError: outcome.failure, memory: { baselineFootprintMiB, peakFootprintMiB }, memoryPressure };
  }
  const { commandPeakFootprintMiB, ...result } = outcome.result;
  return { ...result, memory: { ...fileMemory(outcome, commandPeakFootprintMiB), processMemory: outcome.processMemory, heap: outcome.heap }, memoryPressure: outcome.memoryPressure };
}

/**
 * The timings.jsonl row of a file: its timing, or the first line of its failure (at most 200
 * characters), then its footprint, rounded to whole MiB like the heap figures, and its heap
 * figures. `processMemory` is left to the console line: read when the job answers, it counts
 * garbage not yet collected.
 */
export function timingRow (key: { engine: string, suite: string, file: string, cases: number }, result: FileResult): TimingRow {
  const { baselineFootprintMiB, peakFootprintMiB, heap } = result.memory ?? {};
  const round = (mib: number | null | undefined) => mib == null ? mib : Math.round(mib);
  const footprint = result.memory ? { baselineFootprintMiB: round(baselineFootprintMiB), peakFootprintMiB: round(peakFootprintMiB) } : {};
  const failure = (result.buildError ?? 'no timing').split('\n')[0].slice(0, 200);
  return { ...key, ...(result.timing ?? { failure }), ...footprint, ...heap };
}

/** Throws when an engine that sends workbooks off the machine would be given a private suite. */
export function refusePrivateSuites (engine: Pick<Engine, 'name' | 'sendsWorkbooks'>, suites: string[]): void {
  const refused = engine.sendsWorkbooks === true ? suites.filter(isPrivateSuite) : [];
  if (refused.length) {
    throw new Error(`${engine.name} sends workbooks off the machine, so it never replays a private suite: ${refused.join(', ')}`);
  }
}

export async function replay (engine: Engine, engineModule: string, argv = process.argv.slice(2)): Promise<void> {
  const engineId = `${engine.name}-${engine.version}`;
  const versionTag = `${engine.name}@${engine.version}`;
  const suiteFilter: string[] = [];
  let source: string | null = null;
  let out: string | null = null;
  let only: string | null = null;
  let append = false;
  let timeoutMs = TIMEOUT_MS;
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--suite') suiteFilter.push(argv[++i]);
    else if (argv[i] === '--source') source = argv[++i];
    else if (argv[i] === '--out') out = argv[++i];
    else if (argv[i] === '--only') only = argv[++i];
    else if (argv[i] === '--append') append = true;
    else if (argv[i] === '--timeout') timeoutMs = Number(argv[++i]);
  }
  if (!source) throw new Error('--source <report folder> is required');
  const sourceManifest: ReplayManifest = JSON.parse(readFileSync(join(source, 'manifest.json'), 'utf8'));
  const sourceRevision = (sourceManifest.measurement as { revision?: unknown } | undefined)?.revision;
  if (sourceRevision !== MEASUREMENT_REVISION) {
    // Its cases lack what this revision measures, such as Excel's spill ranges, without which every
    // spill range that the engine reports would fail as a wrong spill size.
    throw new Error(`the source report is at measurement revision ${String(sourceRevision ?? 'unrecorded')}, not ${MEASUREMENT_REVISION}: replay a report of the current revision`);
  }
  const unknown = suiteFilter.filter(suite => !Object.hasOwn(sourceManifest.suites, suite));
  if (unknown.length) throw new Error(`--suite names what is not a suite of the source report: ${unknown.join(', ')}`);
  const suites = suiteFilter.length ? suiteFilter : Object.keys(sourceManifest.suites);
  refusePrivateSuites(engine, suites);
  if (!out) {
    out = developmentReportDir(engineId);
  }
  mkdirSync(out, { recursive: true });
  if (realpathSync(out) === realpathSync(source)) {
    throw new Error('Source and output must be different report directories.');
  }
  const manifestPath = join(out, 'manifest.json');
  const previous: ReplayManifest | null = existsSync(manifestPath) ? JSON.parse(readFileSync(manifestPath, 'utf8')) : null;
  const engineMetadata = { name: engine.name, version: engine.version, source: repoPath(fileURLToPath(engineModule)), hash: createHash('sha256').update(readFileSync(new URL(engineModule))).digest('hex'), options: engine.options ?? {} };
  const replayMetadata = { source: repoPath(resolve(source)), sourceManifestHash: createHash('sha256').update(JSON.stringify(sourceManifest)).digest('hex'), timeoutMs };
  const environment = executionEnvironment();
  if (append && previous && (JSON.stringify(previous.engine) !== JSON.stringify(engineMetadata) || JSON.stringify(previous.replay) !== JSON.stringify(replayMetadata) || JSON.stringify(previous.measurement) !== JSON.stringify(measurement(timeoutMs)) || JSON.stringify(previous.environment) !== JSON.stringify(environment))) {
    throw new Error('Cannot append with a different engine, source report, run options or execution environment; start a new report.');
  }
  // The cases of a suite to evaluate, one workbook at a time, streamed from the source report: a
  // suite of real workbooks can yield more cases than fit in memory or in one string.
  async function * casesByFile (suite: string, sourceDir: string): AsyncGenerator<[ string, Case[] ]> {
    for await (const [ file, group ] of caseGroups(join(sourceDir, `${suite}.jsonl`))) {
      if (only && !file.startsWith(only)) continue;
      const cases = group.filter(r => !r.loadError && !r.excelDroppedFormula);
      if (cases.length) yield [ file, cases ];
    }
  }
  const filesOf = async (suite: string, sourceDir: string) => {
    const files = new Set<string>();
    for await (const [ file ] of casesByFile(suite, sourceDir)) files.add(file);
    return files;
  };
  if (append && previous) {
    // A suite's entry records one fixtures digest, so rows kept from the earlier run must have been
    // evaluated in the current fixtures.
    const changed: string[] = [];
    for (const suite of suites) {
      const recorded = previous.suites[suite]?.fixtures;
      if (recorded == null || recorded === fixturesDigest(resolve('suites', suite))) continue;
      const files = await filesOf(suite, source);
      if (readRows<ReplayRow>(join(out, `${suite}.jsonl`)).some(row => !files.has(row.file))) changed.push(suite);
    }
    if (changed.length) {
      throw new Error(`Cannot append: the fixtures of ${changed.join(', ')} changed since the rows the append would keep were evaluated; start a new report.`);
    }
  }
  const manifest: ReplayManifest = append && previous ? previous : {
    ...sourceManifest,
    engine: engineMetadata,
    environment,
    replay: replayMetadata,
    measurement: measurement(timeoutMs),
    suites: {},
  };
  manifest.createdAt = new Date().toISOString();
  delete manifest.apiary;
  // The source's memory pressure is not the replay's; an append continues its own report's.
  const appendedTo = append ? previous : null;
  const pressure = await PressureWatch.start();
  const writeManifest = () => {
    manifest.memoryPressure = continuedPressure(appendedTo, perFile(pressure.record()));
    writeFileSync(manifestPath, JSON.stringify(manifest, null, 2) + '\n');
  };
  const timingsPath = join(out, 'timings.jsonl');
  if (!append) {
    for (const suite of Object.keys(previous?.suites ?? {})) {
      removeJsonl(join(out, `${suite}.jsonl`));
      rmSync(join(out, `${suite}.md`), { force: true });
    }
    writePlainJsonl(timingsPath, '');
  }
  writeManifest();

  // Each file is evaluated in a fresh job process (replay/worker.ts), so a runaway formula or
  // exhausted memory costs one file, not the run, and no file inherits another's memory: on
  // timeout, crash or the memory limit the process tree is killed and the file is recorded as a
  // load error.
  const jobWorker = new JobWorker<WorkerRequest, WorkerResult>(
    () => forkJobProcess(new URL('./worker.ts', import.meta.url), { engineModule, versionTag, timeoutMs } satisfies WorkerData),
    timeoutMs,
  );

  async function evaluateFile (fixture: string, cases: Case[]): Promise<FileResult> {
    return fileResult(await jobWorker.run({ fixture, cases }));
  }

  for (const suite of suites) {
    const files = await filesOf(suite, source);
    const outPath = join(out, `${suite}.jsonl`);
    delete manifest.suites[suite];
    writeManifest();
    const kept = append ? readRows<ReplayRow>(outPath).filter(row => !files.has(row.file)) : [];
    writeRows(outPath, kept);
    writeRows(timingsPath, readRows<TimingRow>(timingsPath).filter(row => row.suite !== suite || !files.has(row.file)));
    const completedFiles = new Set(kept.map(row => row.file));
    let completedCases = kept.length;
    for await (const [file, cases] of casesByFile(suite, source)) {
      const t0 = Date.now();
      const fixture = resolve('suites', suite, 'excel', file.replace(/\.xls$/, '.xlsx'));
      const result: FileResult = existsSync(fixture) ? await evaluateFile(fixture, cases) : { buildError: 'fixture missing' };
      if (result.buildError != null && engine.stopsReplay?.(result.buildError)) {
        await jobWorker.stop();
        manifest.memoryPressure = continuedPressure(appendedTo, perFile(await pressure.stop()));
        writeFileSync(manifestPath, JSON.stringify(manifest, null, 2) + '\n');
        const resume = suites.slice(suites.indexOf(suite)).map(s => `--suite ${s}`).join(' ');
        throw new Error(`${versionTag} stopped at ${suite}/${file}: ${result.buildError}\n`
          + `The report in ${out} is incomplete: its manifest lists only the suites completed before ${suite}. `
          + `To resume: node ${repoPath(fileURLToPath(engineModule))} --source ${source} --append --out ${out} ${resume}`);
      }
      // A file that fails to build still owes an answer for each of its cases.
      // Collapsing it to one row drops the rest out of the denominator, which
      // flatters an engine that crashes on a large workbook over one that
      // answers every cell in it wrongly, and leaves engines incomparable
      // because each drops a different set of cases.
      const rows: ReplayRow[] = result.rows ?? cases.map(c => ({
        ...c, apiaryValue: null, apiaryVersion: versionTag, match: false, loadError: result.buildError, ...excelSpill(c),
      }));
      appendFileSync(outPath, rows.map(r => JSON.stringify(r)).join('\n') + '\n');
      completedFiles.add(file);
      completedCases += rows.length;
      if (engine.noTimings !== true) {
        const timing: TimingRow = { ...timingRow({ engine: versionTag, suite, file, cases: rows.length }, result), ...rowPressure(result.memoryPressure) };
        appendFileSync(timingsPath, JSON.stringify(timing) + '\n');
      }
      const bad = rows.filter(r => !r.match).length;
      console.log(`${suite}/${file}: ${rows.length} cases, ${bad} not matching, ${Date.now() - t0} ms${result.memory ? `, ${formatJobMemory(result.memory)}` : ''}${result.buildError ? ` (${result.buildError})` : ''}`);
    }
    reportIfSmall(outPath, `${versionTag} (${suite})`);
    compressJsonl(outPath);
    manifest.suites[suite] = { ...sourceManifest.suites[suite], fixtures: fixturesDigest(resolve('suites', suite)), files: completedFiles.size, cases: completedCases };
    writeManifest();
    console.log(`# ${suite} done`);
  }
  await jobWorker.stop();
  compressJsonl(timingsPath);
  manifest.memoryPressure = continuedPressure(appendedTo, perFile(await pressure.stop()));
  writeFileSync(manifestPath, JSON.stringify(manifest, null, 2) + '\n');
  console.log('report:', out);
}
