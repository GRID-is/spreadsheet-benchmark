// A replay's job process, which JobWorker in replay/common.ts forks for each file: it loads the
// engine module named in its job data, evaluates the warm-up fixture unless the engine opts out,
// then evaluates the one file it is sent.
import { basename } from 'node:path';
import { jobData, serveJobs, WARM_UP_FIXTURE, type HeapProbe } from '../src/job-worker.ts';
import { type Case, type Spill } from '../src/model/case.ts';
import { FootprintReader } from '../src/memory.ts';
import { commandFootprint, commandReader } from './child-process.ts';
import { observeSpill } from '../src/compare.ts';
import { areaRefs, parseArea, spillArea } from '../src/model/spill.ts';
import { excelSpill, harnessTiming, valuesMatch, type CaseCell, type Engine, type EngineValue, type ReplayRow, type ValueReader, type WorkerData, type WorkerRequest, type WorkerResult } from './common.ts';

const data = jobData<WorkerData>();
process.env.REPLAY_TIMEOUT_MS = String(data.timeoutMs ?? 60000);
const { engine }: { engine: Engine } = await import(data.engineModule);
// An engine evaluated outside this process may run commands, whose footprint runLimited samples
// through this reader, which starts and answers once before the fixture is timed. If it fails, each
// command starts its own and fails the same way.
if (engine.inProcess === false) {
  commandReader.reader = new FootprintReader();
  await commandReader.reader.pressure().catch(() => {});
}

async function evaluate (request: WorkerRequest, heap?: HeapProbe): Promise<WorkerResult> {
  commandFootprint.peakMiB = null;
  const result = await evaluateFile(request, heap);
  return commandFootprint.peakMiB == null ? result : { ...result, commandPeakFootprintMiB: commandFootprint.peakMiB };
}

// The spilled cells of a case's Excel spill range, which the source report recorded.
function spilledCells (c: Case): CaseCell[] {
  return c.spill === undefined ? [] : areaRefs(spillArea(c.ref, c.spill.excelRange)).flat().filter(ref => ref !== c.ref).map(ref => ({ sheet: c.sheet, ref }));
}

// The engine's value of a spilled cell; one it fails to read is empty, as an unanswered case is.
function readSpilled (value: ValueReader<CaseCell>, cell: CaseCell): EngineValue {
  try {
    return value(cell) ?? null;
  }
  catch {
    return null;
  }
}

// A row's spill: Excel's side from the source case, the engine's from its reader.
function spillOf (c: Case, v: EngineValue, value: ValueReader<CaseCell>): Spill | undefined {
  const excelArea = spillArea(c.ref, c.spill?.excelRange);
  const excel = {
    range: c.spill?.excelRange,
    value: (ref: string) => {
      const at = parseArea(ref);
      return ref === c.ref || at === undefined ? c.excelValue : c.spill?.excelValues[at.top - excelArea.top]?.[at.left - excelArea.left] ?? null;
    },
  };
  return observeSpill(c.ref, excel, { range: value.spillRange?.(c), value: ref => ref === c.ref ? v : readSpilled(value, { sheet: c.sheet, ref }) }, c.eps);
}

async function evaluateFile ({ fixture, cases }: WorkerRequest, heap?: HeapProbe): Promise<WorkerResult> {
  try {
    harnessTiming.ms = 0;
    const t0 = Date.now();
    // The engine's reader for the cases, then the spilled cells of Excel's spill ranges.
    const value = await engine.buildFile(fixture, [ ...cases, ...cases.flatMap(spilledCells) ], heap);
    const buildMs = Date.now() - t0 - harnessTiming.ms;
    const t1 = Date.now();
    const rows = cases.map((source): ReplayRow => {
      const { spill: _sourceSpill, ...c } = source;
      const loadError = value.loadError?.(source);
      if (loadError != null) {
        return { ...c, apiaryValue: null, apiaryVersion: data.versionTag, match: false, loadError, ...excelSpill(source) };
      }
      let v = null;
      let error;
      try {
        v = value(source) ?? null;
      }
      catch (err) {
        error = err instanceof Error ? err.message : String(err);
        v = null;
      }
      const spill = spillOf(source, v, value);
      return { ...c, apiaryValue: v, apiaryVersion: data.versionTag, match: valuesMatch(c.excelValue, v, c.eps), ...(spill ? { spill } : {}), ...(error ? { evalError: error } : {}) };
    });
    const readMs = Date.now() - t1;
    // The reader holds the engine's evaluated workbook.
    if (engine.inProcess !== false) heap?.measure(value);
    return { rows, timing: { harnessMs: harnessTiming.ms, buildMs, readMs, ...(value.phasesMs ? { enginePhasesMs: value.phasesMs } : {}) } };
  }
  catch (err) {
    return { buildError: err instanceof Error ? err.message : String(err) };
  }
}

// The warm-up evaluates its fixture as a measured file is evaluated. `Engine.warmUp: false` skips it.
const warmUpCase: Case = {
  suite: 'warm-up',
  category: 'warm-up',
  file: basename(WARM_UP_FIXTURE),
  sheet: 'Sheet1',
  ref: 'A3',
  id: `${basename(WARM_UP_FIXTURE)}#Sheet1#A3`,
  formula: 'A1+A2',
  functions: [],
  excelValue: 3,
  apiaryValue: null,
  match: false,
  eps: 0,
  excelDroppedFormula: false,
  oracleApp: 'Excel',
  apiaryVersion: '',
};
await serveJobs(evaluate, engine.warmUp === false ? undefined : async () => {
  const result = await evaluate({ fixture: WARM_UP_FIXTURE, cases: [ warmUpCase ] });
  if ('buildError' in result) {
    throw new Error(result.buildError);
  }
});
