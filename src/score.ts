// The benchmark's board, computed from the reports: every case of the public suites, pooled. A case
// counts unless `classify` skips it, and an engine's score is its passes over the cases that count.
//
//   npm run score                 the board, best first
//   npm run score -- --json       the same as JSON, with each engine's per-suite results
//   npm run score -- <report>...  only these report folders
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { classify, type Status } from './model/classify.ts';
import { loadCatalogue, readCases } from './report.ts';

export const PUBLIC_SUITES = [ 'ironcalc', 'libreoffice', 'univer', 'hyperformula', 'poi', 'gnumeric', 'fast-formula-parser', 'ironcalc-unit' ];

export type SuiteScore = Record<Status, number> & { counted: number };
export type EngineScore = SuiteScore & { engine: string, version: string, ranAt: string, report: string, suites: Record<string, SuiteScore> };

const emptyScore = (): SuiteScore => ({ pass: 0, fail: 0, unsupported: 0, error: 0, skip: 0, counted: 0 });

function suiteFile (report: string, suite: string): string {
  const compressed = join(report, `${suite}.jsonl.bz2`);
  return existsSync(compressed) ? compressed : join(report, `${suite}.jsonl`);
}

export function scoreReport (report: string, catalogue = loadCatalogue()): EngineScore {
  const manifest = JSON.parse(readFileSync(join(report, 'manifest.json'), 'utf8'));
  const total = emptyScore();
  const suites: Record<string, SuiteScore> = {};
  for (const suite of PUBLIC_SUITES) {
    const score = emptyScore();
    for (const c of readCases(suiteFile(report, suite))) {
      const k = classify(c, catalogue);
      score[k.status]++;
      total[k.status]++;
      if (k.parity) {
        score.counted++;
        total.counted++;
      }
    }
    suites[suite] = score;
  }
  return { ...total, engine: manifest.engine.name, version: manifest.engine.version, ranAt: manifest.createdAt, report, suites };
}

export const rate = (s: SuiteScore): number => (s.counted ? s.pass / s.counted : 0);

function publishedReports (): string[] {
  return readdirSync('reports', { withFileTypes: true })
    .filter(entry => entry.isDirectory() && existsSync(join('reports', entry.name, 'manifest.json')))
    .map(entry => join('reports', entry.name));
}

if (import.meta.main) {
  const args = process.argv.slice(2);
  const json = args.includes('--json');
  const reports = args.filter(arg => arg !== '--json');
  const board = (reports.length ? reports : publishedReports()).map(report => scoreReport(report)).sort((a, b) => rate(b) - rate(a));
  if (json) {
    console.log(JSON.stringify(board, null, 2));
  }
  else {
    console.log('| engine | version | ran | score | wrong per 1,000 | pass | counted |\n|---|---|---|---:|---:|---:|---:|');
    for (const s of board) {
      const wrong = (1000 * (s.counted - s.pass) / s.counted).toFixed(1);
      console.log(`| ${s.engine} | ${s.version} | ${s.ranAt.slice(0, 10)} | ${(100 * rate(s)).toFixed(2)}% | ${wrong} | ${s.pass} | ${s.counted} |`);
    }
  }
}
