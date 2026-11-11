import { readFileSync, statSync, writeFileSync } from 'node:fs';
import { basename, join } from 'node:path';
import { valuesMatch } from './compare.ts';
import { type Case } from './model/case.ts';
import { type Catalogue } from './model/catalogue.ts';
import { classify, type Classification, type Status } from './model/classify.ts';
import { jsonlLines, plainPath, readJsonl } from './report-jsonl.ts';

const STATUSES: Status[] = [ 'pass', 'fail', 'unsupported', 'error', 'skip' ];

type Row = { c: Case, k: Classification };

export function developmentReportDir (engineId: string): string {
  const timestamp = new Date().toISOString().replace(/[-:]/g, '').replace('T', '-').slice(0, 15);
  return join('reports', 'local', `${timestamp}-${engineId}`);
}

export function loadCatalogue (): Catalogue {
  return JSON.parse(readFileSync(new URL('../data/functions.json', import.meta.url), 'utf8'));
}

function tally (rows: Row[]): Record<Status, number> {
  const t: Record<Status, number> = { pass: 0, fail: 0, unsupported: 0, error: 0, skip: 0 };
  for (const r of rows) {
    t[r.k.status]++;
  }
  return t;
}

function line (label: string, rows: Row[]): string {
  const t = tally(rows);
  const counted = rows.filter(r => r.k.parity).length;
  const rate = counted ? (100 * t.pass / counted).toFixed(1) : '-';
  return `| ${label} | ${rows.length} | ${STATUSES.map(s => t[s]).join(' | ')} | ${rate}% |`;
}

function table (label: string, groups: Map<string, Row[]>, all: Row[]): string[] {
  return [
    `| ${label} | cases | ${STATUSES.join(' | ')} | parity pass rate |`,
    `|---|---|${STATUSES.map(() => '---').join('|')}|---|`,
    ...[ ...groups.keys() ].sort().map(k => line(k, groups.get(k) ?? [])),
    line('**total**', all),
    '',
  ];
}

function groupBy (rows: Row[], key: (r: Row) => string): Map<string, Row[]> {
  const m = new Map<string, Row[]>();
  for (const r of rows) {
    const k = key(r);
    m.set(k, [ ...(m.get(k) ?? []), r ]);
  }
  return m;
}

const show = (v: Case['excelValue']): string => JSON.stringify(v);

/**
 * The cases of a suite's `.jsonl`, in either form (src/report-jsonl.ts), which is blank or empty
 * for a suite left without cases.
 */
export function readCases (jsonlPath: string): Case[] {
  return readJsonl(jsonlPath).split('\n').filter(line => line.trim() !== '').map(line => JSON.parse(line));
}

/**
 * A report's cases one workbook at a time, read line by line: a suite of real workbooks can yield a
 * file of several GB, more than one string holds. Reports write each workbook's cases together, so
 * a group ends where the file name changes.
 */
export async function * caseGroups (jsonlPath: string): AsyncGenerator<[ string, Case[] ]> {
  let file: string | null = null;
  let group: Case[] = [];
  for await (const line of jsonlLines(jsonlPath)) {
    if (line.trim() === '') continue;
    const c: Case = JSON.parse(line);
    if (c.file !== file) {
      if (file !== null) yield [ file, group ];
      file = c.file;
      group = [];
    }
    group.push(c);
  }
  if (file !== null) yield [ file, group ];
}

/** The text summary of a report, or a note that it is too large for one (see src/real-models.ts). */
export function reportIfSmall (jsonlPath: string, title: string): string {
  // A run summarizes the plain form, before it compresses it.
  const size = statSync(jsonlPath).size;
  return size > 400 * 1024 * 1024 ? `${title}: ${Math.round(size / 1048576)} MiB of cases, too many for the text summary; node src/real-models.ts reads it` : report(jsonlPath, title);
}

export function report (jsonlPath: string, title: string, catalogue: Catalogue = loadCatalogue()): string {
  const cases = readCases(jsonlPath);
  const rows: Row[] = cases.map(c => ({ c, k: classify(c, catalogue) }));
  const suite = basename(plainPath(jsonlPath), '.jsonl');
  const fn = (r: Row): string => r.c.functions[0] ?? '';
  const worst = [ ...groupBy(rows.filter(r => r.k.status !== 'pass' && fn(r)), fn).entries() ]
    .sort((a, b) => b[1].length - a[1].length)
    .slice(0, 25);
  const errors = rows.filter(r => r.k.status === 'error');
  const skips = groupBy(rows.filter(r => r.k.status === 'skip'), r => r.k.skipReason ?? '');
  const annotated = rows.filter(r => r.c.upstreamLiteral !== undefined);
  const parityAnnotated = annotated.filter(r => r.k.parity);
  const disagree = parityAnnotated.filter(r => !valuesMatch(r.c.upstreamLiteral ?? null, r.c.excelValue, r.c.eps)).length;
  const outside = [ ...new Set(rows.filter(r => r.k.skipReason === 'excel-lacks-function').map(fn)) ].sort();
  const gaps = [ ...new Set(rows.filter(r => r.k.neverSupported).map(fn)) ].sort();
  const propagated = rows.filter(r => r.k.propagatedError && r.k.status === 'fail').length;
  const wrongErrors = rows.filter(r => r.k.failKind === 'wrong-error').length;
  const spillFails = rows.filter(r => r.k.failKind === 'wrong-spill-size' || r.k.failKind === 'wrong-spilled-value').length;
  const summary = [
    `# ${suite} vs ${title}`,
    '',
    `Catalogue: Apiary ${catalogue.source.apiaryVersion}. Pass rate is over parity cases only. Skips: ${[ ...skips.entries() ].map(([ k, v ]) => `${k} ${v.length}`).join(', ') || 'none'}. Fails where Apiary's error came from another cell: ${propagated}. Fails where both answered an error, but a different one: ${wrongErrors}. Fails where the case's own cell matches but its spill does not: ${spillFails}.`,
    `Functions outside Excel's list (${outside.length}): ${outside.join(', ') || 'none'}. Never-supported functions present: ${gaps.join(', ') || 'none'}.`,
    ...(annotated.length ? [ `Upstream literal differs from Excel's value in ${disagree} of ${parityAnnotated.length} annotated parity cases.` ] : []),
    '',
    ...table('category', groupBy(rows, r => r.c.category), rows),
    ...table('saved by', groupBy(rows, r => r.c.oracleApp || '(none)'), rows),
    '## Functions with most non-passing cases',
    '',
    '| function | non-pass | statuses | example |',
    '|---|---|---|---|',
    ...worst.map(([ name, group ]) => {
      const t = tally(group);
      const ex = group[0].c;
      return `| ${name} | ${group.length} | ${STATUSES.filter(s => t[s]).map(s => `${s} ${t[s]}`).join(', ')} | \`${ex.file}!${ex.ref}\` =${ex.formula.slice(0, 60)} → Excel ${show(ex.excelValue)}, Apiary ${show(ex.apiaryValue)} |`;
    }),
    '',
    ...(errors.length ? [ '## Errors', '', ...errors.map(e => `- \`${e.c.file}\`: ${e.c.loadError ?? ''}`), '' ] : []),
  ].join('\n');
  writeFileSync(plainPath(jsonlPath).replace(/\.jsonl$/, '.md'), summary);
  return summary;
}

if (process.argv[1] && import.meta.url.endsWith(process.argv[1].split('/').pop() ?? '')) {
  const path = process.argv[2];
  if (!path) {
    throw new Error('usage: node src/report.ts <suite.jsonl[.bz2]> [title]');
  }
  console.log(report(path, process.argv[3] ?? 'Apiary'));
}
