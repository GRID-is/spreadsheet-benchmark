import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { basename, join } from 'node:path';
import { fileURLToPath } from 'node:url';

export const SUITES_DIR = fileURLToPath(new URL('../suites', import.meta.url));

/**
 * A suite of real workbooks rather than tests (`"kind": "real-workbooks"` in its manifest). It
 * answers another question than the parity rate, so it stays out of the default suite list, the
 * baselines and their totals, and runs only when named with --suite.
 */
export const REAL_WORKBOOKS = 'real-workbooks';

/**
 * A suite whose workbooks stay on the machine that runs it: Git ignores its folder, and only
 * aggregate numbers from its reports may be shared.
 */
export function isPrivateSuite (name: string): boolean {
  return name.startsWith('private-');
}

export function excelPassCommand (suite: string): string {
  const pass = isPrivateSuite(basename(suite)) ? 'excel-pass' : 'excel-windows-pass';
  return `npm run ${pass} -- --suite ${suite}`;
}

function kindOf (suitesDir: string, name: string): string | undefined {
  try {
    return JSON.parse(readFileSync(join(suitesDir, name, 'manifest.json'), 'utf8')).kind;
  }
  catch {
    return undefined;
  }
}

/**
 * The suites under suites/ with a manifest, sorted by name: the test suites only, unless
 * `includeAll`, which adds real-workbook and private suites.
 */
export function listSuites (suitesDir = SUITES_DIR, { includeAll = false }: { includeAll?: boolean } = {}): string[] {
  return readdirSync(suitesDir, { withFileTypes: true })
    .filter(entry => entry.isDirectory() && existsSync(join(suitesDir, entry.name, 'manifest.json')))
    .map(entry => entry.name)
    .filter(name => includeAll || (!isPrivateSuite(name) && kindOf(suitesDir, name) !== REAL_WORKBOOKS))
    .sort();
}
