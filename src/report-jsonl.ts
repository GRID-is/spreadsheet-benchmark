// The JSONL files of a report: a suite's cases and timings.jsonl, as `<name>.jsonl.bz2` or as the
// plain `<name>.jsonl` of an unfinished run (docs/viewer-requirements.md#report-format-input).
// Where both exist the plain one is current, and the `.bz2` may be stale or partial.
//
// Functions here take the plain path, `<name>.jsonl`, and resolve the form on disk. A path that
// names the `.bz2` form reads that file where it exists, so list a folder with jsonlNames, not by
// its file names.
import { execFile, execFileSync, spawn } from 'node:child_process';
import { copyFileSync, createReadStream, existsSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { basename, join } from 'node:path';
import { createInterface } from 'node:readline';
import { promisify } from 'node:util';

const BZ2 = '.bz2';

/** The plain form of a report JSONL path, whichever form it names. */
export function plainPath (path: string): string {
  return path.endsWith(BZ2) ? path.slice(0, -BZ2.length) : path;
}

/** The file on disk that holds the JSONL file `path`, or null where neither form exists. */
export function storedPath (path: string): string | null {
  const plain = plainPath(path);
  const candidates = path.endsWith(BZ2) ? [ path, plain ] : [ plain, plain + BZ2 ];
  return candidates.find(candidate => existsSync(candidate)) ?? null;
}

/** The name of a report JSONL file without its `.jsonl` or `.jsonl.bz2`, or null for another file. */
export function jsonlName (fileName: string): string | null {
  return /^(.+)\.jsonl(?:\.bz2)?$/.exec(fileName)?.[1] ?? null;
}

/**
 * The names of the report JSONL files in the directory `dir`, each once whichever forms it has.
 * Read and write each as `<name>.jsonl`, so that the plain form wins over a stale `.bz2`.
 */
export function jsonlNames (dir: string): string[] {
  return [ ...new Set(readdirSync(dir).map(jsonlName).filter(name => name != null)) ].sort();
}

function stored (path: string): string {
  const found = storedPath(path);
  if (found == null) {
    throw Object.assign(new Error(`ENOENT: no such file, neither ${plainPath(path)} nor its ${BZ2}`), { code: 'ENOENT' });
  }
  return found;
}

/** The text of the JSONL file `path`, from either form; throws ENOENT where neither exists. */
export function readJsonl (path: string): string {
  const file = stored(path);
  return file.endsWith(BZ2) ? execFileSync('bzip2', [ '-dc', file ], { maxBuffer: Infinity, encoding: 'utf8' }) : readFileSync(file, 'utf8');
}

const execFileAsync = promisify(execFile);

/** readJsonl without blocking the event loop, for the viewer server. */
export async function readJsonlAsync (path: string): Promise<string> {
  const file = stored(path);
  return file.endsWith(BZ2) ? (await execFileAsync('bzip2', [ '-dc', file ], { maxBuffer: Infinity, encoding: 'utf8' })).stdout : await readFile(file, 'utf8');
}

/**
 * The lines of the JSONL file `path`, streamed: a suite of real workbooks can yield a file of
 * several GB, more than one string holds.
 */
export async function * jsonlLines (path: string): AsyncGenerator<string> {
  const file = stored(path);
  if (!file.endsWith(BZ2)) {
    yield * createInterface({ input: createReadStream(file), crlfDelay: Infinity });
    return;
  }
  const child = spawn('bzip2', [ '-dc', file ], { stdio: [ 'ignore', 'pipe', 'pipe' ] });
  let stderr = '';
  child.stderr.setEncoding('utf8').on('data', (chunk: string) => { stderr += chunk; });
  const exited = new Promise<number | null>((resolve, reject) => {
    child.on('error', reject);
    child.on('close', resolve);
  });
  // A consumer that stops early never awaits the exit.
  exited.catch(() => {});
  try {
    yield * createInterface({ input: child.stdout, crlfDelay: Infinity });
    const code = await exited;
    if (code !== 0) {
      throw new Error(`bzip2 -dc ${file}: ${stderr.trim() || `exit code ${code}`}`);
    }
  }
  finally {
    if (child.exitCode == null) child.kill();
  }
}

/** Writes the plain form of the JSONL file `path`, for a run to append to, and removes its `.bz2`. */
export function writePlainJsonl (path: string, text: string): void {
  const plain = plainPath(path);
  writeFileSync(plain, text);
  rmSync(plain + BZ2, { force: true });
}

/** Compresses the plain form of the JSONL file `path`, where it exists, replacing it. */
export function compressJsonl (path: string): void {
  const plain = plainPath(path);
  if (existsSync(plain)) {
    // -f replaces a .bz2 left by an earlier run. bzip2 removes the plain file only once the .bz2 is
    // complete.
    execFileSync('bzip2', [ '-f', plain ]);
  }
}

/** Writes the JSONL file `path` in its compressed form, removing the plain form. */
export function writeJsonl (path: string, text: string): void {
  writePlainJsonl(path, text);
  compressJsonl(path);
}

/** Removes both forms of the JSONL file `path`. */
export function removeJsonl (path: string): void {
  const plain = plainPath(path);
  rmSync(plain, { force: true });
  rmSync(plain + BZ2, { force: true });
}

/**
 * Copies the JSONL file `from`, in the form it is stored in, into the directory `toDir`, and
 * removes the other form there. Returns the path of the copy.
 */
export function copyJsonl (from: string, toDir: string): string {
  const file = stored(from);
  const target = join(toDir, basename(file));
  removeJsonl(target);
  copyFileSync(file, target);
  return target;
}
