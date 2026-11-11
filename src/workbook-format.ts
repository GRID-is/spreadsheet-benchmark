import { closeSync, openSync, readSync } from 'node:fs';
import { extname } from 'node:path';

const SIGNATURES: Array<[string, string]> = [
  [ 'd0cf11e0', '.xls' ],
  [ '504b0304', '.xlsx' ],
  [ '1f8b', '.gnumeric' ],
];

export function detectFormat (path: string): string {
  const fd = openSync(path, 'r');
  const head = Buffer.alloc(4);
  try {
    readSync(fd, head, 0, 4, 0);
  }
  finally {
    closeSync(fd);
  }
  const hex = head.toString('hex');
  return SIGNATURES.find(([ magic ]) => hex.startsWith(magic))?.[1] ?? `unknown (${hex})`;
}

// Extensions whose content is another format's container: an .xlsm is a zip like an .xlsx.
const CONTAINER_OF: Record<string, string> = { '.xlsm': '.xlsx' };

export function assertFormatMatchesExtension (path: string): void {
  const format = detectFormat(path);
  const extension = extname(path).toLowerCase();
  if (format !== (CONTAINER_OF[extension] ?? extension)) {
    throw new Error(`${path}: content is ${format}, extension says ${extname(path)}`);
  }
}
