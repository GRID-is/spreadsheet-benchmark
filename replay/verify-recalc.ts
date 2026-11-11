import { fileURLToPath } from 'node:url';
import { convertAndRead, type Converter } from './common.ts';

export const CANARY = fileURLToPath(new URL('recalc-canary.xlsx', import.meta.url));

// Does the converter recalculate, or copy Excel's cached values through? Prints
// the verdict and sets the exit code; run it after upgrading the engine.
export async function verifyRecalc (convert: Converter, label: string): Promise<void> {
  const got = (await convertAndRead(convert, CANARY))({ sheet: 'Sheet1', ref: 'A3' });
  const ok = got === 3;
  console.log('recalc-canary.xlsx: A3 is =A1+A2 over 1 and 2, with its cached value set to 99999.');
  console.log(`${label} returned ${got} -> ${ok
    ? 'RECALCULATED (scores are real)'
    : 'PASSED THE CACHE THROUGH (scores would be meaningless)'}`);
  process.exitCode = ok ? 0 : 1;
}
