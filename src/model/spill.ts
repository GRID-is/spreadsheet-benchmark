// Spill ranges (CONTEXT.md) as the case record stores them: `C1:D3`, or the anchor alone, `C1`.

export type Area = { top: number, left: number, bottom: number, right: number };

function column (letters: string): number {
  let col = 0;
  for (const ch of letters) {
    col = col * 26 + ch.charCodeAt(0) - 64;
  }
  return col - 1;
}

function letters (col: number): string {
  let s = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  }
  return s;
}

/** The area of `C1:D3` or `C1`, zero-based; undefined for any other text. */
export function parseArea (range: string): Area | undefined {
  const m = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(range);
  if (!m || m[1] === undefined || m[2] === undefined) {
    return undefined;
  }
  const [ c1, r1 ] = [ column(m[1]), Number(m[2]) - 1 ];
  const [ c2, r2 ] = m[3] === undefined || m[4] === undefined ? [ c1, r1 ] : [ column(m[3]), Number(m[4]) - 1 ];
  return { top: Math.min(r1, r2), left: Math.min(c1, c2), bottom: Math.max(r1, r2), right: Math.max(c1, c2) };
}

export function cellRef (row: number, col: number): string {
  return `${letters(col)}${row + 1}`;
}

/** `C1:D3`, or `C1` for a single cell. */
export function areaRange (area: Area): string {
  const start = cellRef(area.top, area.left);
  return area.bottom === area.top && area.right === area.left ? start : `${start}:${cellRef(area.bottom, area.right)}`;
}

/** The references of an area's cells, by row. */
export function areaRefs (area: Area): string[][] {
  const rows: string[][] = [];
  for (let r = area.top; r <= area.bottom; r++) {
    const row: string[] = [];
    for (let c = area.left; c <= area.right; c++) {
      row.push(cellRef(r, c));
    }
    rows.push(row);
  }
  return rows;
}

/**
 * The spill range anchored at `anchor` that `range` names, in the record's form: the anchor alone
 * where `range` is absent, not a range, or not anchored there.
 */
export function spillArea (anchor: string, range: string | null | undefined): Area {
  const at = parseArea(anchor) ?? { top: 0, left: 0, bottom: 0, right: 0 };
  const area = range == null ? undefined : parseArea(range);
  return area !== undefined && area.top === at.top && area.left === at.left ? area : at;
}

export function isSingleCell (area: Area): boolean {
  return area.top === area.bottom && area.left === area.right;
}
