import type { CellValue, DiffItem, WorkbookModel } from './types';

export const colName = (index: number) => {
  let n = index + 1, out = '';
  while (n > 0) { const r = (n - 1) % 26; out = String.fromCharCode(65 + r) + out; n = Math.floor((n - 1) / 26); }
  return out;
};
export const address = (row: number, col: number) => `${colName(col)}${row + 1}`;
const same = (a: CellValue, b: CellValue) => String(a ?? '') === String(b ?? '');

export function diffWorkbooks(left: WorkbookModel, right: WorkbookModel): DiffItem[] {
  const keys = new Set([...left.cells.keys(), ...right.cells.keys()]);
  return [...keys].flatMap((key) => {
    const [row, col] = key.split(':').map(Number);
    const before = left.cells.get(key) ?? null;
    const after = right.cells.get(key) ?? null;
    if (same(before, after)) return [];
    const kind: DiffItem['kind'] = before === null ? 'added' : after === null ? 'removed' : 'changed';
    return [{ key, row, col, address: address(row, col), kind, before, after }];
  }).sort((a, b) => a.row - b.row || a.col - b.col);
}

export function filterDiffs(diffs: DiffItem[], query: string, kind: string): DiffItem[] {
  const q = query.trim().toLowerCase();
  return diffs.filter((d) => (kind === 'all' || d.kind === kind) && (!q || [d.address, d.kind, d.before, d.after].some((v) => String(v ?? '').toLowerCase().includes(q))));
}
