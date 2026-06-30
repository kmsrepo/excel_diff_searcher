export type CellValue = string | number | boolean | null;
export type CellMap = Map<string, CellValue>;
export interface WorkbookModel { name: string; rows: CellValue[][]; cells: CellMap; maxRows: number; maxCols: number; }
export type DiffKind = 'added' | 'removed' | 'changed';
export interface DiffItem { key: string; row: number; col: number; address: string; kind: DiffKind; before: CellValue; after: CellValue; }
