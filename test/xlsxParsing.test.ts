import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { getAllRowsFromWasm } from '../src/wasmParser';
import { diffWorkbooks } from '../src/diff';
import type { CellValue, WorkbookModel } from '../src/types';

function esc(value: string) {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

async function makeWorkbook(cells: Record<string, string | number>) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`);
  zip.folder('_rels')!.file('.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);
  zip.folder('xl')!.file('workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`);
  zip.folder('xl')!.folder('_rels')!.file('workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`);

  const rows = new Map<number, string[]>();
  for (const [address, value] of Object.entries(cells)) {
    const rowNumber = Number(address.match(/\d+$/)?.[0]);
    const cellXml = typeof value === 'number'
      ? `<c r="${address}"><v>${value}</v></c>`
      : `<c r="${address}" t="inlineStr"><is><t>${esc(value)}</t></is></c>`;
    rows.set(rowNumber, [...(rows.get(rowNumber) ?? []), cellXml]);
  }
  const sheetData = [...rows.entries()].sort(([a], [b]) => a - b).map(([row, xml]) => `<row r="${row}">${xml.join('')}</row>`).join('');
  zip.folder('xl')!.folder('worksheets')!.file('sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sheetData}</sheetData></worksheet>`);
  return zip.generateAsync({ type: 'uint8array' });
}

function normalize(value: unknown) {
  return value === undefined || value === '' ? null : value as CellValue;
}

async function parseModel(name: string, bytes: Uint8Array): Promise<WorkbookModel> {
  const rows = (await getAllRowsFromWasm(bytes)).map((row) => row.map(normalize));
  const cells = new Map<string, CellValue>();
  let maxCols = 0;
  rows.forEach((row, r) => {
    maxCols = Math.max(maxCols, row.length);
    row.forEach((value, c) => {
      if (value !== null) cells.set(`${r}:${c}`, value);
    });
  });
  return { name, rows, cells, maxRows: rows.length, maxCols };
}

describe('sample xlsx parsing and diffing', () => {
  it('parses generated XLSX files and detects changed, added, and removed cells', async () => {
    const original = await parseModel('original.xlsx', await makeWorkbook({ A1: 'Name', B1: 'Count', A2: 'Apples', B2: 3, A3: 'Oranges', B3: 5 }));
    const changed = await parseModel('changed.xlsx', await makeWorkbook({ A1: 'Name', B1: 'Count', A2: 'Apples', B2: 4, A4: 'Pears', B4: 2 }));

    expect(original.cells.get('1:1')).toBe(3);
    expect(changed.cells.get('1:1')).toBe(4);

    const diffs = diffWorkbooks(original, changed);
    expect(diffs).toEqual([
      expect.objectContaining({ address: 'B2', kind: 'changed', before: 3, after: 4 }),
      expect.objectContaining({ address: 'A3', kind: 'removed', before: 'Oranges', after: null }),
      expect.objectContaining({ address: 'B3', kind: 'removed', before: 5, after: null }),
      expect.objectContaining({ address: 'A4', kind: 'added', before: null, after: 'Pears' }),
      expect.objectContaining({ address: 'B4', kind: 'added', before: null, after: 2 }),
    ]);
  });
});
