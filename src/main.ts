import './style.css';
import { getAllRows } from 'xlsx-wasm-parser';
import { diffWorkbooks, filterDiffs, colName } from './diff';
import type { DiffItem, WorkbookModel } from './types';

let left: WorkbookModel | null = null;
let right: WorkbookModel | null = null;
let diffs: DiffItem[] = [];
let selected: DiffItem | null = null;

const app = document.querySelector<HTMLDivElement>('#app')!;
app.innerHTML = `
  <main class="app">
    <header class="ribbon"><h1>Excel Diff Searcher</h1><small>Offline WASM-powered XLSX comparison</small></header>
    <section class="toolbar">
      <div class="picker"><label>Original workbook</label><input id="leftFile" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"></div>
      <div class="picker"><label>Changed workbook</label><input id="rightFile" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"></div>
      <div class="actions"><input id="search" placeholder="Search address or value"><select id="kind"><option value="all">All differences</option><option value="changed">Changed</option><option value="added">Added</option><option value="removed">Removed</option></select><button id="compare" disabled>Compare</button></div>
    </section>
    <section class="layout">
      <section class="sheet"><div class="formula"><div class="namebox" id="namebox">A1</div><div class="fx">fx</div><input id="formula" readonly value="Load two .xlsx files to begin"></div><div class="gridWrap"><table class="grid" id="grid"></table></div><div class="tabs"><span class="tab">Sheet 1</span><span class="muted">WASM parser compares the first worksheet.</span></div></section>
      <aside class="panel"><h2>Difference search</h2><div><div class="stats"><div class="stat"><strong id="changed">0</strong>Changed</div><div class="stat"><strong id="added">0</strong>Added</div><div class="stat"><strong id="removed">0</strong>Removed</div></div><div class="results" id="results"></div></div></aside>
    </section>
  </main>`;

const $ = <T extends HTMLElement>(id: string) => document.querySelector<T>(id)!;
const compareBtn = $('#compare') as HTMLButtonElement;
const grid = $('#grid') as HTMLTableElement;
const results = $('#results');
const formula = $('#formula') as HTMLInputElement;
const namebox = $('#namebox');

function normalize(value: unknown) {
  return value === undefined || value === '' ? null : value;
}

async function parseFile(file: File): Promise<WorkbookModel> {
  const rows = getAllRows(new Uint8Array(await file.arrayBuffer())).map((row) => row.map(normalize));
  const cells = new Map<string, ReturnType<typeof normalize>>();
  let maxCols = 0;
  rows.forEach((row, r) => {
    maxCols = Math.max(maxCols, row.length);
    row.forEach((value, c) => { if (value !== null) cells.set(`${r}:${c}`, value); });
  });
  return { name: file.name, rows, cells, maxRows: rows.length, maxCols } as WorkbookModel;
}


async function handleFile(input: HTMLInputElement, side: 'left' | 'right') {
  const file = input.files?.[0]; if (!file) return;
  formula.value = `Parsing ${file.name} with xlsx-wasm-parser...`;
  const model = await parseFile(file);
  if (side === 'left') left = model; else right = model;
  compareBtn.disabled = !(left && right);
  formula.value = `${file.name}: ${model.maxRows} rows × ${model.maxCols} columns ready.`;
  renderGrid();
}

function compare() {
  if (!left || !right) return;
  diffs = diffWorkbooks(left, right);
  selected = diffs[0] ?? null;
  renderGrid(); renderResults();
}

function renderGrid() {
  const source = right ?? left;
  if (!source) { grid.innerHTML = '<tbody><tr><td class="muted">Choose two Excel workbooks.</td></tr></tbody>'; return; }
  const diffMap = new Map(diffs.map((d) => [d.key, d]));
  const rows = Math.max(source.maxRows, left?.maxRows ?? 0, right?.maxRows ?? 0, selected ? selected.row + 1 : 0, 30);
  const cols = Math.max(source.maxCols, left?.maxCols ?? 0, right?.maxCols ?? 0, selected ? selected.col + 1 : 0, 12);
  let html = '<thead><tr><th class="corner"></th>';
  for (let c = 0; c < cols; c++) html += `<th>${colName(c)}</th>`;
  html += '</tr></thead><tbody>';
  for (let r = 0; r < rows; r++) {
    html += `<tr><th class="rowHead">${r + 1}</th>`;
    for (let c = 0; c < cols; c++) {
      const key = `${r}:${c}`; const diff = diffMap.get(key); const value = source.cells.get(key) ?? '';
      html += `<td data-key="${key}" class="${diff?.kind ?? ''}" title="${String(value)}">${String(value)}</td>`;
    }
    html += '</tr>';
  }
  grid.innerHTML = html + '</tbody>';
}

function renderResults() {
  const visible = filterDiffs(diffs, ($('#search') as HTMLInputElement).value, ($('#kind') as HTMLSelectElement).value);
  for (const k of ['changed','added','removed']) $(`#${k}`).textContent = String(diffs.filter((d) => d.kind === k).length);
  results.innerHTML = visible.length ? visible.slice(0, 1000).map((d) => `<article class="result ${d.kind}" data-key="${d.key}"><strong>${d.address}</strong> <span class="muted">${d.kind}</span><br><span>${String(d.before ?? '∅')}</span> → <span>${String(d.after ?? '∅')}</span></article>`).join('') : '<p class="muted">No matching differences.</p>';
}

$('#leftFile').addEventListener('change', (e) => handleFile(e.target as HTMLInputElement, 'left'));
$('#rightFile').addEventListener('change', (e) => handleFile(e.target as HTMLInputElement, 'right'));
compareBtn.addEventListener('click', compare);
$('#search').addEventListener('input', renderResults);
$('#kind').addEventListener('change', renderResults);
results.addEventListener('click', (event) => { const card = (event.target as HTMLElement).closest<HTMLElement>('[data-key]'); if (!card) return; selected = diffs.find((d) => d.key === card.dataset.key) ?? null; if (selected) { namebox.textContent = selected.address; formula.value = `${selected.kind}: ${String(selected.before ?? '∅')} → ${String(selected.after ?? '∅')}`; document.querySelector(`[data-key="${selected.key}"]`)?.scrollIntoView({ block: 'center', inline: 'center' }); } });
renderGrid();
