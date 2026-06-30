import initWasm from 'xlsx-wasm-browser/index_bg.wasm?init';
import * as wasmBindings from 'xlsx-wasm-browser/index_bg.js';
import { __wbg_set_wasm, get_all_rows } from 'xlsx-wasm-browser/index_bg.js';

let wasmReady: Promise<void> | null = null;

async function ensureWasmReady() {
  wasmReady ??= initWasm({ './index_bg.js': wasmBindings }).then((instance) => {
    __wbg_set_wasm(instance.exports);
  });
  await wasmReady;
}

export async function getAllRowsFromWasm(bytes: Uint8Array | ArrayBuffer) {
  await ensureWasmReady();
  return get_all_rows(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
}
