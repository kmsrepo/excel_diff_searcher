declare module 'xlsx-wasm-browser/index_bg.js' {
  export function __wbg_set_wasm(wasm: WebAssembly.Exports): void;
  export function get_all_rows(bytes: Uint8Array): (string | number | boolean | undefined)[][];
}
