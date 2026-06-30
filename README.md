# Excel Diff Searcher

An offline, single-HTML Excel comparison tool with an Excel-like grid UI and searchable cell-level differences.

## Features

- Runs fully in the browser after build; no server or network calls are needed.
- Parses `.xlsx` files with `xlsx-wasm-parser`, a WebAssembly wrapper around Rust Calamine.
- Compares the first worksheet of two workbooks by cell address.
- Highlights added, removed, and changed cells in an Excel-style grid.
- Provides searchable/filterable difference results.
- Builds to `dist/index.html`, then GitHub Actions zips it for release upload.

## Local development

```bash
npm install
npm run dev
npm run build
```

Open `dist/index.html` directly after `npm run build` to use the offline app.

## Release

Push a tag such as `v1.0.0`. GitHub Actions builds the app, compresses `dist/index.html` into `excel-diff-searcher.zip`, and uploads the zip to the GitHub Release.
