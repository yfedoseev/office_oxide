# Getting Started with office_oxide (WebAssembly)

The `office-oxide-wasm` package ships office_oxide compiled to WebAssembly. Zero native dependencies, a single Rust-core binary, and three subpath entry points so it runs anywhere JavaScript does — browsers, Node.js, and bundlers like Vite, Webpack, and Rollup.

> For Node.js without the WASM overhead, use the native [`office-oxide`](getting-started-javascript.md) package instead.

## Installation

```bash
npm install office-oxide-wasm
```

No postinstall step. Works on Node 18+, modern browsers with WebAssembly support, and anything with an ES-module / CommonJS loader.

## Subpath exports

The package exposes three builds behind `exports`:

| Import path | Target | Format |
|---|---|---|
| `office-oxide-wasm` (default) | bundler-friendly ESM (Vite, Webpack, Rollup) | ESM |
| `office-oxide-wasm/node` | Node.js | CommonJS |
| `office-oxide-wasm/web` | Browsers via `<script type="module">` or `import` | ESM, needs `init()` |

Pick the variant that matches your runtime.

## Quickstart

Extract plain text from a DOCX file.

### Node.js (CJS)

```javascript
const { readFileSync } = require('node:fs');
const { WasmDocument } = require('office-oxide-wasm/node');

const data = readFileSync('report.docx');
const doc = new WasmDocument(data, 'docx');
try {
  console.log(doc.plainText());
} finally {
  doc.free();
}
```

### Bundler (ESM)

```javascript
import { WasmDocument } from 'office-oxide-wasm';

const res = await fetch('/report.docx');
const data = new Uint8Array(await res.arrayBuffer());
const doc = new WasmDocument(data, 'docx');
try {
  console.log(doc.plainText());
} finally {
  doc.free();
}
```

### Browser (raw, no bundler)

```html
<script type="module">
  import init, { WasmDocument } from 'https://cdn.jsdelivr.net/npm/office-oxide-wasm/web/office_oxide.js';
  await init();

  const res = await fetch('/report.docx');
  const data = new Uint8Array(await res.arrayBuffer());
  const doc = new WasmDocument(data, 'docx');
  try {
    document.body.textContent = doc.plainText();
  } finally {
    doc.free();
  }
</script>
```

> **Browser only:** you must `await init()` before constructing `WasmDocument`. The Node and bundler entry points handle initialization internally.

## Core API

`WasmDocument` is the single handle — there's no separate `Document` / `EditableDocument` split in the WASM build (editing features live in the native bindings).

```javascript
import { WasmDocument } from 'office-oxide-wasm';

const doc = new WasmDocument(bytes, 'xlsx');
try {
  console.log(doc.formatName());   // "xlsx"
  console.log(doc.plainText());    // string
  console.log(doc.toMarkdown());   // string
  console.log(doc.toHtml());       // string
  const ir = doc.toIr();           // parsed object
} finally {
  doc.free();   // <-- required; WASM memory is not GC-managed
}
```

All methods are **camelCase**: `plainText`, `toMarkdown`, `toHtml`, `toIr`, `formatName`.

`bytes` must be a `Uint8Array`. `format` must be one of:

```
'docx' | 'xlsx' | 'pptx' | 'doc' | 'xls' | 'ppt'
```

The legacy binary formats (`doc`, `xls`, `ppt`) are parsed in WASM too.

## Editing (not in WASM)

`EditableDocument` — text replacement and cell writes — is **not** currently exposed through the WASM binding. For that, use:

- [native Node binding](getting-started-javascript.md) (`office-oxide`)
- [Python binding](getting-started-python.md)
- [.NET binding](getting-started-csharp.md)
- [Rust crate](getting-started-rust.md) directly

For read-only text / markdown / html / IR workflows, WASM is fully featured.

## Advanced

### Format-agnostic IR

```javascript
const doc = new WasmDocument(bytes, 'docx');
const ir = doc.toIr();
doc.free();

for (const section of ir.sections) {
  console.log(section.title);
  for (const el of section.elements) {
    // el.kind: "Heading" | "Paragraph" | "Table" | "List" | "Image" | ...
  }
}
```

The IR schema is the same as the Rust `DocumentIR`, so server-side and client-side pipelines can share a single processor.

### Bytes in, bytes out

The WASM build is bytes-in-only; there's no file I/O surface. Examples:

```javascript
// From <input type="file">
const file = inputEl.files[0];
const data = new Uint8Array(await file.arrayBuffer());
const doc = new WasmDocument(data, file.name.split('.').pop().toLowerCase());

// From fetch
const res = await fetch(url);
const data = new Uint8Array(await res.arrayBuffer());

// In Node
const data = new Uint8Array(require('node:fs').readFileSync('file.docx'));
```

### Bundlers

Vite / Webpack / Rollup resolve the `exports['.']['import']` path automatically. If your bundler complains about the `.wasm` import, add an asset rule:

```javascript
// Vite — usually zero-config
import { WasmDocument } from 'office-oxide-wasm';

// Webpack 5 — add to module.rules:
//   { test: /\.wasm$/, type: 'asset/resource' }
```

### TypeScript

Type definitions (`office_oxide.d.ts`) ship alongside the JS glue for every subpath export. Importing from the root picks up the bundler types automatically; `import type { WasmDocument } from 'office-oxide-wasm/node'` works too.

### Memory management

```javascript
const doc = new WasmDocument(bytes, 'docx');
try {
  // ... work
} finally {
  doc.free();
}
```

Forgetting `free()` leaks the backing WASM memory until the instance is torn down. If you're targeting Node 22+, the TC39 explicit resource management proposal is available behind `--harmony-explicit-resource-management`, but `try/finally` works everywhere today.

## Error Handling

Failures surface as regular `Error` instances with a descriptive message:

```javascript
try {
  const doc = new WasmDocument(new Uint8Array([0, 1, 2]), 'docx');
} catch (e) {
  console.error(`failed: ${e.message}`);
}
```

Because the WASM build uses `JsValue::from_str`, error codes aren't exposed numerically — check the message string.

## Troubleshooting

| Symptom | Fix |
|---|---|
| `ReferenceError: WebAssembly is not defined` | Your target doesn't support WASM. Use the native `office-oxide` package instead. |
| Browser: `TypeError: ... before init()` | You forgot to `await init()` when using `office-oxide-wasm/web`. |
| Bundler complains about `.wasm` | Add a WASM asset rule or switch your bundler target to `esnext`. |
| `unsupported format: pdf` | Only six formats are accepted: `docx`, `xlsx`, `pptx`, `doc`, `xls`, `ppt`. |
| Memory grows unboundedly in a hot loop | You're not calling `doc.free()` between iterations. |

## Links

- Binding source: `src/wasm.rs`
- Generated JS/TS: `wasm-pkg/{node,web,bundler}/office_oxide.{js,d.ts}`
- Package on npm: https://www.npmjs.com/package/office-oxide-wasm
- Native Node alternative: [`office-oxide`](getting-started-javascript.md)
- GitHub: https://github.com/yfedoseev/office_oxide
