# office-oxide-wasm — Office Documents in the Browser and Edge Runtimes

Fast Office document processing (DOCX, XLSX, PPTX, DOC, XLS, PPT) compiled to WebAssembly.
Rust core, zero JavaScript runtime dependencies. Works in Node.js, bundlers
(Webpack/Vite/Rollup), and modern browsers.

[![npm (wasm)](https://img.shields.io/npm/v/office-oxide-wasm?label=npm%20wasm)](https://www.npmjs.com/package/office-oxide-wasm)
[![License: MIT OR Apache-2.0](https://img.shields.io/badge/License-MIT%20OR%20Apache--2.0-blue.svg)](https://opensource.org/licenses)

> **Part of the [office_oxide](https://github.com/yfedoseev/office_oxide) toolkit.** Same Rust core, same pass rate as the
> [Rust](https://docs.rs/office_oxide), [Python](../python/README.md),
> [Go](../go/README.md), [C# / .NET](../csharp/OfficeOxide/README.md),
> and [Node.js native](../js/README.md) bindings.
>
> For Node.js without a bundler, consider the [native addon](../js/README.md) (`office-oxide`) — it has lower overhead and doesn't require WASM init.

## Quick Start

```bash
npm install office-oxide-wasm
```

### Node.js (CommonJS or ESM)

```js
import { readFileSync } from "node:fs";
import { WasmDocument } from "office-oxide-wasm";

const bytes = readFileSync("report.docx");
const doc = new WasmDocument(bytes, "docx");
console.log(doc.plainText());
console.log(doc.toMarkdown());
doc.free();
```

### Bundlers (Webpack, Vite, Rollup, esbuild)

```js
import { WasmDocument } from "office-oxide-wasm";

const response = await fetch("/report.docx");
const bytes = new Uint8Array(await response.arrayBuffer());
const doc = new WasmDocument(bytes, "docx");
console.log(doc.plainText());
doc.free();
```

### Browser (native ES modules, no bundler)

```html
<script type="module">
  import init, { WasmDocument } from "office-oxide-wasm/web";
  await init(); // loads the .wasm asynchronously

  const bytes = new Uint8Array(
    await (await fetch("/report.docx")).arrayBuffer()
  );
  const doc = new WasmDocument(bytes, "docx");
  document.body.textContent = doc.plainText();
  doc.free();
</script>
```

## Why office-oxide-wasm?

- **Portable** — runs anywhere JavaScript runs: browsers, Node.js, Deno, Bun, Cloudflare Workers
- **Reliable** — 98.4% pass rate on 6,062 files; zero failures on legitimate Office documents
- **Complete** — 6 formats: DOCX, XLSX, PPTX + legacy DOC, XLS, PPT
- **Permissive** — MIT / Apache-2.0, no AGPL or GPL restrictions
- **No native build** — pure WASM bundle, works without node-gyp or a C compiler

## API

`WasmDocument` is the single entry point.

| Method | Description |
| --- | --- |
| `new WasmDocument(bytes, format)` | Open from `Uint8Array`. `format`: `"docx" \| "xlsx" \| "pptx" \| "doc" \| "xls" \| "ppt"`. |
| `.formatName()` | Detected format as a string. |
| `.plainText()` | Extract plain text. |
| `.toMarkdown()` | Convert to Markdown. |
| `.toHtml()` | Convert to an HTML fragment. |
| `.toIr()` | Return the format-agnostic document IR (JS object). |
| `.free()` | Release the underlying WASM memory. `[Symbol.dispose]` is also supported. |

TypeScript definitions are shipped.

## Package Layout

`office-oxide-wasm` exports three targets via the `exports` map:

| Subpath | Target | Use when |
| --- | --- | --- |
| `office-oxide-wasm` | bundler (ESM) | Default for bundlers; Node picks CJS automatically |
| `office-oxide-wasm/node` | Node CommonJS | Pure Node without a bundler |
| `office-oxide-wasm/web` | browser ESM | Direct `<script type="module">` in the browser |
| `office-oxide-wasm/bundler` | bundler ESM | Explicit bundler entry |

## Other Languages

office_oxide ships the same Rust core through six bindings:

- **Rust** — `cargo add office_oxide` — see [docs.rs/office_oxide](https://docs.rs/office_oxide)
- **Python** — `pip install office-oxide` — see [python/README.md](../python/README.md)
- **Go** — `go get github.com/yfedoseev/office_oxide/go` — see [go/README.md](../go/README.md)
- **JavaScript (native, no WASM)** — `npm install office-oxide` — see [js/README.md](../js/README.md)
- **C# / .NET** — `dotnet add package OfficeOxide` — see [csharp/OfficeOxide/README.md](../csharp/OfficeOxide/README.md)

## Why I Built This

I needed Office document processing that works in the browser without a server round-trip — and without pulling in a JVM or a GPL-licensed dependency. The same Rust core that powers the CLI and all native bindings compiles cleanly to WASM, so feature parity is automatic.

If something's broken or missing, [open an issue](https://github.com/yfedoseev/office_oxide/issues).

— Yury

## License

MIT OR Apache-2.0

---

**WASM** + **Rust core** | MIT / Apache-2.0 | 98.4% pass rate on 6,062 files | Browser + Node.js + edge runtimes | 6 formats
