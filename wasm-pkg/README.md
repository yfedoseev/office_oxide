# office-oxide-wasm

Fast Office document processing (DOCX, XLSX, PPTX, DOC, XLS, PPT) compiled to WebAssembly.
Rust core, zero JavaScript runtime dependencies. Works in Node.js, bundlers
(Webpack/Vite/Rollup), and modern browsers.

## Install

```bash
npm install office-oxide-wasm
```

## Quick start

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

## Package layout

`office-oxide-wasm` exports three targets via the `exports` map:

| Subpath | Target | Use when |
| --- | --- | --- |
| `office-oxide-wasm` | bundler (ESM) | Default for bundlers; Node picks CJS automatically |
| `office-oxide-wasm/node` | Node CommonJS | Pure Node without a bundler |
| `office-oxide-wasm/web` | browser ESM | Direct `<script type="module">` in the browser |
| `office-oxide-wasm/bundler` | bundler ESM | Explicit bundler entry |

## License

MIT OR Apache-2.0
