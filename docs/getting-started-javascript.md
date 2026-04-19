# Getting Started with office_oxide (Node.js — native)

The `office-oxide` npm package is the **native** Node.js binding: it dynamically loads `liboffice_oxide` via [koffi](https://koffi.dev/) and exposes a small, idiomatic JavaScript API. It runs on Node 18+ and is faster and smaller than the WebAssembly build, but ships platform-specific prebuilds instead of a single portable binary.

> If you need a single artifact that runs in browsers, bundlers, or Node without native libraries, use [`office-oxide-wasm`](getting-started-wasm.md) instead.

## Installation

```bash
npm install office-oxide
```

On `postinstall`, the package verifies that a matching prebuilt native library is present under `node_modules/office-oxide/prebuilds/<platform>-<arch>/`. If none is found, configure one of these before first use:

1. **`OFFICE_OXIDE_LIB`** — absolute path to the library (e.g. `/opt/office_oxide/liboffice_oxide.so`).
2. **Prebuild directory** — drop the library into `prebuilds/<platform>-<arch>/` (e.g. `linux-x64`, `darwin-arm64`, `win32-x64`) inside the package.
3. **System library path** — install so that `liboffice_oxide.so` / `.dylib` / `office_oxide.dll` is discoverable via `LD_LIBRARY_PATH` / `DYLD_LIBRARY_PATH` / `PATH`.

## Quickstart

Extract plain text from a DOCX file:

```javascript
import { Document } from 'office-oxide';

const doc = Document.open('report.docx');
try {
  console.log(doc.plainText());
} finally {
  doc.close();
}
```

Or using TC39 explicit resource management (Node 22+):

```javascript
import { Document } from 'office-oxide';

using doc = Document.open('report.docx');
console.log(doc.plainText());
// doc.close() is called automatically at scope exit
```

One-shot helper:

```javascript
import { extractText } from 'office-oxide';
console.log(extractText('report.docx'));
```

## Core API

### `Document`

```javascript
import { Document } from 'office-oxide';

const doc = Document.open('file.xlsx');
try {
  console.log(doc.format);         // "xlsx"
  console.log(doc.plainText());    // string
  console.log(doc.toMarkdown());   // string
  console.log(doc.toHtml());       // string
  const ir = doc.toIr();           // parsed object (see "Advanced")

  // Save/convert — target format inferred from extension.
  doc.saveAs('file.docx');
} finally {
  doc.close();
}
```

Open from raw bytes (Uint8Array or Buffer):

```javascript
import { readFileSync } from 'node:fs';
import { Document } from 'office-oxide';

const data = readFileSync('report.pptx');   // Buffer is a Uint8Array
using doc = Document.fromBytes(data, 'pptx');
console.log(doc.toMarkdown());
```

`format` must be `'docx' | 'xlsx' | 'pptx' | 'doc' | 'xls' | 'ppt'`.

Module-level helpers:

```javascript
import { extractText, toMarkdown, toHtml, detectFormat, version } from 'office-oxide';

extractText('file.docx');      // string
toMarkdown('file.pptx');       // string
toHtml('file.xlsx');           // string
detectFormat('mystery.bin');   // "docx" | ... | null
version();                     // "0.1.0"
```

### `EditableDocument`

Editable handles preserve every unmodified OPC part (images, charts, relationships) on save. DOCX, XLSX, and PPTX only.

```javascript
import { EditableDocument } from 'office-oxide';

using ed = EditableDocument.open('template.docx');
const n = ed.replaceText('{{name}}', 'Alice');
console.log(`${n} replacements`);
ed.save('out.docx');
```

## Editing Examples

### Replace text across DOCX / PPTX

```javascript
import { EditableDocument } from 'office-oxide';

using ed = EditableDocument.open('slides.pptx');
ed.replaceText('Q3', 'Q4');
ed.replaceText('2024', '2025');
ed.save('slides_q4.pptx');
```

### Set XLSX cells

```javascript
import { EditableDocument } from 'office-oxide';

using wb = EditableDocument.open('budget.xlsx');

wb.setCell(0, 'A1', 'Total');   // string
wb.setCell(0, 'B1', 42.5);      // number
wb.setCell(0, 'C1', true);      // boolean
wb.setCell(0, 'D1', null);      // empty

wb.save('budget.xlsx');
```

`sheetIndex` is zero-based; `cellRef` uses standard spreadsheet notation (`A1`, `AA12`).

## Advanced

### Format-agnostic IR

`doc.toIr()` returns a parsed JavaScript object mirroring the Rust `DocumentIR`:

```javascript
using doc = Document.open('report.docx');
const ir = doc.toIr();

for (const section of ir.sections) {
  console.log(section.title);
  for (const el of section.elements) {
    // el.kind: "Heading" | "Paragraph" | "Table" | "List" | ...
  }
}
```

### Bytes-based pipelines

```javascript
import { Document } from 'office-oxide';

const res = await fetch('https://example.com/report.docx');
const data = new Uint8Array(await res.arrayBuffer());
using doc = Document.fromBytes(data, 'docx');
console.log(doc.toMarkdown());
```

### Legacy formats (DOC, XLS, PPT)

The legacy CFB parsers are first-class. Extension detection routes automatically, and `saveAs` converts via the IR:

```javascript
using doc = Document.open('old.xls');
doc.saveAs('modern.xlsx');
```

### CommonJS

The package also exposes `./lib/index.cjs`:

```javascript
const { Document } = require('office-oxide');
const doc = Document.open('file.docx');
try { console.log(doc.plainText()); } finally { doc.close(); }
```

## Error Handling

Failures throw `OfficeOxideError` with numeric `code` and operation name:

```javascript
import { Document, OfficeOxideError } from 'office-oxide';

try {
  using doc = Document.open('missing.docx');
} catch (e) {
  if (e instanceof OfficeOxideError) {
    console.error(`code=${e.code} op=${e.operation}`);
  } else {
    throw e;
  }
}
```

### Error codes

| Code | Meaning |
|---:|---|
| 0 | ok |
| 1 | invalid argument |
| 2 | io error |
| 3 | parse error |
| 4 | extraction failed |
| 5 | internal error |
| 6 | unsupported format |

## Troubleshooting

| Symptom | Fix |
|---|---|
| `office-oxide: failed to load native library` | The loader tried `OFFICE_OXIDE_LIB`, `prebuilds/<platform>-<arch>/`, and the system search path and found nothing. Set `OFFICE_OXIDE_LIB` to the absolute library path, or drop a matching prebuild into the package. |
| `koffi: ABI mismatch` | Platform/arch prebuild doesn't match this Node process (e.g. x64 Node loading arm64 binary). Reinstall or force a fresh prebuild. |
| `TypeError: data must be a Uint8Array or Buffer` | `Document.fromBytes` only accepts binary types; convert strings with `new TextEncoder().encode(...)` (rarely what you want — use `Buffer.from(base64, 'base64')` instead). |
| `Document is closed` | You called a method after `close()` or after leaving a `using` scope. Open a new handle. |
| Legacy `.doc` opens but renders gibberish | Encrypted Word 97 documents are not decrypted by office_oxide — decrypt first, e.g. via LibreOffice. |

## Links

- Binding source: `js/lib/index.js`, `js/lib/native.js`
- C header it calls: `include/office_oxide_c/office_oxide.h`
- Package on npm: https://www.npmjs.com/package/office-oxide
- WASM alternative: [`office-oxide-wasm`](getting-started-wasm.md)
- GitHub: https://github.com/yfedoseev/office_oxide
