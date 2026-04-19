# office-oxide (Node native)

Native Node.js bindings for [office_oxide](https://github.com/yfedoseev/office_oxide) — a fast Rust library for parsing, converting, and editing Office documents (DOCX, XLSX, PPTX, DOC, XLS, PPT).

Links directly against the Rust C FFI via [koffi](https://koffi.dev). No `node-gyp` build step. Pre-built native libraries are shipped for Linux, macOS, and Windows (x64 + arm64).

> For running in bundlers, browsers, or edge runtimes, use the sibling [`office-oxide-wasm`](../wasm-pkg) package instead.

## Install

```bash
npm install office-oxide
```

The native shared library is resolved (in order):

1. `OFFICE_OXIDE_LIB` environment variable (absolute path).
2. `prebuilds/<platform>-<arch>/liboffice_oxide.{so|dylib|dll}` inside the npm package.
3. The system library search path.

## Quick start

```js
import { Document } from 'office-oxide';

const doc = Document.open('report.docx');
try {
  console.log(doc.format);       // "docx"
  console.log(doc.plainText());
  console.log(doc.toMarkdown());
  console.log(doc.toIr());       // structured, format-agnostic IR
} finally { doc.close(); }
```

With the new disposable protocol (Node 22+):

```js
using doc = Document.open('report.docx');
console.log(doc.plainText());
```

### Editing

```js
import { EditableDocument } from 'office-oxide';

using ed = EditableDocument.open('template.docx');
ed.replaceText('{{NAME}}', 'Alice');
ed.save('out.docx');
```

### Spreadsheet cells

```js
using ed = EditableDocument.open('report.xlsx');
ed.setCell(0, 'A1', 'Revenue');
ed.setCell(0, 'B1', 12345.67);
ed.setCell(0, 'C1', true);
ed.save('report.edited.xlsx');
```

### One-shot helpers

```js
import { extractText, toMarkdown, toHtml } from 'office-oxide';

console.log(extractText('doc.docx'));
console.log(toMarkdown('deck.pptx'));
console.log(toHtml('data.xlsx'));
```

## API

TypeScript definitions ship with the package (`office-oxide/lib/index.d.ts`).

| Export | Description |
| --- | --- |
| `Document.open(path)` / `fromBytes(data, format)` | Parse a read-only document. |
| `Document#format` | `"docx" \| "xlsx" \| …` |
| `Document#plainText()` / `toMarkdown()` / `toHtml()` / `toIr()` | Extraction methods. |
| `Document#saveAs(path)` | Save/convert to a different format. |
| `EditableDocument.open(path)` | Open DOCX/XLSX/PPTX for editing. |
| `EditableDocument#replaceText(find, replace)` | In-place replace. Returns count. |
| `EditableDocument#setCell(sheet, ref, value)` | Write an XLSX cell. |
| `EditableDocument#save(path)` | Persist to disk. |
| `version()` / `detectFormat(path)` | Library info. |
| `extractText(path)` / `toMarkdown(path)` / `toHtml(path)` | One-shot helpers. |

## License

MIT OR Apache-2.0
