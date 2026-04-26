# office-oxide for Node.js — The Fastest Office Document Library for JavaScript & TypeScript

Native Node.js bindings for [office_oxide](https://github.com/yfedoseev/office_oxide) — a fast Rust library for parsing, converting, and editing Office documents (DOCX, XLSX, PPTX, DOC, XLS, PPT).

Links directly against the Rust C FFI via [koffi](https://koffi.dev). No `node-gyp` build step. Pre-built native libraries are shipped for Linux, macOS, and Windows (x64 + arm64).

[![npm](https://img.shields.io/npm/v/office-oxide?label=npm%20native)](https://www.npmjs.com/package/office-oxide)
[![License: MIT OR Apache-2.0](https://img.shields.io/badge/License-MIT%20OR%20Apache--2.0-blue.svg)](https://opensource.org/licenses)

> **Part of the [office_oxide](https://github.com/yfedoseev/office_oxide) toolkit.** Same Rust core, same pass rate as the
> [Rust](https://docs.rs/office_oxide), [Python](../python/README.md),
> [Go](../go/README.md), [C# / .NET](../csharp/OfficeOxide/README.md),
> and [WASM](../wasm-pkg/README.md) bindings.
>
> For running in browsers, Deno, Bun, or edge runtimes, use the sibling [`office-oxide-wasm`](../wasm-pkg) package instead.

## Quick Start

```bash
npm install office-oxide
```

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

With the disposable protocol (Node 22+):

```js
using doc = Document.open('report.docx');
console.log(doc.plainText());
```

## Why office-oxide?

- **Fast** — 0.8ms mean DOCX, 5.0ms mean XLSX, 0.7ms mean PPTX; up to 100× faster than python-docx / openpyxl / python-pptx
- **Reliable** — 98.4% pass rate on 6,062 files; zero failures on legitimate Office documents
- **Complete** — 6 formats: DOCX, XLSX, PPTX + legacy DOC, XLS, PPT
- **Permissive** — MIT / Apache-2.0, no AGPL or GPL restrictions
- **No node-gyp** — Ships pre-built native libraries; no C++ build toolchain required
- **Full TypeScript support** — Type definitions ship in the package

## Performance

Benchmarked on 6,062 files from 11 independent public test suites. Single-thread, release build, warm disk cache.

| Library | Format | Mean | Pass Rate | License |
|---------|--------|------|-----------|---------|
| **office_oxide** | **DOCX** | **0.8ms** | **98.9%** | **MIT** |
| python-docx | DOCX | 11.8ms | 95.1% | MIT |
| **office_oxide** | **XLSX** | **5.0ms** | **97.8%** | **MIT** |
| openpyxl | XLSX | 94.5ms | 96.2% | MIT |
| **office_oxide** | **PPTX** | **0.7ms** | **98.4%** | **MIT** |
| python-pptx | PPTX | 32.5ms | 86.7% | MIT |

## Installation

```bash
npm install office-oxide
```

The native shared library is resolved (in order):

1. `OFFICE_OXIDE_LIB` environment variable (absolute path).
2. `prebuilds/<platform>-<arch>/liboffice_oxide.{so|dylib|dll}` inside the npm package.
3. The system library search path.

| Platform | x64 | ARM64 |
|---|---|---|
| Linux (glibc) | Yes | Yes |
| macOS | Yes | Yes (Apple Silicon) |
| Windows | Yes | Yes |

Requires Node.js 18 or newer. TypeScript definitions ship in the package.

## Editing

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

## One-shot helpers

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

## Other languages

office_oxide ships the same Rust core through six bindings:

- **Rust** — `cargo add office_oxide` — see [docs.rs/office_oxide](https://docs.rs/office_oxide)
- **Python** — `pip install office-oxide` — see [python/README.md](../python/README.md)
- **Go** — `go get github.com/yfedoseev/office_oxide/go` — see [go/README.md](../go/README.md)
- **C# / .NET** — `dotnet add package OfficeOxide` — see [csharp/OfficeOxide/README.md](../csharp/OfficeOxide/README.md)
- **WASM** — `npm install office-oxide-wasm` — see [wasm-pkg/README.md](../wasm-pkg/README.md)

## Why I built this

I needed a library that could read all six Office formats at once — not six separate packages — and I needed it without pulling in a JVM, a Python runtime, or a GPL-licensed dependency. The same Rust binary powers Python via PyO3, Node.js via koffi, Go via cgo, C# via P/Invoke, and the browser via WASM — one fix lands everywhere.

If something's broken or missing, [open an issue](https://github.com/yfedoseev/office_oxide/issues).

— Yury

## License

MIT OR Apache-2.0

---

**Node.js** + **Rust core** | MIT / Apache-2.0 | 98.4% pass rate on 6,062 files | Up to 100× faster than alternatives | 6 formats
