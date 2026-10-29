# Office Oxide — The Fastest Native Office Document Library

A fast, memory-safe library for text extraction from Office documents. Rust core with **first-class bindings for Python, Go, C#/.NET, Node.js (native and WASM), and a stable C FFI**. Handles DOCX, XLSX, PPTX, DOC, XLS, and PPT. 5–10× faster than python-docx, openpyxl, python-pptx, and xlrd; faster than python-calamine on XLSX. **Opens 97.3% of a 10,386-file test collection**; the rest are encrypted, truncated or corrupt files. MIT/Apache-2.0 dual-licensed.

> **Scope of "fastest".** Benchmarks compare Office Oxide against other
> native / embeddable libraries (no JVM runtime required): python-docx,
> openpyxl, python-pptx, python-calamine, xlrd, markitdown, catdoc,
> antiword, xls2csv, calamine (Rust), dotext, docx-rs. Apache POI and
> Apache Tika are out of scope for this comparison because they require
> a JVM and are targeted at a different deployment shape. POI/Tika
> numbers may be added in a future release.

[![Crates.io](https://img.shields.io/crates/v/office_oxide.svg)](https://crates.io/crates/office_oxide)
[![docs.rs](https://docs.rs/office_oxide/badge.svg)](https://docs.rs/office_oxide)
[![PyPI](https://img.shields.io/pypi/v/office-oxide.svg)](https://pypi.org/project/office-oxide/)
[![npm (wasm)](https://img.shields.io/npm/v/office-oxide-wasm?label=npm%20wasm)](https://www.npmjs.com/package/office-oxide-wasm)
[![npm (native)](https://img.shields.io/npm/v/office-oxide?label=npm%20native)](https://www.npmjs.com/package/office-oxide)
[![NuGet](https://img.shields.io/nuget/v/OfficeOxide)](https://www.nuget.org/packages/OfficeOxide)
[![Go Reference](https://pkg.go.dev/badge/github.com/yfedoseev/office_oxide/go.svg)](https://pkg.go.dev/github.com/yfedoseev/office_oxide/go)
[![Build Status](https://github.com/yfedoseev/office_oxide/workflows/CI/badge.svg)](https://github.com/yfedoseev/office_oxide/actions)
[![License: MIT OR Apache-2.0](https://img.shields.io/badge/License-MIT%20OR%20Apache--2.0-blue.svg)](https://opensource.org/licenses)
<!-- [![OpenSSF Scorecard](https://api.scorecard.dev/projects/github.com/yfedoseev/office_oxide/badge)](https://scorecard.dev/viewer/?uri=github.com/yfedoseev/office_oxide) -->
<!-- [![OpenSSF Best Practices](https://www.bestpractices.dev/projects/NNNN/badge)](https://www.bestpractices.dev/projects/NNNN) -->

## Available bindings

| Language | Package | Directory | Docs |
| --- | --- | --- | --- |
| Rust | `office_oxide` on crates.io | [`src/`](src/) | [lib.rs](src/lib.rs) |
| Python | `office-oxide` on PyPI | [`python/`](python/) | [python/](python/office_oxide/) |
| Go | `github.com/yfedoseev/office_oxide/go` | [`go/`](go/) | [go/README.md](go/README.md) |
| C# / .NET | `OfficeOxide` on NuGet | [`csharp/`](csharp/) | [csharp/OfficeOxide/README.md](csharp/OfficeOxide/README.md) |
| Node.js (native) | `office-oxide` on npm | [`js/`](js/) | [js/README.md](js/README.md) |
| Node.js / browser (WASM) | `office-oxide-wasm` on npm | [`wasm-pkg/`](wasm-pkg/) | [wasm-pkg/README.md](wasm-pkg/README.md) |
| C / other | header-only via FFI | [`include/office_oxide_c/`](include/office_oxide_c/) | [office_oxide.h](include/office_oxide_c/office_oxide.h) |
| CLI | `office-oxide` binary | [`crates/office_oxide_cli/`](crates/office_oxide_cli/) | |
| MCP server | `office-oxide-mcp` binary | [`crates/office_oxide_mcp/`](crates/office_oxide_mcp/) | |

Ready-to-run demos (`extract`, `replace`, `read_xlsx`) exist for every binding under [`examples/`](examples/). Deeper language-specific guides live in [`docs/`](docs/): [Rust](docs/getting-started-rust.md) · [Python](docs/getting-started-python.md) · [Go](docs/getting-started-go.md) · [C#](docs/getting-started-csharp.md) · [JavaScript (native)](docs/getting-started-javascript.md) · [WASM](docs/getting-started-wasm.md) · [C FFI](docs/getting-started-c.md).

## Quick Start

### Python
```python
from office_oxide import Document, extract_text, to_markdown

# One-liner text extraction
text = extract_text("report.docx")
markdown = to_markdown("data.xlsx")

# Context-managed document; accepts str or pathlib.Path
with Document.open("slides.pptx") as doc:
    print(doc.format)       # "pptx"
    print(doc.plain_text())
    print(doc.to_markdown())
```

```bash
pip install office-oxide
```

### Rust
```rust
use office_oxide::Document;

let doc = Document::open("report.docx")?;
let text = doc.plain_text();
let markdown = doc.to_markdown();
let ir = doc.to_ir(); // Format-agnostic intermediate representation
```

```toml
[dependencies]
office_oxide = "0.1.16"
```

### JavaScript / WASM

Browser + bundlers:

```javascript
import { WasmDocument } from "office-oxide-wasm";

const doc = new WasmDocument(buffer, "docx");
console.log(doc.plainText());
console.log(doc.toMarkdown());
doc.free();
```

Node.js native (koffi + C FFI, no node-gyp):

```javascript
import { Document } from "office-oxide";

using doc = Document.open("report.docx");
console.log(doc.format);          // "docx"
console.log(doc.plainText());
console.log(doc.toMarkdown());
console.log(doc.toIr());
```

```bash
npm install office-oxide         # native addon
npm install office-oxide-wasm    # portable WASM
```

### Go

```go
import oo "github.com/yfedoseev/office_oxide/go"

doc, _ := oo.Open("report.docx")
defer doc.Close()
text, _ := doc.PlainText()
md,   _ := doc.ToMarkdown()
```

### C# / .NET

```csharp
using OfficeOxide;

using var doc = Document.Open("report.docx");
Console.WriteLine(doc.Format);        // "docx"
Console.WriteLine(doc.PlainText());
Console.WriteLine(doc.ToMarkdown());
```

### C (raw FFI)

Include [`include/office_oxide_c/office_oxide.h`](include/office_oxide_c/office_oxide.h) and link against `liboffice_oxide`. See [`examples/c/extract.c`](examples/c/extract.c) for a working sample.

## Why office_oxide?

- **Fast** — 5–10× faster than python-docx, openpyxl, python-pptx, xlrd; faster than python-calamine on XLSX
- **Reliable** — opens 97.3% of a 10,386-file test collection; the files it does not open are encrypted, truncated or corrupt (the opt-in `default-password` feature opens the ones Office protects with its fixed default password)
- **Complete** — 6 formats: DOCX, XLSX, PPTX + legacy DOC, XLS, PPT
- **Multi-platform** — Rust, Python, Go, JavaScript/TypeScript, C#/.NET, WASM, CLI, and MCP server — one library, all platforms
- **Permissive** — MIT / Apache-2.0, no AGPL or GPL restrictions

## Performance

Benchmarked on **10,386 files** (the project's test collection: 3,519 .docx, 3,367 .xlsx, 1,123 .pptx, 780 .doc, 1,180 .xls, 417 .ppt), run on 2026-10-29. Single-thread, release build with LTO, warm disk cache. Rust figures are the median of three runs; Python and command-line figures are one run. Full methodology, pinned versions and the failure breakdown are in [BENCHMARKS.md](BENCHMARKS.md).

### DOCX

| Library | Language | Mean | Pass Rate | License |
|---------|----------|------|-----------|---------|
| **office_oxide** | **Rust** | **1.7ms** | **98.7%** | **MIT** |
| docx-rs (crate) | Rust | 5.0ms | 98.0% | MIT |
| python-docx | Python | 9.5ms | 95.7% | MIT |
| markitdown | Python | 147.5ms | 98.4% | MIT |

### XLSX

| Library | Language | Mean | Pass Rate | License |
|---------|----------|------|-----------|---------|
| calamine (crate) | Rust | 4.2ms | 98.2% | MIT |
| **office_oxide** | **Rust** | **6.2ms** | **98.6%** | **MIT** |
| python-calamine | Rust (Python) | 8.7ms | 97.8% | MIT |
| openpyxl | Python | 60.7ms | 96.0% | MIT |
| markitdown | Python | 307.5ms | 96.2% | MIT |

### PPTX

| Library | Language | Mean | Pass Rate | License |
|---------|----------|------|-----------|---------|
| **office_oxide** | **Rust** | **2.6ms** | **98.1%** | **MIT** |
| python-pptx | Python | 13.4ms | 90.0% | MIT |
| markitdown | Python | 74.1ms | 90.2% | MIT |

### Legacy formats

| Library | Language | Mean | Pass Rate | License |
|---------|----------|------|-----------|---------|
| **office_oxide** | **Rust** | **4.1ms** | **90.0%** | **MIT** |
| catdoc | C | 5.6ms | 95.5% | GPL-2.0 |
| antiword | C | 5.6ms | 79.5% | GPL-2.0 |

| Library | Language | Mean | Pass Rate | License |
|---------|----------|------|-----------|---------|
| **office_oxide** | **Rust** | **5.1ms** | **95.3%** | **MIT** |
| calamine (crate) | Rust | 9.0ms | 93.0% | MIT |
| xlrd | Python | 26.9ms | 92.5% | BSD-3 |
| xls2csv (catdoc) | C | 36.5ms | 97.7% | GPL-2.0 |
| python-calamine | Rust (Python) | 44.5ms | 92.6% | MIT |

| Library | Language | Mean | Pass Rate | License |
|---------|----------|------|-----------|---------|
| **office_oxide** | **Rust** | **6.7ms** | **92.3%** | **MIT** |
| catppt (catdoc) | C | 80.2ms | 99.8% | GPL-2.0 |

"Pass" means the library opened the file and returned without an error. The command-line tools exit 0 on many files they cannot read and print little or nothing, so their rate overstates what they extracted. The collection also holds deliberately damaged and password-protected files, which is where most of office_oxide's failures are. The earlier April 2026 run (a different 6,062-file collection) is kept in [BENCHMARKS.md](BENCHMARKS.md).

## Supported Formats

| Format | Extension | Read | Write | Edit | Convert | Text | Markdown | HTML | IR |
|--------|-----------|------|-------|------|---------|------|----------|------|----|
| Word (OOXML) | .docx | Yes | Yes | Yes | — | Yes | Yes | Yes | Yes |
| Excel (OOXML) | .xlsx | Yes | Yes | Yes | — | Yes | Yes | Yes | Yes |
| PowerPoint (OOXML) | .pptx | Yes | Yes | Yes | — | Yes | Yes | Yes | Yes |
| Word (Legacy) | .doc | Yes | — | — | → .docx | Yes | Yes | Yes | Yes |
| Excel (Legacy) | .xls | Yes | — | — | → .xlsx | Yes | Yes | Yes | Yes |
| PowerPoint (Legacy) | .ppt | Yes | — | — | → .pptx | Yes | Yes | Yes | Yes |

Legacy formats can be converted to modern OOXML with `save_as()`:

```python
doc = Document.open("report.doc")
doc.save_as("report.docx")  # Converts DOC → DOCX
```

## Python API

```python
from office_oxide import Document, extract_text, to_markdown, to_html

# Quick extraction
text = extract_text("report.docx")
markdown = to_markdown("spreadsheet.xlsx")
html = to_html("slides.pptx")

# Document object
doc = Document.open("presentation.pptx")
text = doc.plain_text()
md = doc.to_markdown()
html = doc.to_html()
ir = doc.to_ir()  # Structured JSON intermediate representation
fmt = doc.format_name()  # "pptx"

# From bytes
with open("file.docx", "rb") as f:
    doc = Document.from_bytes(f.read(), "docx")
```

All 6 formats supported. Works with `str`, `pathlib.Path`, or raw bytes.

## Rust API

```rust
use office_oxide::{Document, DocumentFormat};

// Open from path (format auto-detected from extension)
let doc = Document::open("report.docx")?;

// Open from reader with explicit format
let file = std::fs::File::open("data.xlsx")?;
let doc = Document::from_reader(file, DocumentFormat::Xlsx)?;

// Extract content
let text = doc.plain_text();
let markdown = doc.to_markdown();
let html = doc.to_html();
let ir = doc.to_ir(); // Format-agnostic DocumentIR

// Access format-specific types
if let Some(docx) = doc.as_docx() {
    println!("Paragraphs: {}", docx.body.elements.len());
}

// Create documents from IR
use office_oxide::create::create_from_ir;
create_from_ir(&ir, DocumentFormat::Docx, "output.docx")?;
```

### Sub-modules

Each format is available as a sub-module for direct access:

```rust
use office_oxide::docx::DocxDocument;
use office_oxide::xlsx::XlsxDocument;
use office_oxide::pptx::PptxDocument;
use office_oxide::doc::DocDocument;
use office_oxide::xls::XlsDocument;
use office_oxide::ppt::PptDocument;
```

### Limits for untrusted input

A spreadsheet can reference one shared string from every cell, so its
rendered text can be the string times the cell count — gigabytes from a
few hundred kilobytes. Every renderer stops at a per-document text budget
(256 Mi characters by default) and appends a visible `[output truncated: …]`
notice; the `.xls` reader also flags `Metadata::text_truncated`. A process
that reads larger workbooks whole raises the limit once at startup:

```rust
office_oxide::limits::set_max_text_chars(1 << 30);
```

## Installation

### Python

```bash
pip install office-oxide
```

Wheels available for Linux, macOS, and Windows. Python 3.10–3.14.

### Rust

```toml
[dependencies]
office_oxide = "0.1.16"
```

### JavaScript/WASM

```bash
npm install office-oxide-wasm    # portable WASM (browser + Node.js)
npm install office-oxide         # native addon via koffi (Node.js only, no node-gyp)
```

### Go

```bash
go get github.com/yfedoseev/office_oxide/go
```

See [go/README.md](go/README.md) for setup details.

### C# / .NET

```bash
dotnet add package OfficeOxide
```

See [csharp/OfficeOxide/README.md](csharp/OfficeOxide/README.md) for setup details.

### CLI

```bash
cargo install office_oxide_cli       # Cargo
cargo binstall office_oxide_cli      # Pre-built binary
```

### MCP Server

```bash
cargo install office_oxide_mcp
```

## CLI

`office-oxide` provides fast Office document processing from your terminal:

```bash
office-oxide text report.docx                    # Extract plain text
office-oxide markdown data.xlsx                  # Convert to Markdown
office-oxide markdown report.docx --embed-images  # Markdown with inline base64 images
office-oxide html slides.pptx                    # Convert to HTML
office-oxide ir document.docx                    # Dump IR as JSON
office-oxide info report.docx                    # Show format and metadata
office-oxide replace report.docx OLD NEW         # Find/replace text (DOCX/PPTX), in place
office-oxide replace report.docx OLD NEW --output edited.docx   # …written to a new file
office-oxide text report.docx --include-hidden --include-deleted  # also hidden text and tracked deletions
```

All six formats are supported for reading (docx, xlsx, pptx, doc, xls, ppt); `replace` writes
DOCX and PPTX. Use `--help` for all options.

## MCP Server

`office-oxide-mcp` lets AI assistants (Claude, Cursor, etc.) read Office documents locally via the [Model Context Protocol](https://modelcontextprotocol.io/).

Add to your MCP client configuration:

```json
{
  "mcpServers": {
    "office-oxide": { "command": "office-oxide-mcp" }
  }
}
```

The server exposes `extract`, `replace_text`, and `info` tools. All processing runs locally — no files leave your machine.

## Building from Source

```bash
git clone https://github.com/yfedoseev/office_oxide
cd office_oxide
cargo build --release
cargo test

# Python bindings
maturin develop --features python

# Shared library for Go, JS/TS (koffi), and C# bindings
cargo build --release --lib
# Output: target/release/liboffice_oxide.{so,dylib} or office_oxide.dll
```

## Documentation

- **[Documentation Site](https://office.oxide.fyi)** — Guides and examples
- **[Getting Started (Rust)](docs/getting-started-rust.md)** — Rust guide
- **[Getting Started (Python)](docs/getting-started-python.md)** — Python guide
- **[Getting Started (Go)](go/README.md)** — Go guide
- **[Getting Started (JavaScript / TypeScript)](js/README.md)** — Node.js native guide
- **[Getting Started (WASM)](wasm-pkg/README.md)** — Browser and Node.js WASM guide
- **[Getting Started (C# / .NET)](csharp/OfficeOxide/README.md)** — .NET guide
- **[Getting Started (C FFI)](docs/getting-started-c.md)** — C FFI guide
- **[API Docs (Rust)](https://docs.rs/office_oxide)** — Full Rust API reference
- **[Architecture](docs/ARCHITECTURE.md)** — System design and module structure

## Use Cases

- **RAG / LLM pipelines** — Extract clean text or Markdown from Office documents for retrieval-augmented generation
- **Document processing at scale** — Parse thousands of documents in seconds
- **Data extraction** — Pull structured data from spreadsheets, tables, and presentations
- **Format conversion** — Convert between formats via the intermediate representation
- **python-docx / openpyxl alternative** — 5–10× faster, supports all 6 formats in one library

## Why I built this

I needed a library that could read all six Office formats at once — not six separate packages — and I needed it without pulling in a JVM, a Python runtime, or a GPL-licensed dependency. Nothing existed that combined speed, correctness, and a permissive license across the full DOCX / XLSX / PPTX / DOC / XLS / PPT surface, so I wrote it in Rust and wrapped it for every language I use day-to-day. The same binary powers Python via PyO3, Node.js via koffi, Go via cgo, C# via P/Invoke, and the browser via WASM — one fix lands everywhere.

If it saves you a dependency, a license audit, or a weekend, consider leaving a star. If something's broken or missing, [open an issue](https://github.com/yfedoseev/office_oxide/issues) — I read all of them.

— Yury

## License & trademark

The **code** is dual-licensed under [MIT](LICENSE-MIT) or [Apache-2.0](LICENSE-APACHE) at your option. No AGPL, no GPL, no copyleft restrictions. Use freely in commercial and open-source projects.

The **name and brand** are not covered by the code license: **"OfficeOxide"** (the product name) and **"office_oxide"** (the package name), together with any associated logo, are trademarks of Yury Fedoseev. You may say your product "uses OfficeOxide"; please don't name a different or modified product "OfficeOxide" or "office_oxide". See [TRADEMARKS.md](TRADEMARKS.md).

© 2026 Yury Fedoseev and the OfficeOxide contributors. Contributions are accepted under the project's [DCO](CONTRIBUTING.md#developer-certificate-of-origin-dco) and [CLA](CLA.md); contributors retain ownership of their work.

## Contributing

We welcome contributions! See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

```bash
cargo build && cargo test && cargo fmt && cargo clippy -- -D warnings
```

## Citation

```bibtex
@software{office_oxide,
  title = {Office Oxide: Fast Office Document Processing for Rust, Python, Go, JavaScript, and C#},
  author = {Yury Fedoseev},
  year = {2026},
  url = {https://github.com/yfedoseev/office_oxide}
}
```

---

**Rust** + **Python** + **Go** + **JS/TS** + **C#** + **WASM** + **CLI** + **MCP** | MIT/Apache-2.0 | Opens 97.3% of a 10,386-file test collection | 5–10× faster than the common Python libraries | 6 formats
