# Office Oxide — The Fastest Native Office Document Library

A fast, memory-safe library for text extraction from Office documents. Rust core with **first-class bindings for Python, Go, C#/.NET, Node.js (native and WASM), and a stable C FFI**. Handles DOCX, XLSX, PPTX, DOC, XLS, and PPT. Up to 100× faster than python-docx, openpyxl, python-pptx, and xlrd. Beats python-calamine on XLSX. **98.4% pass rate on 6,062 files** — zero failures on legitimate Office documents. MIT/Apache-2.0 dual-licensed.

> **Scope of "fastest".** Benchmarks compare Office Oxide against other
> native / embeddable libraries (no JVM runtime required): python-docx,
> openpyxl, python-pptx, python-calamine, xlrd, markitdown, catdoc,
> antiword, xls2csv, calamine (Rust), dotext, docx-rs. Apache POI and
> Apache Tika are out of scope for this comparison because they require
> a JVM and are targeted at a different deployment shape. POI/Tika
> numbers may be added in a future release.

[![Crates.io](https://img.shields.io/crates/v/office_oxide.svg)](https://crates.io/crates/office_oxide)
[![PyPI](https://img.shields.io/pypi/v/office-oxide.svg)](https://pypi.org/project/office-oxide/)
[![npm (wasm)](https://img.shields.io/npm/v/office-oxide-wasm?label=npm%20wasm)](https://www.npmjs.com/package/office-oxide-wasm)
[![npm (native)](https://img.shields.io/npm/v/office-oxide?label=npm%20native)](https://www.npmjs.com/package/office-oxide)
[![NuGet](https://img.shields.io/nuget/v/OfficeOxide)](https://www.nuget.org/packages/OfficeOxide)
[![Go Reference](https://pkg.go.dev/badge/github.com/yfedoseev/office_oxide/go.svg)](https://pkg.go.dev/github.com/yfedoseev/office_oxide/go)
[![Build Status](https://github.com/yfedoseev/office_oxide/workflows/CI/badge.svg)](https://github.com/yfedoseev/office_oxide/actions)
[![License: MIT OR Apache-2.0](https://img.shields.io/badge/License-MIT%20OR%20Apache--2.0-blue.svg)](https://opensource.org/licenses)
[![OpenSSF Scorecard](https://api.scorecard.dev/projects/github.com/yfedoseev/office_oxide/badge)](https://scorecard.dev/viewer/?uri=github.com/yfedoseev/office_oxide)

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
office_oxide = "0.1.0"
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

- **Fast** — 8-100× faster than python-docx, openpyxl, python-pptx, xlrd; beats calamine on XLSX
- **Reliable** — 98.4% pass rate on a 6,062-file corpus. The 97 non-passing files are all invalid archives, non-Office inputs (WordPerfect, pre-OLE2 Excel 3/4, empty, mislabeled), or fuzz/CVE fixtures. Zero failures on legitimate Word 97+ / Excel 97+ / PowerPoint 97+ files
- **Complete** — 6 formats: DOCX, XLSX, PPTX + legacy DOC, XLS, PPT
- **Multi-platform** — Rust, Python, JavaScript/WASM — one library, all platforms
- **Permissive** — MIT / Apache-2.0, no AGPL or GPL restrictions

## Performance

Benchmarked on **6,062 files** from 11 independent public test suites. Single-thread, release build with LTO, warm disk cache (steady-state), median of three runs on an idle system. Full methodology in [BENCHMARKS.md](BENCHMARKS.md).

### DOCX — 2,538 files

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **0.8ms** | **3.9ms** | **98.9%** | **MIT** |
| python-docx | Python | 11.8ms | 98ms | 95.1% | MIT |

### XLSX — 1,802 files

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **5.0ms** | **40ms** | **97.8%** | **MIT** |
| python-calamine | Rust/Python | 13.9ms | 183ms | 96.6% | MIT |
| openpyxl | Python | 94.5ms | 698ms | 96.2% | MIT |

### PPTX — 806 files

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **0.7ms** | **3.9ms** | **98.4%** | **MIT** |
| python-pptx | Python | 32.5ms | 174ms | 86.7% | MIT |

### Legacy Formats — 916 files

| Library | Format | Mean | p99 | Pass Rate | License |
|---------|--------|------|-----|-----------|---------|
| **office_oxide** | **DOC** (246) | **0.3ms** | **3.4ms** | **94.7%** | **MIT** |
| catdoc | DOC | 4.3ms | 41ms | 90.2% | GPL-2.0 |
| antiword | DOC | 4.5ms | 66ms | 76.8% | GPL-2.0 |
| **office_oxide** | **XLS** (494) | **2.8ms** | 75ms | **99.2%** | **MIT** |
| xls2csv (catdoc) | XLS | 6.9ms | **58ms** | 84.0% | GPL-2.0 |
| python-calamine | XLS | 9.0ms | 96ms | 90.7% | MIT |
| xlrd | XLS | 36.6ms | 503ms | 93.1% | BSD-3 |
| **office_oxide** | **PPT** (176) | **0.7ms** | **6.6ms** | **100%** | **MIT** |
| catppt (catdoc) | PPT | 2.8ms | 8ms | 77.8% | GPL-2.0 |

On .xls, xls2csv has a tighter p99 (58ms vs 75ms) because it emits truncated/lossy output on complex sheets. office_oxide is 2.4× faster on the mean and passes 15pp more of the corpus. No other Rust or Python library supports .doc, .xls, and .ppt text extraction without a JVM (Apache Tika) or external binaries.

### Corpus

| Source | Files | License |
|--------|------:|---------|
| [LibreOffice Core](https://github.com/LibreOffice/core) | 2,185 | MPL-2.0 |
| [Apache POI](https://github.com/apache/poi) | 1,298 | Apache-2.0 |
| [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) | 707 | MIT |
| [ClosedXML](https://github.com/ClosedXML/ClosedXML) | 371 | MIT |
| [Pandoc](https://github.com/jgm/pandoc) | 224 | GPL-2.0 |
| [python-docx](https://github.com/python-openxml/python-docx) + [python-pptx](https://github.com/scanny/python-pptx) | 111 | MIT |
| [Apache Tika](https://github.com/apache/tika) | 108 | Apache-2.0 |
| [calamine](https://github.com/tafia/calamine) | 28 | MIT |
| [openpreserve](https://github.com/openpreserve/format-corpus) | 20 | CC0 |
| [oletools](https://github.com/decalage2/oletools) | 17 | BSD-2 |
| LibreOffice (legacy) | 12 | MPL-2.0 |
| **Total** | **6,062** | |

### Pass Rate — 98.4% (5,965 / 6,062)

All 97 non-passing files are invalid inputs:

| Category | Count | Notes |
|----------|------:|-------|
| Invalid ZIP / CFB archive | 43 | Truncated, missing EOCD, bad CFB magic |
| Missing required part | 21 | Encrypted, password-protected, or stream absent |
| Malformed XML | 18 | XML bombs, ill-formed tags, fuzz-corrupted content |
| Invalid CFB header | 15 | WordPerfect / IBM DisplayWrite / Excel 3/4 misnamed as .doc/.xls, CVE-exploit fixtures |

**Zero failures on legitimate Word 97+ / Excel 97+ / PowerPoint 97+ files.** Zero panics, zero timeouts, zero false negatives on valid documents. Full breakdown in [BENCHMARKS.md](BENCHMARKS.md).

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

## Installation

### Python

```bash
pip install office-oxide
```

Wheels available for Linux, macOS, and Windows. Python 3.8–3.14.

### Rust

```toml
[dependencies]
office_oxide = "0.1.0"
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

### MCP Server (for AI assistants)

Give Claude, Cursor, or any MCP-compatible tool the ability to read Office documents:

```bash
cargo install office_oxide_mcp
```

Add to Claude Desktop `claude_desktop_config.json`:
```json
{
  "mcpServers": {
    "office-oxide": { "command": "office-oxide-mcp" }
  }
}
```

### CLI

```bash
cargo install office_oxide_cli
office-oxide text report.docx
office-oxide markdown data.xlsx
office-oxide html slides.pptx
office-oxide ir document.docx
```

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

- **[API Docs (Rust)](https://docs.rs/office_oxide)** — Full Rust API reference
- **[Documentation Site](https://office.oxide.fyi)** — Guides and examples
- **[Architecture](docs/ARCHITECTURE.md)** — System design and module structure

## Why I built this

I needed a library that could read all six Office formats at once — not six separate packages — and I needed it without pulling in a JVM, a Python runtime, or a GPL-licensed dependency. Nothing existed that combined speed, correctness, and a permissive license across the full DOCX / XLSX / PPTX / DOC / XLS / PPT surface, so I wrote it in Rust and wrapped it for every language I use day-to-day. The same binary powers Python via PyO3, Node.js via koffi, Go via cgo, C# via P/Invoke, and the browser via WASM — one fix lands everywhere.

If it saves you a dependency, a license audit, or a weekend, consider leaving a star. If something's broken or missing, [open an issue](https://github.com/yfedoseev/office_oxide/issues) — I read all of them.

— Yury

## Use Cases

- **RAG / LLM pipelines** — Extract clean text or Markdown from Office documents for retrieval-augmented generation
- **Document processing at scale** — Parse thousands of documents in seconds
- **Data extraction** — Pull structured data from spreadsheets, tables, and presentations
- **Format conversion** — Convert between formats via the intermediate representation
- **python-docx / openpyxl alternative** — Up to 100× faster, supports all 6 formats in one library

## License

Dual-licensed under [MIT](LICENSE-MIT) or [Apache-2.0](LICENSE-APACHE) at your option. No AGPL, no GPL, no copyleft restrictions. Use freely in commercial and open-source projects.

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

**Rust** + **Python** + **Go** + **JS/TS** + **C#** + **WASM** + **CLI** + **MCP** | MIT/Apache-2.0 | 98.4% pass rate on 6,062 files (zero failures on legitimate Office docs) | Up to 100× faster than alternatives | 6 formats
