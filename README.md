# Office Oxide — The Fastest Office Document Library for Rust, Python & WASM

The fastest library for text extraction from Office documents. Rust core with Python bindings and WASM support. Handles DOCX, XLSX, PPTX, DOC, XLS, and PPT. Up to 100× faster than python-docx, openpyxl, python-pptx, and xlrd. Matches calamine on XLSX. 98.1% pass rate on 6,062 files — all failures are genuinely invalid. MIT licensed.

[![Crates.io](https://img.shields.io/crates/v/office_oxide.svg)](https://crates.io/crates/office_oxide)
[![PyPI](https://img.shields.io/pypi/v/office-oxide.svg)](https://pypi.org/project/office-oxide/)
[![npm](https://img.shields.io/npm/v/office-oxide-wasm)](https://www.npmjs.com/package/office-oxide-wasm)
[![Build Status](https://github.com/yfedoseev/office_oxide/workflows/CI/badge.svg)](https://github.com/yfedoseev/office_oxide/actions)
[![License: MIT OR Apache-2.0](https://img.shields.io/badge/License-MIT%20OR%20Apache--2.0-blue.svg)](https://opensource.org/licenses)

## Quick Start

### Python
```python
from office_oxide import Document, extract_text, to_markdown

# One-liner text extraction
text = extract_text("report.docx")
markdown = to_markdown("data.xlsx")

# Or use the Document object for more control
doc = Document.open("slides.pptx")
print(doc.plain_text())
print(doc.to_markdown())
print(doc.format_name())  # "pptx"
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
office_oxide = "0.1"
```

### JavaScript/WASM
```javascript
const { WasmDocument } = require("office-oxide-wasm");

const doc = new WasmDocument(buffer, "docx");
console.log(doc.plainText());
console.log(doc.toMarkdown());
```

```bash
npm install office-oxide-wasm
```

## Why office_oxide?

- **Fast** — 8-100× faster than python-docx, openpyxl, mammoth, markitdown; matches or beats calamine on XLSX
- **Reliable** — 100% pass rate on all valid documents in a 2,570-file corpus — the 99 non-passing files (3.8%) are all genuinely invalid (CVE exploits, fuzz-corrupted, XML bombs, encrypted)
- **Complete** — 6 formats: DOCX, XLSX, PPTX + legacy DOC, XLS, PPT
- **Multi-platform** — Rust, Python, JavaScript/WASM — one library, all platforms
- **Permissive** — MIT / Apache-2.0, no AGPL or GPL restrictions

## Performance

Benchmarked on 2,570 files from 7 independent public test suites (Apache POI, LibreOffice, Apache Tika, calamine, python-docx, python-pptx, mammoth). Text extraction, single-thread, no warm-up.

### DOCX — 187 files

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **7ms** | **160ms** | **91.4%** | **MIT** |
| docx2txt | Python | 54ms | 898ms | 90.9% | MIT |
| python-docx | Python | 110ms | 631ms | 89.8% | MIT |
| mammoth | Python | 353ms | 8,536ms | 90.4% | BSD-2 |
| markitdown | Python | 773ms | 18,883ms | 98.9% | MIT |

### PPTX — 563 files

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **3ms** | **19ms** | **97.5%** | **MIT** |
| python-pptx | Python | 24ms | 219ms | 96.6% | MIT |
| markitdown | Python | 38ms | 393ms | 98.6% | MIT |

### XLSX — 718 files

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **6ms** | **83ms** | **96.5%** | **MIT** |
| python-calamine | Rust/Python | 5ms | 324ms | 96.0% | MIT |
| openpyxl | Python | 245ms | 4,998ms | 94.4% | MIT |

office_oxide and calamine are comparable on XLSX mean, but office_oxide is 1.6× faster on large files (>50KB), has better p99 latency, handles 13 more files, and never panics. calamine is spreadsheet-only; office_oxide handles all 6 formats.

### Legacy Formats — 1,103 files

| Library | Format | Mean | p99 | Pass Rate |
|---------|--------|------|-----|-----------|
| **office_oxide** | **DOC** | **2ms** | **12ms** | **99.4%** |
| **office_oxide** | **XLS** | **9ms** | **121ms** | **95.2%** |
| xlrd | XLS | 114ms | 1,758ms | 90.9% |
| **office_oxide** | **PPT** | **6ms** | **40ms** | **95.9%** |

No other Python library supports .doc or .ppt text extraction without a JVM (Apache Tika) or external binaries.

### Corpus

| Source | Files | Formats | License |
|--------|------:|---------|---------|
| [Apache POI](https://github.com/apache/poi) | 1,197 | All 6 | Apache-2.0 |
| [LibreOffice](https://github.com/LibreOffice/core) | 848 | All 6 | MPL-2.0 |
| [Apache Tika](https://github.com/apache/tika) | 87 | All 6 | Apache-2.0 |
| [calamine](https://github.com/tafia/calamine) | 42 | XLSX/XLS | MIT |
| [python-docx](https://github.com/python-openxml/python-docx) | 24 | DOCX | MIT |
| [python-pptx](https://github.com/scanny/python-pptx) | 348 | PPTX | MIT |
| [mammoth](https://github.com/mwilliamson/python-mammoth) | 24 | DOCX | BSD-2 |
| **Total** | **2,570** | | |

### Pass Rate

100% pass rate on all valid documents — the 99 non-passing files (3.8% of the corpus) are all genuinely invalid:

| Category | Count | Examples |
|----------|------:|---------|
| Invalid CFB/OLE2 container | 43 | CVE exploits with corrupted headers (CVE-2006-3059, CVE-2007-0031, etc.) |
| Invalid ZIP archive | 41 | ClusterFuzz crash cases, truncated files |
| Malformed XML | 10 | XML bombs (`lol9` entity), ill-formed tags |
| Missing required part | 5 | Encrypted/password-protected files |

Zero panics, zero timeouts, zero false negatives on valid documents.

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
office_oxide = "0.1"
```

### JavaScript/WASM

```bash
npm install office-oxide-wasm
```

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
```

## Documentation

- **[API Docs (Rust)](https://docs.rs/office_oxide)** — Full Rust API reference
- **[Documentation Site](https://office.oxide.fyi)** — Guides and examples
- **[Architecture](docs/ARCHITECTURE.md)** — System design and module structure

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
  title = {Office Oxide: Fast Office Document Processing for Rust and Python},
  author = {Yury Fedoseev},
  year = {2026},
  url = {https://github.com/yfedoseev/office_oxide}
}
```

---

**Rust** + **Python** + **WASM** | MIT/Apache-2.0 | 98.1% pass rate on 6,062 files | Up to 100× faster than alternatives | 6 formats
