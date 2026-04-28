# office-oxide for Python — The Fastest Office Document Library for Python

The fastest Python library for text extraction, Markdown conversion, and editing across all six Microsoft Office formats. Powered by a Rust core via PyO3. Up to 100× faster than python-docx, openpyxl, and python-pptx. 100% pass rate on valid Office files — zero failures on legitimate Word/Excel/PowerPoint documents. MIT / Apache-2.0 licensed.

[![PyPI](https://img.shields.io/pypi/v/office-oxide.svg)](https://pypi.org/project/office-oxide/)
[![License: MIT OR Apache-2.0](https://img.shields.io/badge/License-MIT%20OR%20Apache--2.0-blue.svg)](https://opensource.org/licenses)

> **Part of the [office_oxide](https://github.com/yfedoseev/office_oxide) toolkit.** Same Rust core, same pass rate as the
> [Rust](https://docs.rs/office_oxide), [Go](../go/README.md),
> [JavaScript (native)](../js/README.md), [C# / .NET](../csharp/OfficeOxide/README.md),
> and [WASM](../wasm-pkg/README.md) bindings.

## Quick Start

```bash
pip install office-oxide
```

```python
from office_oxide import Document, EditableDocument, extract_text, to_markdown

# One-shot helpers
text = extract_text("report.docx")
md = to_markdown("spreadsheet.xlsx")

# Context-managed document
with Document.open("slides.pptx") as doc:
    print(doc.format)          # "pptx"
    print(doc.plain_text())
    print(doc.to_markdown())
    ir = doc.to_ir()           # format-agnostic IR as nested dicts
```

## Why office-oxide?

- **Fast** — 0.8ms mean DOCX, 5.0ms mean XLSX, 0.7ms mean PPTX; up to 100× faster than python-docx / openpyxl / python-pptx
- **Reliable** — 100% pass rate on valid Office files (6,062-file corpus); zero failures on legitimate documents
- **Complete** — 6 formats: DOCX, XLSX, PPTX + legacy DOC, XLS, PPT — one library
- **Permissive** — MIT / Apache-2.0, unlike many alternatives that require GPL or AGPL
- **Drop-in** — `extract_text()` and `to_markdown()` replace python-docx / openpyxl in one line

## Performance

Benchmarked on 6,062 files from 11 independent public test suites. Single-thread, release build, warm disk cache.

| Library | Format | Mean | Pass Rate | License |
|---------|--------|------|-----------|---------|
| **office_oxide** | **DOCX** | **0.8ms** | **98.9%** | **MIT** |
| python-docx | DOCX | 11.8ms | 95.1% | MIT |
| **office_oxide** | **XLSX** | **5.0ms** | **97.8%** | **MIT** |
| python-calamine | XLSX | 13.9ms | 96.6% | MIT |
| openpyxl | XLSX | 94.5ms | 96.2% | MIT |
| **office_oxide** | **PPTX** | **0.7ms** | **98.4%** | **MIT** |
| python-pptx | PPTX | 32.5ms | 86.7% | MIT |

Full methodology and corpus breakdown in [BENCHMARKS.md](../BENCHMARKS.md).

## Installation

```bash
pip install office-oxide
```

Pre-built wheels for Linux (x86_64, aarch64, musl), macOS (x86_64, arm64), and Windows (x86_64). Python 3.8–3.14. No system dependencies, no Rust toolchain required.

## API

### Extraction

```python
from office_oxide import Document, extract_text, to_markdown, to_html

# One-shot helpers
text = extract_text("report.docx")
md   = to_markdown("spreadsheet.xlsx")
html = to_html("slides.pptx")

# Document object
doc = Document.open("presentation.pptx")
text = doc.plain_text()
md   = doc.to_markdown()
html = doc.to_html()
ir   = doc.to_ir()           # structured dict (format-agnostic IR)
fmt  = doc.format            # "pptx"
```

`Document.open` accepts both `str` and `pathlib.Path`. Can also be used as a context manager.

### From bytes

```python
with open("file.docx", "rb") as f:
    doc = Document.from_bytes(f.read(), "docx")
```

### Editing

```python
with EditableDocument.open("template.docx") as ed:
    ed.replace_text("{{NAME}}", "Alice")
    ed.save("out.docx")
```

### Spreadsheet cells

```python
with EditableDocument.open("data.xlsx") as ed:
    ed.set_cell(0, "A1", "Revenue")
    ed.set_cell(0, "B1", 12345.67)
    ed.set_cell(0, "C1", True)
    ed.save("data.edited.xlsx")
```

## Other languages

office_oxide ships the same Rust core through six bindings:

- **Rust** — `cargo add office_oxide` — see [docs.rs/office_oxide](https://docs.rs/office_oxide)
- **Go** — `go get github.com/yfedoseev/office_oxide/go` — see [go/README.md](../go/README.md)
- **JavaScript (native)** — `npm install office-oxide` — see [js/README.md](../js/README.md)
- **C# / .NET** — `dotnet add package OfficeOxide` — see [csharp/OfficeOxide/README.md](../csharp/OfficeOxide/README.md)
- **WASM** — `npm install office-oxide-wasm` — see [wasm-pkg/README.md](../wasm-pkg/README.md)

## Why I built this

I needed a library that could read all six Office formats at once — not six separate packages — and I needed it without pulling in a JVM or a GPL-licensed dependency. Nothing existed that combined speed, correctness, and a permissive license across the full DOCX / XLSX / PPTX / DOC / XLS / PPT surface, so I wrote it in Rust and wrapped it for every language I use day-to-day.

If it saves you a dependency or a license audit, a star on GitHub genuinely helps. If something's broken or missing, [open an issue](https://github.com/yfedoseev/office_oxide/issues).

— Yury

## License

Dual-licensed under [MIT](../LICENSE-MIT) or [Apache-2.0](../LICENSE-APACHE) at your option.

---

**Python** + **Rust core** | MIT / Apache-2.0 | 100% pass rate on valid Office files (6,062-file corpus) | Up to 100× faster than alternatives | 6 formats
