# office_oxide for Go — The Fastest Office Document Library for Go

Fast Office document processing (DOCX, XLSX, PPTX, DOC, XLS, PPT) for Go, powered by a Rust core.
Up to 100× faster than python-docx, openpyxl, and python-pptx on equivalent tasks.
100% pass rate on valid Office files — zero failures on legitimate Word/Excel/PowerPoint documents. MIT / Apache-2.0 licensed.

[![Go Reference](https://pkg.go.dev/badge/github.com/yfedoseev/office_oxide/go.svg)](https://pkg.go.dev/github.com/yfedoseev/office_oxide/go)
[![License: MIT OR Apache-2.0](https://img.shields.io/badge/License-MIT%20OR%20Apache--2.0-blue.svg)](https://opensource.org/licenses)

> **Part of the [office_oxide](https://github.com/yfedoseev/office_oxide) toolkit.** Same Rust core, same pass rate as the
> [Rust](https://docs.rs/office_oxide), [Python](../python/README.md),
> [JavaScript (native)](../js/README.md), [C# / .NET](../csharp/OfficeOxide/README.md),
> and [WASM](../wasm-pkg/README.md) bindings.

## Quick Start

```bash
go get github.com/yfedoseev/office_oxide/go
```

The binding uses cgo and links against the `office_oxide` Rust static library.
You must either:

1. **Build locally from the monorepo.** From the office_oxide repo root:

   ```bash
   cargo build --release --lib
   go test -tags office_oxide_dev ./go/...
   ```

   The `office_oxide_dev` build tag points cgo at `target/release/`.

2. **Point cgo at an installed prefix.** Build the library once, place
   `liboffice_oxide.a` (or `.so` / `.dylib`) somewhere on your link path,
   and the header somewhere on your include path, then set:

   ```bash
   export CGO_CFLAGS="-I/path/to/include"
   export CGO_LDFLAGS="-L/path/to/lib -loffice_oxide"
   ```

```go
package main

import (
    "fmt"
    "log"

    oo "github.com/yfedoseev/office_oxide/go"
)

func main() {
    doc, err := oo.Open("report.docx")
    if err != nil {
        log.Fatal(err)
    }
    defer doc.Close()

    fmt.Println("format:", mustStr(doc.Format()))
    fmt.Println(mustStr(doc.PlainText()))

    md, err := doc.ToMarkdown()
    if err != nil {
        log.Fatal(err)
    }
    fmt.Println(md)
}

func mustStr(s string, err error) string {
    if err != nil {
        log.Fatal(err)
    }
    return s
}
```

## Why office_oxide?

- **Fast** — 0.8ms mean DOCX, 5.0ms mean XLSX, 0.7ms mean PPTX; up to 100× faster than python-docx / openpyxl / python-pptx
- **Reliable** — 100% pass rate on valid Office files (6,062-file corpus); zero failures on legitimate documents
- **Complete** — 6 formats: DOCX, XLSX, PPTX + legacy DOC, XLS, PPT
- **Permissive** — MIT / Apache-2.0, no AGPL or GPL restrictions
- **Idiomatic Go** — Sentinel errors, `defer doc.Close()`, clean return signatures

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

Full methodology and corpus breakdown in [BENCHMARKS.md](../BENCHMARKS.md).

## Editing

```go
ed, err := oo.OpenEditable("template.docx")
if err != nil { log.Fatal(err) }
defer ed.Close()

if _, err := ed.ReplaceText("{{NAME}}", "Alice"); err != nil {
    log.Fatal(err)
}
if err := ed.Save("out.docx"); err != nil {
    log.Fatal(err)
}
```

## API

| Function / method | Description |
| --- | --- |
| `Version() string` | Library version. |
| `DetectFormat(path) string` | Short format name (`"docx"`, …) or `""`. |
| `Open(path) (*Document, error)` | Open from a file path. |
| `OpenFromBytes(data, format) (*Document, error)` | Open from memory. |
| `Document.Close() error` | Release the handle (safe to defer). |
| `Document.Format() (string, error)` | Detected format. |
| `Document.PlainText() (string, error)` | Extract plain text. |
| `Document.ToMarkdown() (string, error)` | Convert to Markdown. |
| `Document.ToHTML() (string, error)` | Convert to an HTML fragment. |
| `Document.ToIRJSON() (string, error)` | Format-agnostic IR as JSON. |
| `Document.SaveAs(path) error` | Save/convert to a new format. |
| `OpenEditable(path) (*EditableDocument, error)` | Open for editing (DOCX/XLSX/PPTX). |
| `EditableDocument.ReplaceText(find, replace) (int64, error)` | In-place replace. |
| `EditableDocument.SetCell(sheet, ref, value) error` | Write an XLSX cell. |
| `EditableDocument.Save(path) error` | Persist to disk. |
| `EditableDocument.SaveToBytes() ([]byte, error)` | Persist to memory. |
| `ExtractText(path)` / `ToMarkdown(path)` / `ToHTML(path)` | One-shot helpers. |

Errors returned from FFI are `*Error` values that expose the underlying
code via `Error.Code` and the originating operation via `Error.Op`.

## Other languages

office_oxide ships the same Rust core through six bindings:

- **Rust** — `cargo add office_oxide` — see [docs.rs/office_oxide](https://docs.rs/office_oxide)
- **Python** — `pip install office-oxide` — see [python/README.md](../python/README.md)
- **JavaScript (native)** — `npm install office-oxide` — see [js/README.md](../js/README.md)
- **C# / .NET** — `dotnet add package OfficeOxide` — see [csharp/OfficeOxide/README.md](../csharp/OfficeOxide/README.md)
- **WASM** — `npm install office-oxide-wasm` — see [wasm-pkg/README.md](../wasm-pkg/README.md)

A bug fix in the Rust core lands in every binding on the next release.

## Why I built this

I needed a library that could read all six Office formats at once — not six separate packages — and I needed it without pulling in a JVM, a Python runtime, or a GPL-licensed dependency. Nothing existed that combined speed, correctness, and a permissive license across the full DOCX / XLSX / PPTX / DOC / XLS / PPT surface, so I wrote it in Rust and wrapped it for every language I use day-to-day.

If it saves you a dependency or a license audit, a star on GitHub genuinely helps. If something's broken or missing, [open an issue](https://github.com/yfedoseev/office_oxide/issues).

— Yury

## License

Dual-licensed under [MIT](../LICENSE-MIT) or [Apache-2.0](../LICENSE-APACHE) at your option.

---

**Go** + **Rust core** | MIT / Apache-2.0 | 100% pass rate on valid Office files (6,062-file corpus) | Up to 100× faster than alternatives | 6 formats
