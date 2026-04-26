# OfficeOxide for .NET — The Fastest Office Document Library for C#

Idiomatic C# bindings for [office_oxide](https://github.com/yfedoseev/office_oxide) — a fast Rust library for parsing, converting, and editing Microsoft Office documents (DOCX, XLSX, PPTX, DOC, XLS, PPT).

- `IDisposable` / `using` pattern for native handles.
- `LibraryImport` source generator — NativeAOT-compatible, trim-safe.
- `net8.0` and `net10.0` target frameworks.
- Async helpers (`Document.OpenAsync`) that offload to the thread pool.

[![NuGet](https://img.shields.io/nuget/v/OfficeOxide)](https://www.nuget.org/packages/OfficeOxide)
[![License: MIT OR Apache-2.0](https://img.shields.io/badge/License-MIT%20OR%20Apache--2.0-blue.svg)](https://opensource.org/licenses)

> **Part of the [office_oxide](https://github.com/yfedoseev/office_oxide) toolkit.** Same Rust core, same pass rate as the
> [Rust](https://docs.rs/office_oxide), [Python](../../python/README.md),
> [Go](../../go/README.md), [JavaScript (native)](../../js/README.md),
> and [WASM](../../wasm-pkg/README.md) bindings.

## Quick Start

```bash
dotnet add package OfficeOxide
```

```csharp
using OfficeOxide;

using var doc = Document.Open("report.docx");
Console.WriteLine(doc.Format);        // "docx"
Console.WriteLine(doc.PlainText());
Console.WriteLine(doc.ToMarkdown());
```

## Why OfficeOxide?

- **Fast** — 0.8ms mean DOCX, 5.0ms mean XLSX, 0.7ms mean PPTX; up to 100× faster than python-docx / openpyxl / python-pptx
- **Reliable** — 98.4% pass rate on 6,062 files; zero failures on legitimate Office documents
- **Complete** — 6 formats: DOCX, XLSX, PPTX + legacy DOC, XLS, PPT
- **Permissive** — MIT / Apache-2.0, no AGPL or GPL restrictions
- **NativeAOT-safe** — `LibraryImport` source generator, trim-compatible

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
dotnet add package OfficeOxide
```

The NuGet package bundles pre-built native libraries for:

| Platform | x64 | ARM64 |
|---|---|---|
| Linux (glibc) | Yes | Yes |
| macOS | Yes | Yes (Apple Silicon) |
| Windows | Yes | Yes |

## Editing

```csharp
using var ed = EditableDocument.Open("template.docx");
ed.ReplaceText("{{NAME}}", "Alice");
ed.Save("out.docx");
```

### Spreadsheet cells

```csharp
using var ed = EditableDocument.Open("report.xlsx");
ed.SetCell(0, "A1", "Revenue");
ed.SetCell(0, "B1", 12345.67);
ed.SetCell(0, "C1", true);
ed.Save("report.edited.xlsx");
```

### One-shot helpers

```csharp
string text = OfficeOxide.ExtractText("report.docx");
string md   = OfficeOxide.ToMarkdown("report.docx");
string html = OfficeOxide.ToHtml("report.docx");
```

## Other languages

office_oxide ships the same Rust core through six bindings:

- **Rust** — `cargo add office_oxide` — see [docs.rs/office_oxide](https://docs.rs/office_oxide)
- **Python** — `pip install office-oxide` — see [python/README.md](../../python/README.md)
- **Go** — `go get github.com/yfedoseev/office_oxide/go` — see [go/README.md](../../go/README.md)
- **JavaScript (native)** — `npm install office-oxide` — see [js/README.md](../../js/README.md)
- **WASM** — `npm install office-oxide-wasm` — see [wasm-pkg/README.md](../../wasm-pkg/README.md)

## Why I built this

I needed a library that could read all six Office formats at once — not six separate packages — and I needed it without pulling in a JVM, a Python runtime, or a GPL-licensed dependency. The same Rust binary powers Python via PyO3, Node.js via koffi, Go via cgo, C# via P/Invoke, and the browser via WASM — one fix lands everywhere.

If something's broken or missing, [open an issue](https://github.com/yfedoseev/office_oxide/issues).

— Yury

## License

Dual-licensed under [MIT](../../LICENSE-MIT) or [Apache-2.0](../../LICENSE-APACHE) at your option.

---

**C# / .NET** + **Rust core** | MIT / Apache-2.0 | 98.4% pass rate on 6,062 files | Up to 100× faster than alternatives | 6 formats
