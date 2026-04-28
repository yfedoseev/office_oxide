# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2026-04-27

> Initial public release

### Cross-language bindings

- **Rust core** (`office_oxide` on crates.io): unified `Document` handle for
  all six formats, `EditableDocument` for DOCX/XLSX/PPTX editing, format-
  agnostic `DocumentIR`.
- **Python** (`office-oxide` on PyPI): context-manager `Document` /
  `EditableDocument`, `os.PathLike` support, complete type stubs in
  `_native.pyi` (`Literal` format names, `_Path` alias).
- **Go** (`github.com/yfedoseev/office_oxide/go`): CGo wrapper over the C FFI
  with idiomatic `Open` / `Close` / error-return API, `go/cmd/install` helper
  that fetches the matching native archive and prints the
  `CGO_CFLAGS` / `CGO_LDFLAGS` to export.
- **C# / .NET** (`OfficeOxide` on NuGet): `LibraryImport` P/Invoke,
  `IDisposable`, `async/await`, `IsAotCompatible=true`, `IsTrimmable=true`.
  Four `SetCell` overloads + `SetCellEmpty`. Net 8 and net 10 target
  frameworks.
- **Node.js native** (`office-oxide` on npm): [koffi](https://koffi.dev)-based,
  no node-gyp, ESM + CJS entry points with an `exports` map, TypeScript
  definitions, `Symbol.dispose` support, platform prebuilds staged into
  `prebuilds/<platform>-<arch>/`.
- **WASM** (`office-oxide-wasm` on npm): three sub-path exports — default
  ESM for bundlers, `office-oxide-wasm/node` for CJS, `office-oxide-wasm/web`
  for native-ESM browser imports. TypeScript definitions shipped.
- **C FFI** (`include/office_oxide_c/office_oxide.h`): stable
  `office_document_*` / `office_editable_*` surface with out-param error
  codes and explicit memory ownership. Exported from the cdylib + staticlib;
  the substrate that Go, C#, and Node-native link against.

### Tooling

- **CLI** (`office-oxide` binary): `text`, `markdown`, `html`, `info`, `ir`
  subcommands.
- **MCP server** (`office-oxide-mcp` binary): `extract` and `info` tools
  over JSON-RPC 2.0 / stdio.

### Performance

- Up to 100× faster than `python-docx`, `openpyxl`, `python-pptx`, `xlrd`.
- Beats `calamine` on XLSX and all Rust / Python alternatives on .xls.
- **100% pass rate on valid Office files** (6,062-file corpus: LibreOffice,
  Apache POI, python-pptx, python-docx, Pandoc, etc.). All 97 non-passing
  files are invalid inputs — corrupted ZIPs, missing required parts, malformed
  XML, or non-Office files with Office extensions.

### Documentation & examples

- Per-language getting-started guides in [`docs/`](docs/): Rust, Python,
  Go, C#, JavaScript (native), WASM, and C / raw FFI.
- Per-binding READMEs in [`python/`](python/README.md), [`go/`](go/README.md),
  [`csharp/OfficeOxide/`](csharp/OfficeOxide/README.md),
  [`js/`](js/README.md), [`wasm-pkg/`](wasm-pkg/README.md).
- Identical `extract` / `replace` / `read_xlsx` demos per language under
  [`examples/`](examples/).

### Release CI

- Version parity across `Cargo.toml`, `pyproject.toml`,
  `wasm-pkg/package.json`, `js/package.json`, and
  `csharp/OfficeOxide/OfficeOxide.csproj`.
- 6-target native-lib build matrix producing `.tar.gz` / `.zip` archives
  with the shared library, static archive, and public header.
- 3-target WASM build (bundler / nodejs / web) with per-target module-type
  hints so Node + bundlers + browsers each load the right code.
- NuGet packaging with `runtimes/<rid>/native/` prebuilts, Node
  `prebuilds/<platform>-<arch>/` staging, and a `go/v*` module tag for the
  Go module proxy.

## [0.1.0-draft] - 2026-04-07 (not released)

> Internal milestone before the public release above.

Single consolidated `office_oxide` crate with modules per format, plus
companion workspace crates `office_oxide_cli` and `office_oxide_mcp`.

### Added

- **Unified API** (`crate::Document`)
  - `Document::open()`, `Document::from_reader()`, `plain_text()`,
    `to_markdown()`, `to_html()`, `to_ir()`, `format_name()`
  - Format auto-detection from file extension for all 6 formats
  - `as_docx()` / `as_xlsx()` / `as_pptx()` / `as_doc()` / `as_xls()` /
    `as_ppt()` escape hatches for format-specific types
  - Convenience functions: `extract_text()`, `to_markdown()`, `to_html()`
    at crate root

- **Format-agnostic IR** (`crate::ir::DocumentIR`)
  - Sections, Elements (Heading, Paragraph, Table, List, Image,
    ThematicBreak), serializable to/from JSON
  - DOCX→IR: heading detection via outline level + style resolution,
    list grouping, table vMerge → row_span
  - XLSX→IR: worksheets → sections, cell grids → tables, first row as
    header
  - PPTX→IR: slides → sections, spatial sort, title placeholder → section
    title
  - IR renderers: `plain_text()` and `to_markdown()`

- **OOXML module — `crate::docx`** (36 tests)
  - SAX-style parsing of `document.xml`, `styles.xml`, `numbering.xml`
  - Hyperlink resolution, heading detection, list grouping
  - Plain text, Markdown, HTML extraction

- **OOXML module — `crate::xlsx`** (57 tests)
  - Shared string table, cell parsing, 1900 date system with Lotus bug
  - Built-in + custom number format detection
  - CSV (RFC 4180), Markdown (pipe tables), HTML output
  - Rich-text and style support

- **OOXML module — `crate::pptx`** (40 tests)
  - Slide parsing with spatial sort, shape types (AutoShape, Picture,
    Group, GraphicFrame, Connector), text body extraction
  - Notes slide support, inline hyperlink resolution
  - Plain text, Markdown, HTML extraction

- **Legacy module — `crate::cfb`** (18 tests)
  - Pure Rust CFBF / OLE2 container reader
  - Supports v3 (512-byte) and v4 (4096-byte) sectors, mini-streams,
    case-insensitive stream access, path-based lookup

- **Legacy module — `crate::doc`** (15 tests)
  - Word Binary (.doc) parser built on `crate::cfb`
  - FIB parsing, piece-table extraction, dual encoding
    (compressed CP1252 + Unicode UTF-16LE)
  - Field-code stripping and special-char sanitization

- **Legacy module — `crate::xls`** (24 tests)
  - Excel Binary (.xls) BIFF8 parser built on `crate::cfb`
  - CONTINUE record merging, SST (compressed + wide + rich text),
    RK decode
  - Cell records: LABELSST, NUMBER, RK, MULRK, FORMULA, BOOLERR,
    LABEL, BLANK

- **Legacy module — `crate::ppt`** (15 tests)
  - PowerPoint Binary (.ppt) parser built on `crate::cfb`
  - 8-byte record headers, container vs atom records
  - Text extraction from TextCharsAtom (UTF-16LE) and TextBytesAtom
    (Latin-1), SlideListWithText grouping

- **Shared core — `crate::core`** (55 tests)
  - `OpcReader` / `OpcWriter` for ZIP-based OPC packages
  - Theme parsing, color resolution, unit types (`Twip`, `HalfPoint`,
    `Emu`)
  - Namespace-aware XML utilities with OOXML Strict support

- **Creation API** — write OOXML documents from scratch
  - `crate::docx::write::DocxWriter`, `crate::xlsx::write::XlsxWriter`,
    `crate::pptx::write::PptxWriter`
  - `crate::create::create_from_ir()` and
    `create_from_ir_to_writer()` for IR-to-format conversion

- **Editing API** — modify existing documents while preserving unmodified
  parts
  - `crate::docx::edit::EditableDocx`, `crate::xlsx::edit::EditableXlsx`,
    `crate::pptx::edit::EditablePptx`
  - Unified `crate::edit::EditableDocument` with `replace_text()`
    and `set_cell()`
  - `crate::core::editable::EditablePackage` round-trips rels via
    `RelationshipsBuilder::add_with_id()`

- **Python bindings** — PyO3 0.28, extension module `office_oxide._native`
  - Type stubs (`_native.pyi`) and PEP 561 `py.typed` marker
  - Wheels for Linux, macOS, Windows across Python 3.8–3.14

- **WASM bindings** — `wasm-bindgen`, `WasmDocument` class
  - `new(data, format)`, `plainText()`, `toMarkdown()`, `toHtml()`,
    `toIr()`
  - npm package `office-oxide-wasm`

- **CLI** — `office_oxide_cli` workspace crate, binary `office-oxide`
  - Subcommands: `text`, `markdown`, `html`, `info`, `ir`

- **MCP server** — `office_oxide_mcp` workspace crate, binary
  `office-oxide-mcp`
  - JSON-RPC 2.0 over stdin/stdout
  - Tools: `extract` (text / markdown / html / ir), `info`

- **Feature flags**
  - `python` — PyO3 extension module
  - `wasm` — wasm-bindgen bindings
  - `mmap` — `memmap2`-backed file reading
  - `parallel` — `rayon`-based parallel parsing

### Robustness

- OOXML Strict namespace support via dual-namespace matching
  (`strict_alternate()`) and relationship-type normalization
- Case-insensitive ZIP entry lookup with backslash path normalization
- Namespace-agnostic attribute lookup (`optional_prefixed_attr_str()`)
  for `d3p1:id` and similar
- Percent-encoding decoding in `PartName::new()`
- CRC checksum tolerance in `read_zip_entry()` (accepts data on mismatch)
- Tolerant numeric parsing: `parse_numeric()` strips unit suffixes,
  handles decimals
- Shared-string DoS cap: `MAX_CELL_STRING_LEN = 32_768` in `xlsx`
- Optional parts (`numbering.xml`, theme) degrade gracefully on read
  errors
- XLSX border aliases: `start`/`end` mapped to `left`/`right` for Strict
  OOXML

### Validation

- **98.4% pass rate on a 6,062-file corpus** (5,965 / 6,062) across 11
  open-source test suites (LibreOffice Core, Apache POI, Open XML SDK,
  ClosedXML, Pandoc, python-docx, python-pptx, Apache Tika, calamine,
  openpreserve, oletools)
- Zero failures on legitimate Word 97+ / Excel 97+ / PowerPoint 97+
  files — all 97 non-passing files are invalid inputs: 43 invalid
  ZIP/CFB archives, 21 missing required parts, 18 malformed XML, and
  15 non-Office files (WordPerfect, pre-OLE2 Excel 3/4) misnamed with
  Office extensions
- See [BENCHMARKS.md](BENCHMARKS.md) for per-format timings and the
  full failure breakdown

[0.1.0]: https://github.com/yfedoseev/office_oxide/releases/tag/v0.1.0
