# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2026-04-07

> Initial public release

### Added

- **office_core**: Shared OPC/ZIP/XML primitives for OOXML document processing (55 tests)
  - `OpcReader`/`OpcWriter` for ZIP-based OPC packages
  - Theme parsing, color resolution, unit types (Twip, HalfPoint, Emu)
  - Namespace-aware XML utilities with strict OOXML support

- **docx_oxide**: High-performance Word document (.docx) processing (36 tests)
  - SAX-style parsing of document.xml, styles.xml, numbering.xml
  - Hyperlink resolution, heading detection, list grouping
  - Plain text and Markdown extraction

- **xlsx_oxide**: High-performance Excel spreadsheet (.xlsx) processing (57 tests)
  - Shared string table, cell parsing, date detection (1900 date system with Lotus bug)
  - CSV (RFC 4180) and Markdown (pipe tables) output
  - Rich text and style support

- **pptx_oxide**: High-performance PowerPoint presentation (.pptx) processing (40 tests)
  - Slide parsing with spatial sort, shape types, text body extraction
  - Notes slide support, hyperlink resolution
  - Plain text and Markdown extraction

- **cfb_oxide**: Pure Rust CFBF/OLE2 container reader (18 tests)
  - Supports v3 (512-byte) and v4 (4096-byte) sectors
  - Mini-stream support, case-insensitive stream access

- **doc_oxide**: Legacy Word Binary (.doc) parser (15 tests)
  - FIB parsing, piece table extraction, dual encoding (CP1252 + UTF-16LE)

- **xls_oxide**: Legacy Excel Binary (.xls) BIFF8 parser (24 tests)
  - CONTINUE record merging, SST parsing, RK decode
  - Cell records: LABELSST, NUMBER, RK, MULRK, FORMULA, BOOLERR

- **ppt_oxide**: Legacy PowerPoint Binary (.ppt) parser (15 tests)
  - Container/atom record parsing, TextCharsAtom/TextBytesAtom
  - SlideListWithText container for slide-level grouping

- **office_oxide**: Unified API across all formats (29 tests)
  - `Document::open()`, `plain_text()`, `to_markdown()`, `to_ir()`
  - Format-agnostic IR with Sections, Elements (Heading, Paragraph, Table, List, Image)
  - Feature flags for each format (all default-on)

- **Creation API**: Write DOCX, XLSX, PPTX from scratch via IR
  - `DocxWriter`, `XlsxWriter`, `PptxWriter`
  - `create_from_ir()` for IR-to-format conversion

- **Editing API**: Modify existing documents preserving unmodified parts
  - `EditableDocx`, `EditableXlsx`, `EditablePptx`
  - `replace_text()`, `set_cell()` operations

- **Python bindings**: PyO3-based with type stubs and PEP 561 marker
- **WASM bindings**: wasm-bindgen-based for browser/Node.js usage

### Robustness

- OOXML Strict namespace support with dual namespace matching
- Case-insensitive ZIP entry lookup with backslash path normalization
- CRC checksum tolerance, tolerant numeric parsing
- Shared string DoS cap (32,768 chars), optional parts graceful degradation
- **98.3% pass rate** on 5,146-file OOXML corpus (90 failures all genuinely invalid)

[0.1.0]: https://github.com/yfedoseev/office_oxide/releases/tag/v0.1.0
