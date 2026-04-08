# office_oxide — Architecture

## Project Structure

```
office_oxide/
├── Cargo.toml                  # Workspace root + library package
├── src/
│   ├── lib.rs                  # Unified API: Document, extract_text, to_markdown
│   ├── format.rs               # DocumentFormat enum + detection
│   ├── error.rs                # OfficeError (wraps all format errors)
│   ├── ir.rs                   # Document IR types
│   ├── ir_render.rs            # IR → plain text / markdown renderers
│   ├── convert_docx.rs         # DOCX → IR converter
│   ├── convert_xlsx.rs         # XLSX → IR converter
│   ├── convert_pptx.rs         # PPTX → IR converter
│   ├── convert_doc.rs          # DOC → IR converter
│   ├── convert_xls.rs          # XLS → IR converter
│   ├── convert_ppt.rs          # PPT → IR converter
│   ├── create.rs               # IR → format creation
│   ├── edit.rs                 # Unified editing API
│   ├── python.rs               # PyO3 bindings (feature: python)
│   ├── wasm.rs                 # wasm-bindgen bindings (feature: wasm)
│   ├── core/                   # Shared OPC/XML primitives
│   │   ├── opc.rs              # OpcReader/OpcWriter (ZIP-based packages)
│   │   ├── xml.rs              # XML parsing utilities, namespace constants
│   │   ├── content_types.rs    # [Content_Types].xml parser
│   │   ├── relationships.rs    # .rels parser + builder
│   │   ├── editable.rs         # EditablePackage (round-trip OPC editing)
│   │   ├── theme.rs            # DrawingML theme parser
│   │   ├── properties.rs       # Dublin Core + app properties
│   │   ├── units.rs            # Twip, HalfPoint, Emu
│   │   ├── parallel.rs         # cfg-gated parallel/sequential map
│   │   ├── traits.rs           # OfficeDocument trait
│   │   └── error.rs            # Core error type
│   ├── cfb/                    # Compound File Binary (OLE2) reader
│   ├── docx/                   # DOCX parser, writer, editor
│   ├── xlsx/                   # XLSX parser, writer, editor
│   ├── pptx/                   # PPTX parser, writer, editor
│   ├── doc/                    # Legacy .doc parser
│   ├── xls/                    # Legacy .xls parser
│   └── ppt/                    # Legacy .ppt parser
├── crates/
│   ├── office_oxide_cli/       # CLI binary: office-oxide
│   └── office_oxide_mcp/       # MCP server binary: office-oxide-mcp
├── tests/                      # Integration tests
├── python/                     # Python package (office_oxide/)
├── wasm-pkg/                   # npm package config
└── docs/                       # Architecture + spec references
```

## Module Dependency Graph

```
                    src/lib.rs (unified API + bindings)
                   /        |          \
              src/docx   src/xlsx   src/pptx    (OOXML formats)
              src/doc    src/xls    src/ppt     (legacy formats)
                   \        |          /
                    src/core (OPC, XML, themes)
                         |
                      src/cfb (OLE2 container — legacy formats only)
```

Binary crates depend on the library:
```
office_oxide_cli  ──→  office_oxide (lib)
office_oxide_mcp  ──→  office_oxide (lib)
```

## core — Shared Primitives

All OOXML formats share the Open Packaging Conventions (OPC) layer:

### OPC Layer
- **ZIP archive** read/write (using `zip` crate)
- **Content Types** parser (`[Content_Types].xml`)
- **Relationships** parser (`.rels` files)
- **Part resolution** — URI-based part lookup
- **Case-insensitive** ZIP entry lookup + backslash path normalization

### XML Layer
- **Fast XML parsing** via `quick-xml` (SAX-style, zero-copy)
- **Namespace handling** — OOXML Transitional + Strict dual namespace matching
- **Namespace-agnostic** attribute lookup for prefixed attributes (`d3p1:id`, etc.)
- **Tolerant numeric parsing** — strips unit suffixes, handles decimals

### DrawingML Shared
- **Themes** (`a:theme`) — colors, fonts, effects
- **Colors** — scheme colors, RGB, HSL, system colors, tint/shade
- **Units** — Twip, HalfPoint, Emu

### Core Properties
- **Dublin Core metadata** — title, creator, subject, description
- **App properties** — application, version, word count, etc.

## docx — Word Documents

### Read Path
```
.docx file
  → ZIP extraction (core::opc)
  → Relationship resolution
  → document.xml parsing (SAX-style, non-trimming reader for whitespace preservation)
  → Element tree: Document → Body → [Block-level elements]
      ├── Paragraph (w:p) → Run (w:r) → Text (w:t)
      ├── Table (w:tbl) → Row → Cell → [Block-level elements] (recursive)
      ├── Hyperlinks (resolved via relationships)
      └── Headers/Footers
  → Style resolution (styles.xml)
  → Numbering resolution (numbering.xml)
```

### Write / Edit
- `DocxWriter` — builder API for creating new documents
- `EditableDocx` — load, replace text in `<w:t>` elements, save

## xlsx — Excel Spreadsheets

### Read Path
```
.xlsx file
  → ZIP extraction (core::opc)
  → sharedStrings.xml → String table (parsed first, non-trimming reader)
  → styles.xml → Number formats, cell styles
  → workbook.xml → Sheet list
  → sheet{N}.xml → Cell grid (parallel when feature enabled)
      ├── Cell types: string, number, boolean, error, inline string, date
      ├── Shared string dereference (inline during parse)
      └── Date detection (built-in format IDs + custom format scanning)
```

### Key Design Decisions
- **Shared strings loaded first** — worksheets reference them by index
- **Date handling**: 1900 date system with Lotus 1-2-3 bug compatibility (serial 60 = Feb 29, 1900)
- **Phase 1/Phase 2 parsing**: gather raw data sequentially (requires `&mut archive`), then parse worksheets in parallel via `core::parallel::map_collect`

### Write / Edit
- `XlsxWriter` — builder API with sheets, rows, cells
- `EditableXlsx` — load, set cell values, save

## pptx — PowerPoint Presentations

### Read Path
```
.pptx file
  → ZIP extraction (core::opc)
  → Theme resolution
  → presentation.xml → Slide list
  → slides/slide{N}.xml → Shape tree (parallel when feature enabled)
      └── Shape tree (p:spTree)
          ├── AutoShape (p:sp) → Text body
          ├── Picture (p:pic) → Image reference
          ├── Group shape (p:grpSp) → Nested shapes
          ├── Graphic frame (p:graphicFrame) → Tables
          └── Connector (p:cxnSp)
  → Notes slide extraction
  → Spatial sorting (y then x) for text extraction order
```

### Write / Edit
- `PptxWriter` — builder API with slides, titles, text, bullet lists
- `EditablePptx` — load, replace text in `<a:t>` elements, save

## Legacy Formats (doc, xls, ppt)

All three use the CFB (Compound File Binary / OLE2) container:

```
.doc/.xls/.ppt file
  → cfb::CfbReader (header, FAT, DIFAT, directory, mini-stream)
  → Stream extraction (case-insensitive lookup)
  → Format-specific binary parsing
```

- **doc**: FIB → piece table → text reassembly (CP1252 + UTF-16LE dual encoding)
- **xls**: BIFF8 record iterator with CONTINUE merging → SST, cell records (LABELSST, NUMBER, RK, MULRK, FORMULA)
- **ppt**: 8-byte record headers → container/atom tree → TextCharsAtom (UTF-16LE) / TextBytesAtom (Latin-1)

## Document IR

All format conversions flow through a format-agnostic intermediate representation:

```
Source Format                     Target Format
─────────────                     ─────────────
DOCX ──┐                    ┌──→ Plain Text
XLSX ──┤                    ├──→ Markdown
PPTX ──┤──→ Document IR ────├──→ JSON (structured)
DOC  ──┤    (unified)       └──→ CSV (tables only)
XLS  ──┤
PPT  ──┘
```

IR types (`src/ir.rs`): `DocumentIR`, `Section`, `Element` (Heading, Paragraph, Table, List, Image, ThematicBreak), `InlineContent` (Text, LineBreak). All derive `serde::Serialize`/`Deserialize` for direct JSON serialization.

## Performance Strategy

- **SAX-style XML parsing** (not DOM) — constant memory via `quick-xml`
- **Zero-copy attributes** — parse without allocation using borrowed API
- **Optional memory-mapping** — `mmap` feature for large files
- **Parallel parsing** — `parallel` feature uses Rayon for sheets/slides
- **Lazy parsing** — only parse parts that are accessed
- **Stack-depth protection** — `with_parse_stack` spawns threads when stack is low (Python/WASM interop)

## Testing Strategy

- **310+ unit and integration tests** across all formats
- **5,146-file corpus validation** — 98.3% pass rate, all failures are genuinely invalid files
- **Round-trip testing** — create → write → read back → compare
- **CI matrix** — Linux/macOS/Windows, stable/beta/nightly Rust, Python 3.8–3.14
- **85% code coverage enforcement** via cargo-llvm-cov
