# office_oxide — Architecture

## Workspace Structure

```
office_oxide/
├── Cargo.toml                  # Workspace root
├── docs/
│   ├── MISSION.md              # Project mission and roadmap
│   ├── ARCHITECTURE.md         # This file
│   ├── specs/
│   │   ├── opc_shared_spec.md  # Open Packaging Conventions (shared)
│   │   ├── docx_spec.md        # WordprocessingML spec reference
│   │   ├── xlsx_spec.md        # SpreadsheetML spec reference
│   │   └── pptx_spec.md        # PresentationML spec reference
│   └── architecture/
│       └── conversion_pipeline.md
├── crates/
│   ├── office_core/            # Shared primitives
│   ├── docx_oxide/             # Word document processing
│   ├── xlsx_oxide/             # Excel spreadsheet processing
│   ├── pptx_oxide/             # PowerPoint processing
│   └── office_oxide/           # Unified API + bindings
└── pdf_oxide -> ~/projects/pdf_oxide_fixes  # Symlink or git submodule
```

## Crate Dependency Graph

```
                    office_oxide (unified API + bindings)
                   /        |          \
              docx_oxide  xlsx_oxide  pptx_oxide
                   \        |          /
                    office_core (OPC, XML, themes)
                         |
                     pdf_oxide (for → PDF conversion)
```

## office_core — Shared Primitives

All OOXML formats share the Open Packaging Conventions (OPC) layer. This crate provides:

### OPC Layer
- **ZIP archive** read/write (using `zip` crate)
- **Content Types** parser (`[Content_Types].xml`)
- **Relationships** parser (`.rels` files)
- **Part resolution** — URI-based part lookup

### XML Layer
- **Fast XML parsing** via `quick-xml` (SAX-style, zero-copy)
- **Namespace handling** — resolve OOXML namespaces
- **Shared string table** — common across XLSX, used concept in DOCX/PPTX
- **XML writing** — streaming XML output

### DrawingML Shared
- **Themes** (`a:theme`) — colors, fonts, effects
- **Colors** — scheme colors, RGB, HSL, system colors
- **Fonts** — major/minor font schemes, font fallback
- **Styles** — shared style concepts

### Core Properties
- **Dublin Core metadata** — title, creator, subject, description
- **App properties** — application, version, word count, etc.
- **Custom properties** — user-defined key-value pairs

### Common Types
- **EMU** (English Metric Units) — shared coordinate system
- **ST_Percentage**, **ST_Coordinate**, etc. — shared simple types
- **Color types** — scheme, RGB, theme color with tint/shade

## docx_oxide — Word Documents

### Read Path
```
.docx file
  → ZIP extraction (office_core)
  → Relationship resolution
  → document.xml parsing
  → Element tree: Document → Body → [Block-level elements]
      ├── Paragraph (w:p)
      │   ├── Run (w:r) → Text (w:t)
      │   ├── Hyperlink
      │   ├── Field codes
      │   └── Drawing/Image
      ├── Table (w:tbl)
      │   └── Row → Cell → [Block-level elements] (recursive)
      ├── Structured Document Tag (w:sdt)
      └── Section Properties (w:sectPr)
  → Style resolution (styles.xml)
  → Numbering resolution (numbering.xml)
  → Header/Footer processing
```

### Write Path
```
Builder API
  → Element tree construction
  → Style sheet generation
  → Numbering definitions
  → Relationship tracking
  → XML serialization (quick-xml)
  → ZIP packaging (office_core)
  → .docx output
```

### Conversion Paths
- **DOCX → Text**: Walk element tree, extract w:t content with spacing
- **DOCX → Markdown**: Map elements to Markdown (headings, lists, tables, bold/italic)
- **DOCX → HTML**: Map elements to HTML with CSS styles
- **DOCX → PDF**: Use pdf_oxide writer — map fonts, layout paragraphs, render tables

## xlsx_oxide — Excel Spreadsheets

### Read Path
```
.xlsx file
  → ZIP extraction (office_core)
  → workbook.xml → Sheet list, defined names
  → sharedStrings.xml → String table
  → styles.xml → Number formats, cell styles
  → sheet{N}.xml → Cell grid
      ├── Row (r attribute = row number)
      │   └── Cell (r attribute = "A1" reference)
      │       ├── Type: string (s), number (n), boolean (b), error (e), inline string (inlineStr)
      │       ├── Value (v element)
      │       ├── Formula (f element)
      │       └── Style index (s attribute → styles.xml)
      ├── Merge cells
      ├── Conditional formatting
      └── Data validations
  → Formula evaluation (optional)
```

### Key Design Decisions
- **Shared strings**: XLSX stores repeated strings once in a lookup table. Must load this first.
- **Cell references**: "A1" style (column letters + row number). Need bidirectional conversion.
- **Date handling**: Excel stores dates as serial numbers. Need epoch conversion (1900 vs 1904 date system).
- **Formula evaluation**: MVP = extract formula text. Future = evaluate simple formulas.

## pptx_oxide — PowerPoint Presentations

### Read Path
```
.pptx file
  → ZIP extraction (office_core)
  → presentation.xml → Slide list, slide size
  → slideMasters/ → Base layouts and styles
  → slideLayouts/ → Layout templates
  → slides/slide{N}.xml → Slide content
      └── Shape tree (p:spTree)
          ├── Shape (p:sp) → Text body, geometry
          ├── Picture (p:pic) → Image reference
          ├── Group shape (p:grpSp) → Nested shapes
          ├── Connection (p:cxnSp)
          └── Graphic frame (p:graphicFrame) → Tables, charts
  → Theme resolution
  → Notes extraction
```

### Key Design Decisions
- **Slide inheritance**: Slides inherit from layouts, which inherit from masters. Must resolve the chain.
- **Shape positioning**: Absolute positioning in EMU (1 inch = 914400 EMU).
- **Text extraction order**: Shapes have no guaranteed reading order. Need spatial sorting.

## Conversion Pipeline Architecture

All format conversions flow through an intermediate representation:

```
Source Format                     Target Format
─────────────                     ─────────────
DOCX ──┐                    ┌──→ PDF (via pdf_oxide)
XLSX ──┤                    ├──→ Markdown
PPTX ──┤──→ Document IR ────├──→ HTML
PDF  ──┘    (unified)       ├──→ Plain Text
                            ├──→ JSON (structured)
                            └──→ CSV (tables only)
```

The **Document IR** is a format-agnostic intermediate representation:
- Headings with levels
- Paragraphs with inline formatting (bold, italic, underline, etc.)
- Tables with cells, spans, and alignment
- Images with dimensions and alt text
- Lists (ordered, unordered) with nesting
- Metadata (title, author, dates)
- Page/slide/sheet boundaries

This IR enables any-to-any conversion without N*M converter implementations.

## Performance Strategy

### Zero-Copy Where Possible
- Parse XML attributes without allocation using `quick-xml`'s borrowed API
- Memory-map ZIP entries for large files
- Return `&str` references into parsed data

### Streaming Processing
- SAX-style XML parsing (not DOM) — constant memory
- Process sheets/slides independently
- Stream output for large spreadsheets (millions of rows)

### Parallelism
- Parse independent sheets/slides in parallel (Rayon)
- Parallel style/theme resolution
- Concurrent image extraction

### Caching
- Cache parsed styles, themes, shared strings
- Lazy parsing — only parse parts that are accessed
- LRU cache for frequently accessed elements

## Testing Strategy

### Corpus Testing
- Collect real-world documents (same approach as pdf_oxide's 3,830 PDF corpus)
- Sources: government databases, academic papers, public datasets
- Automated quality scoring

### Round-Trip Testing
- Create document → Write → Read back → Compare
- Ensures write path produces valid files

### Compatibility Testing
- Test output opens correctly in MS Office, LibreOffice, Google Docs
- Test reading files produced by each of these applications

### Fuzz Testing
- Fuzz ZIP parsing layer
- Fuzz XML parsing
- Fuzz cell reference parsing
- Fuzz number format parsing
