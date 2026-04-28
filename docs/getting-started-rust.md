# Getting Started with office_oxide (Rust)

`office_oxide` is a pure-Rust library for parsing, converting, and editing Microsoft Office documents: DOCX, XLSX, PPTX, plus their legacy binary predecessors DOC, XLS, and PPT. One crate, one unified `Document` handle, no native dependencies.

## Installation

Add to your `Cargo.toml`:

```toml
[dependencies]
office_oxide = "0.1.0"
```

### Feature Flags

```toml
[dependencies]
# Default build — full read/write/edit support for all six formats.
office_oxide = "0.1.0"

# Memory-mapped opens for large OOXML files.
office_oxide = { version = "0.1.0", features = ["mmap"] }

# Parallel parsing helpers.
office_oxide = { version = "0.1.0", features = ["parallel"] }
```

## Quickstart

Extract plain text from a DOCX file:

```rust
use office_oxide::Document;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let doc = Document::open("report.docx")?;
    println!("{}", doc.plain_text());
    Ok(())
}
```

Or use the one-shot helper:

```rust
let text = office_oxide::extract_text("report.docx")?;
```

## Core API

The unified `Document` handle works the same for every format — extension detection and magic-byte sniffing pick the right parser.

```rust
use office_oxide::{Document, DocumentFormat};

let doc = Document::open("file.xlsx")?;

// Identify the format at runtime
assert_eq!(doc.format(), DocumentFormat::Xlsx);

// Text extraction
let plain = doc.plain_text();

// Rendered output
let md = doc.to_markdown();
let html = doc.to_html();

// Format-agnostic IR (see "Advanced")
let ir = doc.to_ir();

// Save/convert to any supported format (extension-driven)
doc.save_as("file.docx")?;
```

`Document::open` accepts anything `AsRef<Path>`; `Document::from_reader` takes a `Read + Seek + Send + 'static` with an explicit `DocumentFormat`. Both allocate a boxed inner document, so hold onto the returned handle.

Module-level one-shots for common cases:

```rust
let text = office_oxide::extract_text("file.docx")?;
let md   = office_oxide::to_markdown("file.pptx")?;
let html = office_oxide::to_html("file.xlsx")?;
```

### Format-specific access

When you need richer format-specific data (worksheets, slides, table structures), unwrap the inner document:

```rust
if let Some(xlsx) = doc.as_xlsx() {
    for sheet in xlsx.sheets() {
        println!("sheet: {}", sheet.name());
    }
}
```

The same pattern applies to `as_docx`, `as_pptx`, `as_doc`, `as_xls`, and `as_ppt`.

## Editing

`EditableDocument` implements a read-modify-write workflow that preserves unmodified OPC parts (images, charts, styles, relationships) verbatim. Only DOCX, XLSX, and PPTX are editable.

### Replace text in a DOCX

```rust
use office_oxide::edit::EditableDocument;

let mut doc = EditableDocument::open("template.docx")?;
let n = doc.replace_text("{{name}}", "Alice");
println!("{n} replacements");
doc.save("out.docx")?;
```

`replace_text` walks `<w:t>` elements in DOCX and `<a:t>` elements in PPTX. It returns the number of replacements (0 for XLSX — use `set_cell` instead).

### Set an XLSX cell

```rust
use office_oxide::edit::EditableDocument;
use office_oxide::xlsx::edit::CellValue;

let mut wb = EditableDocument::open("budget.xlsx")?;
wb.set_cell(0, "B2", CellValue::Number(42.0))?;
wb.set_cell(0, "A1", CellValue::String("Total".into()))?;
wb.set_cell(0, "C1", CellValue::Boolean(true))?;
wb.set_cell(0, "D1", CellValue::Empty)?;
wb.save("budget.xlsx")?;
```

Sheet indices are zero-based; `cell_ref` uses standard spreadsheet notation (`A1`, `AA12`).

### Write to any `Write + Seek`

```rust
let mut buf = std::io::Cursor::new(Vec::new());
doc.write_to(&mut buf)?;
let bytes: Vec<u8> = buf.into_inner();
```

## Advanced

### Format-agnostic IR

`DocumentIR` is the structural bridge between formats. It powers `to_html`, `save_as`, and legacy-format conversion.

```rust
let ir = doc.to_ir();
println!("{} sections", ir.sections.len());

// Convert a DOC into a DOCX via IR
let legacy = Document::open("old.doc")?;
legacy.save_as("migrated.docx")?;
```

The IR is `Serialize`/`Deserialize` (via serde) so you can emit JSON for downstream tooling.

### Open from bytes

```rust
use std::io::Cursor;
use office_oxide::{Document, DocumentFormat};

let bytes: Vec<u8> = std::fs::read("file.pptx")?;
let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Pptx)?;
```

### Legacy binary formats

DOC, XLS, and PPT are parsed directly from their CFB/OLE2 containers — no external converter required. Read-only extraction and IR conversion are supported; `save_as` will transparently produce a DOCX / XLSX / PPTX from a legacy input, so you can migrate legacy corpora in one line:

```rust
Document::open("old.xls")?.save_as("modern.xlsx")?;
```

### Memory-mapped opens

With the `mmap` feature enabled, `Document::open_mmap` avoids copying large OOXML files into heap memory:

```rust
let doc = Document::open_mmap("huge.xlsx")?;
```

Only DOCX/XLSX/PPTX are mmap-able; the legacy CFB parsers require owned buffers.

## Error Handling

All fallible entry points return `office_oxide::Result<T>`, i.e. `Result<T, OfficeError>`. The error enum covers IO, parse, unsupported-format, and extraction failures.

```rust
use office_oxide::{Document, OfficeError};

match Document::open("weird.file") {
    Ok(doc) => println!("{}", doc.plain_text()),
    Err(OfficeError::UnsupportedFormat(ext)) => eprintln!("cannot open .{ext}"),
    Err(e) => eprintln!("failed: {e}"),
}
```

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| `UnsupportedFormat("(none)")` | Path has no extension — open via `from_reader` with an explicit `DocumentFormat`. |
| Garbled DOC text | The source file is encrypted or uses an uncommon piece-table encoding; check with a hex-dump for CFB magic `D0 CF 11 E0`. |
| Missing hyperlinks in DOCX | Hyperlinks resolve via `w:rels` — verify the `.rels` sidecar is present inside the ZIP. |
| Stack overflow on tiny-stack threads | `office_oxide` already spawns a 16 MB parse thread when `RLIMIT_STACK < 12 MB`; if you pool your own threads, pass `Builder::stack_size(16 * 1024 * 1024)`. |

## Links

- Crate source: `src/lib.rs`, `src/edit.rs`
- Per-format modules: `src/docx/`, `src/xlsx/`, `src/pptx/`, `src/doc/`, `src/xls/`, `src/ppt/`
- Public API on crates.io: https://crates.io/crates/office_oxide
- Rustdoc: https://docs.rs/office_oxide
- Architecture overview: `docs/ARCHITECTURE.md`
