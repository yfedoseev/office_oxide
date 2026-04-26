# office_oxide examples

Numbered scenarios (01–06 for Rust, 01–04 for other languages) implemented
across every language binding. All numbered examples are **self-contained** —
they create their own test data in memory and exit 0 on success, so they run
unmodified in CI without external fixture files.

## Numbered examples (self-contained, run in CI)

| # | Scenario | Rust | Python | Go | JS | C# |
|---|---|---|---|---|---|---|
| 01 | Extract text, Markdown, IR | ✓ | ✓ | ✓ | ✓ | ✓ |
| 02 | Create documents | Rich DocxWriter / XlsxWriter / PptxWriter | create_from_markdown | create_from_markdown | create_from_markdown | create_from_markdown |
| 03 | Edit documents (replace_text, set_cell) | ✓ | ✓ | ✓ | ✓ | ✓ |
| 04 | XLSX formulas + cell styles | ✓ | batch demo | — | — | — |
| 05 | Edit roundtrip | ✓ | — | — | — | — |
| 06 | Markdown → DOCX / XLSX / PPTX | ✓ | — | — | — | — |

## Classic examples (accept a file path argument)

| Scenario | File stem |
|---|---|
| Extract plain text + Markdown | `extract` |
| Fill a template (text replace) | `replace` |
| Read a spreadsheet into rows | `read_xlsx` |

## Layout

| Binding | Directory |
|---|---|
| Rust (numbered) | [`rust/01_extract.rs` … `rust/06_from_markdown.rs`](rust/) |
| Rust (classic) | [`rust/extract.rs`](rust/extract.rs), [`rust/make_smoke.rs`](rust/make_smoke.rs) |
| Python | [`python/`](python/) |
| Go | [`go/`](go/) |
| C# / .NET | [`csharp/`](csharp/) |
| JavaScript (native, koffi) | [`javascript/`](javascript/) |
| Raw C FFI | [`c/`](c/) |

## Running the Rust examples

```sh
# Self-contained (no file argument needed)
cargo run --example 01_extract
cargo run --example 02_create_rich
cargo run --example 03_create_xlsx
cargo run --example 04_create_pptx
cargo run --example 05_edit
cargo run --example 06_from_markdown

# Classic (supply a file)
cargo run --example extract -- path/to/file.docx
```

## Running other language examples

Each language's numbered examples run automatically in CI after the
respective bindings are built. To run locally, build the native library first:

```sh
cargo build --release --lib
```

Then follow the language-specific instructions in the binding's README.
