# office-oxide CLI

The fastest CLI for Office document processing. Extract text, convert to markdown, dump IR, and inspect documents across all 6 major formats.

## Supported Formats

| Format | Extension | Type |
|--------|-----------|------|
| Word | `.docx` | OOXML |
| Excel | `.xlsx` | OOXML |
| PowerPoint | `.pptx` | OOXML |
| Word Legacy | `.doc` | Binary |
| Excel Legacy | `.xls` | Binary |
| PowerPoint Legacy | `.ppt` | Binary |

## Installation

```bash
# From crates.io
cargo install office_oxide_cli

# Pre-built binaries (via cargo-binstall)
cargo binstall office_oxide_cli

# From source
cargo install --path crates/office_oxide_cli
```

## Usage

```bash
# Extract plain text
office-oxide text document.docx

# Convert to markdown
office-oxide markdown spreadsheet.xlsx

# Show document info (format, size)
office-oxide info presentation.pptx

# Dump document IR as JSON
office-oxide ir document.docx
```

## Performance

Powered by [office_oxide](https://crates.io/crates/office_oxide) — the fastest Office document processing library in Rust. Opens 97.3% of a 10,386-file test collection.

## License

MIT OR Apache-2.0
