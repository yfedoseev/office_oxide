# office-oxide

The fastest Office document library for Python — DOCX, XLSX, PPTX, DOC, XLS, PPT parsing, conversion, and editing, powered by a Rust core.

- **Drop-in replacement** for text extraction workflows currently built on `python-docx`, `openpyxl`, `python-pptx`, `xlrd`. Up to 100× faster.
- **Format coverage**: DOCX, XLSX, PPTX (OOXML) + DOC, XLS, PPT (legacy OLE2).
- **Unified API**: `plain_text()`, `to_markdown()`, `to_html()`, `to_ir()` for every format.
- **Editing**: replace text in DOCX/PPTX, set cells in XLSX.
- **Packaging**: wheels for Linux, macOS, Windows (x64 + arm64), Python 3.8–3.14.

## Install

```bash
pip install office-oxide
```

## Usage

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

# Editing
with EditableDocument.open("template.docx") as ed:
    ed.replace_text("{{NAME}}", "Alice")
    ed.save("out.docx")

# Spreadsheet cells
with EditableDocument.open("data.xlsx") as ed:
    ed.set_cell(0, "A1", "Revenue")
    ed.set_cell(0, "B1", 12345.67)
    ed.set_cell(0, "C1", True)
    ed.save("data.edited.xlsx")
```

`Document.open` and related methods accept both `str` and `pathlib.Path`.

## Links

- Source code & issues: <https://github.com/yfedoseev/office_oxide>
- Other language bindings (Rust crate, Go, C#, Node native, WASM, raw C FFI): see the main repository.
- Licence: MIT OR Apache-2.0
