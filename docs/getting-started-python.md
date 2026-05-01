# Getting Started with office_oxide (Python)

`office_oxide` parses, converts, and edits Microsoft Office documents (DOCX, XLSX, PPTX, DOC, XLS, PPT) with a pure-Rust core exposed through a small, idiomatic Python API.

## Installation

```bash
pip install office-oxide
```

The PyPI distribution name uses a hyphen (`office-oxide`); the Python import is `office_oxide`. Wheels ship for CPython 3.8–3.14 on Linux, macOS, and Windows; there are no runtime dependencies.

## Quickstart

Extract plain text from a DOCX file:

```python
from office_oxide import Document

with Document.open("report.docx") as doc:
    print(doc.plain_text())
```

Or the one-shot module helper:

```python
import office_oxide
print(office_oxide.extract_text("report.docx"))
```

## Core API

### `Document`

`Document.open` detects the format from the file extension (and double-checks via magic bytes). It accepts `str`, `bytes`, or any `os.PathLike`. Use it as a context manager to release native memory deterministically.

```python
from pathlib import Path
from office_oxide import Document

with Document.open(Path("data/deck.pptx")) as doc:
    print(doc.format)           # "pptx"
    print(doc.plain_text())     # str
    print(doc.to_markdown())    # str
    print(doc.to_html())        # str
    ir = doc.to_ir()            # nested dict (see "Advanced")

    # Convert/save to any supported format — extension-driven.
    doc.save_as("deck.docx")    # legacy PPT → PPTX works too
```

Open from raw bytes when the file isn't on disk:

```python
data = open("report.xlsx", "rb").read()
with Document.from_bytes(data, "xlsx") as doc:
    print(doc.plain_text())
```

Module-level shortcuts:

```python
import office_oxide
office_oxide.extract_text("file.docx")   # → str
office_oxide.to_markdown("file.pptx")    # → str
office_oxide.to_html("file.xlsx")        # → str
office_oxide.version()                   # → "0.1.1"
```

### `EditableDocument`

Editing preserves all unmodified OPC parts (images, charts, styles) on save. Only DOCX, XLSX, and PPTX are editable.

```python
from office_oxide import EditableDocument

with EditableDocument.open("template.docx") as ed:
    n = ed.replace_text("{{name}}", "Alice")
    print(f"{n} replacements")
    ed.save("out.docx")
```

## Editing Examples

### Replace text in DOCX / PPTX

```python
from office_oxide import EditableDocument

with EditableDocument.open("slides.pptx") as ed:
    ed.replace_text("Q3", "Q4")
    ed.replace_text("2024", "2025")
    ed.save("slides_q4.pptx")
```

`replace_text` walks `<w:t>` elements in DOCX and `<a:t>` across every slide in PPTX. It returns the replacement count.

### Write XLSX cells

```python
from office_oxide import EditableDocument

with EditableDocument.open("budget.xlsx") as ed:
    ed.set_cell(0, "A1", "Total")     # string
    ed.set_cell(0, "B1", 42.5)        # number (int also accepted)
    ed.set_cell(0, "C1", True)        # boolean
    ed.set_cell(0, "D1", None)        # empty
    ed.save("budget.xlsx")
```

`sheet_index` is zero-based; `cell_ref` uses standard spreadsheet notation.

## Advanced

### Format-agnostic IR

`doc.to_ir()` returns a nested `dict` that mirrors the Rust `DocumentIR`: sections of headings, paragraphs, tables, lists, and images. Useful for building pipelines or feeding LLMs structured context.

```python
ir = doc.to_ir()
for section in ir["sections"]:
    print(section.get("title"))
    for el in section["elements"]:
        kind = el["kind"]  # "Heading" | "Paragraph" | "Table" | "List" | ...
        # ...
```

### Bytes-based workflow

`from_bytes` avoids writing temp files in serverless / streaming pipelines:

```python
import requests
from office_oxide import Document

data = requests.get("https://example.com/doc.docx").content
with Document.from_bytes(data, "docx") as doc:
    print(doc.to_markdown())
```

### Legacy formats (DOC, XLS, PPT)

The legacy CFB parsers are first-class. Extension detection routes automatically:

```python
with Document.open("legacy.doc") as doc:
    print(doc.plain_text())
    doc.save_as("legacy.docx")     # transparent DOC → DOCX
```

### Batch scripts

```python
from pathlib import Path
from office_oxide import Document

for path in Path("corpus").iterdir():
    if path.suffix.lower() in {".docx", ".xlsx", ".pptx", ".doc", ".xls", ".ppt"}:
        with Document.open(path) as doc:
            path.with_suffix(".txt").write_text(doc.plain_text())
        print(f"ok: {path.name}")
```

## Error Handling

All parse / IO failures raise `OfficeOxideError`. `save_as` wraps IO failures in `IOError`.

```python
from office_oxide import Document, OfficeOxideError

try:
    with Document.open("weird.file") as doc:
        print(doc.plain_text())
except OfficeOxideError as e:
    print(f"office_oxide failed: {e}")
except FileNotFoundError:
    print("no such file")
```

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| `OfficeOxideError: unsupported format: ""` | No extension on the path — use `Document.from_bytes(data, "docx")`. |
| `RuntimeError: Document is closed` | You exited the `with` block and still held a reference; open a fresh handle. |
| `ImportError: _native` | Wheel didn't match your platform; reinstall with `pip install --force-reinstall office_oxide`. |
| Legacy DOC renders as gibberish | File may be encrypted (Word 97 RC4). `office_oxide` does not decrypt; decrypt first via LibreOffice. |
| Unicode issues on Windows | Use `pathlib.Path` rather than byte paths; `Document.open` handles platform encoding. |

## Links

- Binding source: `src/python.rs`
- Python package layout: `python/office_oxide/`
- PyPI: https://pypi.org/project/office_oxide/
- Rust crate (the underlying core): https://crates.io/crates/office_oxide
- Module-level API reference: `python/office_oxide/_native.pyi`
