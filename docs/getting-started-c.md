# Getting Started with office_oxide (C FFI)

The office_oxide C API is a stable, thin ABI over the Rust core. All higher-level bindings (Go, C#, Node.js native) call these same entry points, so if an ecosystem doesn't have a ready-made wrapper you can always link the C library directly.

The authoritative surface lives in `include/office_oxide_c/office_oxide.h`; this guide walks through the same API.

## Installation

### Option 1 — build from the Rust source

```bash
git clone https://github.com/yfedoseev/office_oxide
cd office_oxide
cargo build --release --lib
```

You'll find the shared library under `target/release/`:

| OS | File |
|---|---|
| Linux | `liboffice_oxide.so` |
| macOS | `liboffice_oxide.dylib` |
| Windows | `office_oxide.dll` (+ `.lib` import library) |

The header lives at `include/office_oxide_c/office_oxide.h`.

### Option 2 — install a prebuilt library

Copy the header and shared library into a prefix of your choice:

```bash
install -Dm644 include/office_oxide_c/office_oxide.h /usr/local/include/office_oxide.h
install -Dm755 target/release/liboffice_oxide.so     /usr/local/lib/liboffice_oxide.so
```

## Quickstart

Extract plain text from a DOCX file:

```c
#include <stdio.h>
#include <stdlib.h>
#include "office_oxide.h"

int main(void) {
    int err = 0;
    char *text = office_extract_text("report.docx", &err);
    if (!text) {
        fprintf(stderr, "office_oxide failed: code=%d\n", err);
        return 1;
    }
    printf("%s\n", text);
    office_oxide_free_string(text);
    return 0;
}
```

Build and run:

```bash
cc quickstart.c -I/usr/local/include -L/usr/local/lib -loffice_oxide -o quickstart
LD_LIBRARY_PATH=/usr/local/lib ./quickstart
```

## Core API

### Library info

```c
const char *version = office_oxide_version();           // "0.1.1" — don't free
const char *fmt     = office_oxide_detect_format("f"); // "docx"/... or NULL
```

### Document (read-only)

```c
int err = 0;
OfficeDocumentHandle *doc = office_document_open("file.xlsx", &err);
if (!doc) { /* handle err */ }

const char *fmt = office_document_format(doc);         // "xlsx" — don't free

char *text = office_document_plain_text(doc, &err);
char *md   = office_document_to_markdown(doc, &err);
char *html = office_document_to_html(doc, &err);
char *ir   = office_document_to_ir_json(doc, &err);

// Save/convert — format inferred from extension.
if (office_document_save_as(doc, "file.docx", &err) != 0) {
    /* handle err */
}

// Free every heap string from the API:
office_oxide_free_string(text);
office_oxide_free_string(md);
office_oxide_free_string(html);
office_oxide_free_string(ir);

office_document_free(doc);
```

Open from an in-memory buffer:

```c
uint8_t *data = ...;
size_t   len  = ...;
OfficeDocumentHandle *doc =
    office_document_open_from_bytes(data, len, "docx", &err);
```

`format` must be one of the six strings `"docx" | "xlsx" | "pptx" | "doc" | "xls" | "ppt"`.

### EditableDocument

Editing preserves every unmodified OPC part on save. Only DOCX, XLSX, and PPTX are supported.

```c
int err = 0;
OfficeEditableHandle *ed = office_editable_open("template.docx", &err);
if (!ed) { /* handle err */ }

int64_t n = office_editable_replace_text(ed, "{{name}}", "Alice", &err);
if (n < 0) { /* handle err */ }
printf("%lld replacements\n", (long long)n);

if (office_editable_save(ed, "out.docx", &err) != 0) {
    /* handle err */
}

office_editable_free(ed);
```

Set XLSX cells:

```c
OfficeEditableHandle *wb = office_editable_open("budget.xlsx", &err);

office_editable_set_cell(wb, 0, "A1", OFFICE_CELL_STRING, "Total", 0.0,  &err);
office_editable_set_cell(wb, 0, "B1", OFFICE_CELL_NUMBER,  NULL,   42.5, &err);
office_editable_set_cell(wb, 0, "C1", OFFICE_CELL_BOOLEAN, NULL,   1.0,  &err);
office_editable_set_cell(wb, 0, "D1", OFFICE_CELL_EMPTY,   NULL,   0.0,  &err);

office_editable_save(wb, "budget.xlsx", &err);
office_editable_free(wb);
```

For booleans, non-zero `value_num` = true.

Serialize to a heap buffer instead of a file:

```c
size_t   out_len = 0;
uint8_t *buf     = office_editable_save_to_bytes(ed, &out_len, &err);
if (!buf) { /* handle err */ }
/* ... upload or stream buf[0..out_len] ... */
office_oxide_free_bytes(buf, out_len);
```

### One-shot convenience helpers

```c
char *text = office_extract_text("file.docx", &err);
char *md   = office_to_markdown("file.pptx", &err);
char *html = office_to_html("file.xlsx", &err);
/* free each with office_oxide_free_string */
```

## Memory rules (important)

- Strings returned as `char*` from `office_document_*`, `office_editable_*`, and the one-shot helpers must be freed with `office_oxide_free_string(ptr)`.
- Byte buffers returned as `uint8_t*` (with an `out_len` out-parameter) must be freed with `office_oxide_free_bytes(ptr, len)`. The `len` must match the one the API wrote to `out_len`.
- Opaque handles must be freed with their matching `*_free()` call.
- `const char*` values returned by `office_oxide_version`, `office_oxide_detect_format`, and `office_document_format` are static — **do not free**.

## Error handling

Every call that can fail takes an `int *error_code` out-parameter:

```c
int err = 0;
OfficeDocumentHandle *doc = office_document_open("missing.docx", &err);
if (!doc) {
    switch (err) {
        case OFFICE_ERR_IO:       fputs("io error\n", stderr); break;
        case OFFICE_ERR_PARSE:    fputs("parse error\n", stderr); break;
        /* ... */
    }
}
```

### Error codes

| Macro | Value | Meaning |
|---|---:|---|
| `OFFICE_OK` | 0 | success |
| `OFFICE_ERR_INVALID_ARG` | 1 | nil pointer / unknown format string |
| `OFFICE_ERR_IO` | 2 | filesystem error |
| `OFFICE_ERR_PARSE` | 3 | malformed document |
| `OFFICE_ERR_EXTRACTION` | 4 | parsed but rendering failed |
| `OFFICE_ERR_INTERNAL` | 5 | bug — please file an issue |
| `OFFICE_ERR_UNSUPPORTED` | 6 | extension / feature not supported |

### Cell value constants

| Macro | Value |
|---|---:|
| `OFFICE_CELL_EMPTY` | 0 |
| `OFFICE_CELL_STRING` | 1 |
| `OFFICE_CELL_NUMBER` | 2 |
| `OFFICE_CELL_BOOLEAN` | 3 |

## Advanced

### Legacy formats (DOC, XLS, PPT)

Open with `office_document_open`; the format is detected from the extension and verified via magic bytes. `office_document_save_as` can convert legacy → OOXML transparently:

```c
OfficeDocumentHandle *doc = office_document_open("old.xls", &err);
office_document_save_as(doc, "modern.xlsx", &err);
office_document_free(doc);
```

### Thread safety

Each handle is owned by the caller — do not share a single handle between threads. Different handles can be used in parallel; the library itself is reentrant.

### C++

The header is wrapped in `extern "C"` guards, so it drops straight into a C++ translation unit:

```cpp
#include "office_oxide.h"
```

Pair with RAII wrappers (e.g. `std::unique_ptr<OfficeDocumentHandle, decltype(&office_document_free)>`) to make cleanup exception-safe.

## Troubleshooting

| Symptom | Fix |
|---|---|
| Linker: `undefined reference to office_document_open` | Add `-loffice_oxide` and a matching `-L` flag; ensure the library was built with `--release --lib`. |
| Runtime: `error while loading shared libraries: liboffice_oxide.so` | Add the library directory to `LD_LIBRARY_PATH` (Linux), `DYLD_LIBRARY_PATH` (macOS), or next to the exe (Windows). |
| `OFFICE_ERR_INVALID_ARG` on `open_from_bytes` | `format` must be exactly `"docx"|"xlsx"|"pptx"|"doc"|"xls"|"ppt"`, lowercase. |
| Double-free or heap corruption | Check that each `char*` was freed with `office_oxide_free_string`, and each byte buffer with `office_oxide_free_bytes(ptr, len)` using the original `out_len`. |
| `office_editable_replace_text` returns 0 on XLSX | Expected — use `office_editable_set_cell` for spreadsheet edits. |

## Links

- C header: `include/office_oxide_c/office_oxide.h`
- FFI implementation: `src/ffi/`
- Rust crate: https://crates.io/crates/office_oxide
- GitHub: https://github.com/yfedoseev/office_oxide
- Higher-level bindings: [Rust](getting-started-rust.md), [Python](getting-started-python.md), [Go](getting-started-go.md), [C#](getting-started-csharp.md), [Node.js](getting-started-javascript.md), [WASM](getting-started-wasm.md)
