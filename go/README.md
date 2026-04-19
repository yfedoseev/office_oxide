# office_oxide — Go binding

Idiomatic Go bindings for [office_oxide](https://github.com/yfedoseev/office_oxide),
a fast Rust library for parsing, converting, and editing Office documents
(DOCX, XLSX, PPTX, DOC, XLS, PPT).

## Install

```bash
go get github.com/yfedoseev/office_oxide/go
```

The binding uses cgo and links against the `office_oxide` Rust static library.
You must either:

1. **Build locally from the monorepo.** From the office_oxide repo root:

   ```bash
   cargo build --release --lib
   go test -tags office_oxide_dev ./go/...
   ```

   The `office_oxide_dev` build tag points cgo at `target/release/`.

2. **Point cgo at an installed prefix.** Build the library once, place
   `libofice_oxide.a` (or `.so` / `.dylib`) somewhere on your link path,
   and the header somewhere on your include path, then set:

   ```bash
   export CGO_CFLAGS="-I/path/to/include"
   export CGO_LDFLAGS="-L/path/to/lib -loffice_oxide"
   ```

## Quick start

```go
package main

import (
    "fmt"
    "log"

    oo "github.com/yfedoseev/office_oxide/go"
)

func main() {
    doc, err := oo.Open("report.docx")
    if err != nil {
        log.Fatal(err)
    }
    defer doc.Close()

    fmt.Println("format:", mustStr(doc.Format()))
    fmt.Println(mustStr(doc.PlainText()))

    md, err := doc.ToMarkdown()
    if err != nil {
        log.Fatal(err)
    }
    fmt.Println(md)
}

func mustStr(s string, err error) string {
    if err != nil {
        log.Fatal(err)
    }
    return s
}
```

## Editing

```go
ed, err := oo.OpenEditable("template.docx")
if err != nil { log.Fatal(err) }
defer ed.Close()

if _, err := ed.ReplaceText("{{NAME}}", "Alice"); err != nil {
    log.Fatal(err)
}
if err := ed.Save("out.docx"); err != nil {
    log.Fatal(err)
}
```

## API

| Function / method | Description |
| --- | --- |
| `Version() string` | Library version. |
| `DetectFormat(path) string` | Short format name (`"docx"`, …) or `""`. |
| `Open(path) (*Document, error)` | Open from a file path. |
| `OpenFromBytes(data, format) (*Document, error)` | Open from memory. |
| `Document.Close() error` | Release the handle (safe to defer). |
| `Document.Format() (string, error)` | Detected format. |
| `Document.PlainText() (string, error)` | Extract plain text. |
| `Document.ToMarkdown() (string, error)` | Convert to Markdown. |
| `Document.ToHTML() (string, error)` | Convert to an HTML fragment. |
| `Document.ToIRJSON() (string, error)` | Format-agnostic IR as JSON. |
| `Document.SaveAs(path) error` | Save/convert to a new format. |
| `OpenEditable(path) (*EditableDocument, error)` | Open for editing (DOCX/XLSX/PPTX). |
| `EditableDocument.ReplaceText(find, replace) (int64, error)` | In-place replace. |
| `EditableDocument.SetCell(sheet, ref, value) error` | Write an XLSX cell. |
| `EditableDocument.Save(path) error` | Persist to disk. |
| `EditableDocument.SaveToBytes() ([]byte, error)` | Persist to memory. |
| `ExtractText(path)` / `ToMarkdown(path)` / `ToHTML(path)` | One-shot helpers. |

Errors returned from FFI are `*Error` values that expose the underlying
code via `Error.Code` and the originating operation via `Error.Op`.

## License

MIT OR Apache-2.0
