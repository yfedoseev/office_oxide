# Getting Started with office_oxide (Go)

The Go package `github.com/yfedoseev/office_oxide/go` wraps the office_oxide C FFI via cgo and gives you idiomatic Go types for reading, converting, and editing DOCX / XLSX / PPTX / DOC / XLS / PPT files.

## Installation

```bash
go get github.com/yfedoseev/office_oxide/go@latest
```

The binding needs the `office_oxide` C library at link time. Two ways to provide it:

**Option 1 — one-liner installer (recommended):**

```bash
go run github.com/yfedoseev/office_oxide/go/cmd/install@latest
```

The installer downloads the matching prebuilt `liboffice_oxide` (and header) for your OS/arch and places it where cgo can find it. Re-run after upgrading to pick up the new ABI.

**Option 2 — set cgo flags yourself** if you've built the library from the Rust source (`cargo build --release --lib`) or have it in a system prefix:

```bash
export CGO_CFLAGS="-I/usr/local/include"
export CGO_LDFLAGS="-L/usr/local/lib -loffice_oxide"
```

## Quickstart

Extract plain text from a DOCX file:

```go
package main

import (
    "fmt"
    "log"

    officeoxide "github.com/yfedoseev/office_oxide/go"
)

func main() {
    doc, err := officeoxide.Open("report.docx")
    if err != nil {
        log.Fatal(err)
    }
    defer doc.Close()

    text, err := doc.PlainText()
    if err != nil {
        log.Fatal(err)
    }
    fmt.Println(text)
}
```

Or the one-shot helper:

```go
text, err := officeoxide.ExtractText("report.docx")
```

## Core API

### `Document`

```go
doc, err := officeoxide.Open("file.xlsx")
if err != nil { log.Fatal(err) }
defer doc.Close()

format, _ := doc.Format()      // "xlsx"
text, _   := doc.PlainText()
md, _     := doc.ToMarkdown()
html, _   := doc.ToHTML()
irJSON, _ := doc.ToIRJSON()

// Save/convert — target format is inferred from the extension.
err = doc.SaveAs("file.docx")
```

`Document` wraps a C handle; a finalizer frees it if you forget `Close`, but always prefer `defer doc.Close()` for deterministic cleanup.

Open from bytes (useful for streaming / serverless):

```go
data, _ := os.ReadFile("report.pptx")
doc, err := officeoxide.OpenFromBytes(data, "pptx")
```

`format` must be one of `"docx"`, `"xlsx"`, `"pptx"`, `"doc"`, `"xls"`, `"ppt"`.

### `EditableDocument`

Editable handles preserve every unmodified OPC part (images, charts, custom XML) on save. DOCX, XLSX, and PPTX only.

```go
ed, err := officeoxide.OpenEditable("template.docx")
if err != nil { log.Fatal(err) }
defer ed.Close()

n, _ := ed.ReplaceText("{{name}}", "Alice")
fmt.Printf("%d replacements\n", n)

err = ed.Save("out.docx")
```

## Editing Examples

### Replace text across a DOCX or PPTX

```go
ed, _ := officeoxide.OpenEditable("slides.pptx")
defer ed.Close()

ed.ReplaceText("Q3", "Q4")
ed.ReplaceText("2024", "2025")

buf, _ := ed.SaveToBytes()   // []byte ready to upload / stream
_ = os.WriteFile("slides_q4.pptx", buf, 0o644)
```

### Set cells in an XLSX

```go
ed, _ := officeoxide.OpenEditable("budget.xlsx")
defer ed.Close()

ed.SetCell(0, "A1", officeoxide.NewStringCell("Total"))
ed.SetCell(0, "B1", officeoxide.NewNumberCell(42.5))
ed.SetCell(0, "C1", officeoxide.NewBoolCell(true))
ed.SetCell(0, "D1", officeoxide.NewEmptyCell())

ed.Save("budget.xlsx")
```

Use `NewStringCell`, `NewNumberCell`, `NewBoolCell`, or `NewEmptyCell` — the constructor picks the correct variant for the FFI call.

## Advanced

### Format-agnostic IR

`doc.ToIRJSON()` returns JSON that matches the Rust `DocumentIR` schema. Unmarshal into whatever shape you need:

```go
import "encoding/json"

irJSON, _ := doc.ToIRJSON()

var ir struct {
    Sections []struct {
        Title    *string `json:"title"`
        Elements []json.RawMessage `json:"elements"`
    } `json:"sections"`
}
_ = json.Unmarshal([]byte(irJSON), &ir)
fmt.Printf("%d sections\n", len(ir.Sections))
```

### Detect format without opening

```go
fmt := officeoxide.DetectFormat("mystery.bin")  // "" if unsupported
```

### Legacy formats (DOC, XLS, PPT)

Open them exactly like OOXML; `SaveAs` transparently converts through the IR:

```go
doc, _ := officeoxide.Open("old.xls")
defer doc.Close()
_ = doc.SaveAs("modern.xlsx")
```

## Error Handling

Every fallible call returns an `*officeoxide.Error` carrying a typed code plus the originating operation:

```go
if _, err := officeoxide.Open("missing.docx"); err != nil {
    var e *officeoxide.Error
    if errors.As(err, &e) {
        fmt.Printf("code=%d op=%s\n", e.Code, e.Op)
    }
}
```

Using a closed handle returns `officeoxide.ErrClosed`.

### Error codes

| Code | Name | Meaning |
|---:|---|---|
| 0 | `OK` | success |
| 1 | `INVALID_ARG` | nil / empty / wrong format string |
| 2 | `IO` | filesystem error |
| 3 | `PARSE` | malformed document |
| 4 | `EXTRACTION` | parsing succeeded but rendering failed |
| 5 | `INTERNAL` | bug — please file an issue |
| 6 | `UNSUPPORTED` | extension / feature not supported |

## Troubleshooting

| Symptom | Fix |
|---|---|
| `could not determine kind of name for C.office_document_open` | The C headers aren't visible to cgo. Run the installer (`go run .../cmd/install@latest`) or set `CGO_CFLAGS`. |
| `cannot find -loffice_oxide` at link time | Set `CGO_LDFLAGS="-L/path/to/lib -loffice_oxide"` or run the installer. |
| Runtime `cannot open shared object file` | Add the library directory to `LD_LIBRARY_PATH` (Linux), `DYLD_LIBRARY_PATH` (macOS), or copy the DLL next to your binary (Windows). |
| `unsupported format` on `.doc`/`.xls` | Make sure the extension is lowercase; the binding routes via it. Or call `OpenFromBytes(data, "doc")`. |
| Cross-compilation fails | cgo requires a target-matching C toolchain; use `zig cc` or set `CC=aarch64-linux-gnu-gcc` etc. |

## Links

- Binding source: `go/office_oxide.go`
- C header: `include/office_oxide_c/office_oxide.h`
- Installer: `go/cmd/install/`
- Package on pkg.go.dev: https://pkg.go.dev/github.com/yfedoseev/office_oxide/go
- GitHub: https://github.com/yfedoseev/office_oxide
