# office_oxide examples

The same ten scenarios in every binding, with the same data, the same
`step <n>: <what>` lines and the same labelled results, so you can open two
languages side by side and compare them line by line.

Every numbered example is **self-contained**: it builds its input in code,
writes only to the OS temp directory (and cleans up), needs no network, exits 0
on success, and prints **`OK <id>` as its last line** (for example
`OK 04_open_from_bytes`). On a failed check it prints `FAIL: <message>` to
stderr and exits non-zero. A run that does not end in `OK <id>` has failed.

## Scenarios

| id | what it shows |
|---|---|
| `01_extract` | open a DOCX from bytes; plain text, Markdown and IR JSON |
| `02_create_from_markdown` | one Markdown source into a DOCX, an XLSX and a PPTX |
| `03_edit` | replace text in a DOCX; set cells in an XLSX; save to bytes |
| `04_open_from_bytes` | DOCX, XLSX and PPTX built in memory and opened from bytes; a wrong format is an error |
| `05_xlsx_writer` | typed cells, a formula, a style, a merge, a column width; read back |
| `06_pptx_writer` | presentation size, slides, titles, text, an image; read back |
| `07_all_outputs` | one document as plain text, Markdown, HTML and IR JSON |
| `08_limits_strict_damage` | strict mode refuses a damaged archive; a lenient open reports the damage; a truncated archive and a per-document limit |
| `09_concurrency` | 32 documents read from 8 workers at once, with host logging on |
| `10_detect_and_batch` | detect a format from bytes, then process a mixed batch |

## Where each one lives

| id | Rust | Python | Go | C# | JavaScript | C |
|---|---|---|---|---|---|---|
| 01 | `rust/01_extract.rs` | `python/01_extract/main.py` | `go/01_extract/main.go` | `csharp/01_extract/Program.cs` | `javascript/01_extract.mjs` | `c/01_extract.c` |
| 02 | `rust/02_create_from_markdown.rs` | `python/02_create_from_markdown/main.py` | `go/02_create_from_markdown/main.go` | `csharp/02_create_from_markdown/Program.cs` | `javascript/02_create_from_markdown.mjs` | `c/02_create_from_markdown.c` |
| 03 | `rust/03_edit.rs` | `python/03_edit/main.py` | `go/03_edit/main.go` | `csharp/03_edit/Program.cs` | `javascript/03_edit.mjs` | `c/03_edit.c` |
| 04 | `rust/04_open_from_bytes.rs` | `python/04_open_from_bytes/main.py` | `go/04_open_from_bytes/main.go` | `csharp/04_open_from_bytes/Program.cs` | `javascript/04_open_from_bytes.mjs` | `c/04_open_from_bytes.c` |
| 05 | `rust/05_xlsx_writer.rs` | `python/05_xlsx_writer/main.py` | `go/05_xlsx_writer/main.go` | `csharp/05_xlsx_writer/Program.cs` | `javascript/05_xlsx_writer.mjs` | `c/05_xlsx_writer.c` |
| 06 | `rust/06_pptx_writer.rs` | `python/06_pptx_writer/main.py` | `go/06_pptx_writer/main.go` | `csharp/06_pptx_writer/Program.cs` | `javascript/06_pptx_writer.mjs` | `c/06_pptx_writer.c` |
| 07 | `rust/07_all_outputs.rs` | `python/07_all_outputs/main.py` | `go/07_all_outputs/main.go` | `csharp/07_all_outputs/Program.cs` | `javascript/07_all_outputs.mjs` | `c/07_all_outputs.c` |
| 08 | `rust/08_limits_strict_damage.rs` | `python/08_limits_strict_damage/main.py` | `go/08_limits_strict_damage/main.go` | `csharp/08_limits_strict_damage/Program.cs` | `javascript/08_limits_strict_damage.mjs` | `c/08_limits_strict_damage.c` |
| 09 | `rust/09_concurrency.rs` | `python/09_concurrency/main.py` | `go/09_concurrency/main.go` | `csharp/09_concurrency/Program.cs` | `javascript/09_concurrency.mjs` (step 2 skipped) | `c/09_concurrency.c` (step 2 skipped) |
| 10 | `rust/10_detect_and_batch.rs` | `python/10_detect_and_batch/main.py` | `go/10_detect_and_batch/main.go` | `csharp/10_detect_and_batch/Program.cs` | `javascript/10_detect_and_batch.mjs` | `c/10_detect_and_batch.c` |

Paths are relative to this directory. There are no other skipped cells apart
from the one step noted below.

**Intentionally skipped.** In `09_concurrency`, step 2 ("switch logging on")
prints `step 2: skipped, <reason>` in JavaScript and C, because neither
binding has a host logger to switch on. The point of the example, parsing from
many threads cannot deadlock the host, is still proven by steps 3 and 4.

**Two things every language does the same way, because the libraries do not
give a direct answer:**

- *DOCX bytes.* Only Rust can create a DOCX in memory
  (`create_from_markdown_to_writer`). Python, Go, JavaScript, C# and C create
  it from Markdown at a path, in a private temp file that is read back and
  deleted at once, so the DOCX content is identical everywhere.
- *Format detection from bytes.* No binding detects a format from bytes
  (`detect_format` goes by file extension). `10_detect_and_batch` sniffs the
  bytes the same way in all six: a `PK` prefix, then the first of the part
  names `word/document.xml`, `xl/workbook.xml` and `ppt/presentation.xml`
  found in the archive.

Error lines (`  error: <message>`) carry each binding's own message text, so
they are the one printed line that is not word-for-word the same.

## Running them

Build the native library first (every binding except pure Rust needs it):

```sh
cargo build --release --lib        # target/release/liboffice_oxide.so (.dylib, .dll)
```

Run from the repository root. Each command runs all ten.

```sh
# Rust (needs only a Rust toolchain)
for e in 01_extract 02_create_from_markdown 03_edit 04_open_from_bytes 05_xlsx_writer \
         06_pptx_writer 07_all_outputs 08_limits_strict_damage 09_concurrency 10_detect_and_batch; do
  timeout 120 cargo run --quiet --example "$e"
done

# Python (the office_oxide extension installed in the environment, e.g. maturin develop --release)
for d in examples/python/[0-9]*/; do timeout 120 python "${d}main.py"; done

# Go (go >= 1.21; the dev tag links ../target/release)
for d in examples/go/[0-9]*/; do (cd "$d" && timeout 120 go run -tags office_oxide_dev .); done

# C# (.NET 8 or later; the library directory must be on the loader path)
export LD_LIBRARY_PATH="$PWD/target/release"      # DYLD_LIBRARY_PATH on macOS, PATH on Windows
for p in examples/csharp/[0-9]*/*.csproj; do timeout 120 dotnet run --project "$p"; done

# JavaScript (Node; run `npm install` in js/ once for koffi)
export OFFICE_OXIDE_LIB="$PWD/target/release/liboffice_oxide.so"
for f in examples/javascript/[0-9]*.mjs; do timeout 120 node "$f"; done

# C (gcc or clang, POSIX: Linux and macOS; the Makefile wraps each run in timeout 120)
make -C examples/c run
```

To see that a run really finished, check the last line of its output:
`OK <id>`.

## Classic examples (take a file path)

Not part of the numbered set; they stay as they are.

| Scenario | Rust | Python | Go | C# | JavaScript | C |
|---|---|---|---|---|---|---|
| Extract plain text and Markdown | `rust/extract.rs` | `python/extract.py` | `go/extract/` | `csharp/Extract/` | `javascript/extract.mjs` | `c/extract.c` |
| Fill a template (text replace) | | `python/replace.py` | `go/replace/` | `csharp/Replace/` | `javascript/replace.mjs` | |
| Read a spreadsheet into rows | | `python/read_xlsx.py` | `go/read_xlsx/` | | `javascript/read_xlsx.mjs` | |

Other Rust helpers: `rust/create_rich.rs` (a rich DOCX with the `DocxWriter`;
prints `OK create_rich`), `rust/inspect_doc.rs` and `rust/make_smoke.rs`.
