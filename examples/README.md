# office_oxide examples

Identical scenarios implemented against every language binding. Each
language directory ships the same three demos:

| Scenario | File stem |
| --- | --- |
| Extract plain text + Markdown | `extract` |
| Fill a template (text replace) | `replace` |
| Read a spreadsheet into rows | `read_xlsx` |

## Layout

| Binding | Directory |
| --- | --- |
| Python | [`python/`](python/) |
| Go | [`go/`](go/) |
| C# / .NET | [`csharp/`](csharp/) |
| JavaScript (native, koffi) | [`javascript/`](javascript/) |
| WASM | [`../wasm-pkg/`](../wasm-pkg/) README |
| Rust | [`rust/`](rust/) |
| Raw C FFI | [`c/`](c/) |

All examples read an input path from `argv`/`os.Args`, so you can point them
at a docx/xlsx/pptx file of your choice. The Rust monorepo ships a tiny
docx at `examples/rust/make_smoke.rs` to produce a fixture.
