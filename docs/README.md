# office_oxide documentation

## Getting started — per language

| Language | Guide |
| --- | --- |
| Rust | [getting-started-rust.md](getting-started-rust.md) |
| Python | [getting-started-python.md](getting-started-python.md) |
| Go | [getting-started-go.md](getting-started-go.md) |
| C# / .NET | [getting-started-csharp.md](getting-started-csharp.md) |
| JavaScript (Node native) | [getting-started-javascript.md](getting-started-javascript.md) |
| JavaScript / browser (WASM) | [getting-started-wasm.md](getting-started-wasm.md) |
| C / raw FFI | [getting-started-c.md](getting-started-c.md) |

Each guide covers install, the core API (`Document`, `EditableDocument`),
editing, advanced usage (IR, in-memory parsing, legacy format handling),
and error handling.

## Runnable examples

Identical demos (`extract`, `replace`, `read_xlsx`) for every binding live
in [`../examples/`](../examples/). Start there to see a concrete script for
your language.

## Deeper references

- [ARCHITECTURE.md](ARCHITECTURE.md) — how the Rust core is organised
  (format modules, the common OPC reader, the IR converters, the editing
  layer).
- [MISSION.md](MISSION.md) — project goals and non-goals.
- [architecture/](architecture/), [specs/](specs/) — internal deep-dives
  (format notes, parsing invariants, benchmark methodology).

## C FFI reference

The stable C ABI is documented inline in
[`../include/office_oxide_c/office_oxide.h`](../include/office_oxide_c/office_oxide.h).
Every language binding in this repo is built on top of that surface.
