# office_oxide — Mission

## What We're Building

**The fastest document processing suite in the world.** A Rust-native library for reading, writing, converting, and extracting content from all major office document formats — with first-class bindings to Python, Node.js, Go, C#, WASM, and C.

## Why

The LLM era demands fast, reliable document processing at scale:

- **LLM ingestion pipelines** process millions of documents daily. Every millisecond matters. Current tools (LibreOffice headless, Python libs) are 10–100× too slow.
- **AI-powered document generation** — reports, contracts, invoices — needs programmatic creation without heavyweight dependencies.
- **Privacy-first processing** — on-device, in-browser (WASM), no cloud round-trips. Enterprises won't send sensitive docs to third-party APIs.
- **Universal conversion** — the world runs on DOCX, XLSX, PPTX, and PDF. Moving between them should be instant and lossless.

## Current State (v0.1.1 — shipped)

All six Office formats are production-ready:

| Format | Read | Write | Edit | Text | Markdown | HTML | IR |
|--------|------|-------|------|------|----------|------|----|
| DOCX | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| XLSX | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| PPTX | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| DOC  | ✓ | — | — | ✓ | ✓ | ✓ | ✓ |
| XLS  | ✓ | — | — | ✓ | ✓ | ✓ | ✓ |
| PPT  | ✓ | — | — | ✓ | ✓ | ✓ | ✓ |

**98.4% pass rate on a 6,062-file public corpus. Zero failures on legitimate Office files.**

Language bindings: Rust, Python (PyPI), Go, C# / .NET (NuGet), Node.js native (npm), WASM (npm), C FFI.

## PDF ↔ Office Conversion

office_oxide is designed to serve as the Office output layer in document pipelines that process PDFs or other source formats.

### Markdown bridge (available today)

Any library that can emit Markdown can drive office_oxide's creation API directly:

```rust
// Markdown → DOCX (no intermediate serialization format needed)
let markdown = get_markdown_from_somewhere();
office_oxide::create::create_from_markdown(&markdown, DocumentFormat::Docx, "output.docx")?;
```

The same path works for PPTX and XLSX targets.

Going the other direction — Office → any text format — is equally direct:
`to_markdown()` / `to_html()` / `plain_text()` are available on every document type.

### DocumentIR bridge (lossless, planned)

The `DocumentIR` type in this crate is designed as a stable, serializable intermediate representation that other document libraries can target. The planned roadmap:

- Publish `DocumentIR` as a standalone crate (`document_ir`) so external libraries can depend on it without pulling in all of office_oxide.
- External libraries implement `to_ir() -> DocumentIR` to produce structurally faithful output (headings, tables, lists preserved as typed elements, not text).
- `office_oxide::create::create_from_ir()` consumes that IR to produce native Office documents.

This architecture achieves lossless, zero-copy format conversion without serialising through text.

## Design Principles

1. **Speed above all** — zero-copy parsing, memory-mapped I/O, Rayon for parallel workloads
2. **Correctness** — spec-compliant (ISO 29500 / ECMA-376), tested on thousands of real-world documents, zero panics
3. **Minimal dependencies** — no C/C++ runtime, no system library requirements, pure Rust core
4. **Developer experience** — unified API, serde integration, comprehensive docs and per-language examples
5. **LLM-first output** — clean Markdown, structured IR, chunk-friendly text for RAG pipelines

## Target Users

1. **AI/LLM companies** — document ingestion pipelines (RAG, training data)
2. **SaaS platforms** — document preview, conversion, processing features
3. **Enterprise** — on-premise document processing, compliance, archival
4. **Developers** — anyone who needs to work with office documents programmatically

## License

MIT OR Apache-2.0 — no AGPL, no GPL, no copyleft restrictions.
