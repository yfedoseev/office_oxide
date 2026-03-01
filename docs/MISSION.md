# office_oxide — Mission

## What We're Building

**The fastest document processing suite in the world.** A Rust-native library for reading, writing, converting, and extracting content from all major office document formats — with bindings to Python, Node.js, WASM, and C.

## Why

The LLM era demands fast, reliable document processing at scale:

- **LLM ingestion pipelines** process millions of documents daily. Every millisecond matters. Current tools (LibreOffice headless, Python libs) are 10–100x too slow.
- **AI-powered document generation** — reports, contracts, invoices — needs programmatic creation without heavyweight dependencies.
- **Privacy-first processing** — on-device, in-browser (WASM), no cloud round-trips. Enterprises won't send sensitive docs to third-party APIs.
- **Universal conversion** — the world runs on DOCX, XLSX, PPTX, and PDF. Moving between them should be instant and lossless.

## The Problem Today

| Tool | Speed | Quality | Formats | Deployment |
|------|-------|---------|---------|------------|
| LibreOffice headless | Slow (200ms+) | Good | All | Heavy (500MB+) |
| python-docx/openpyxl | Slow | Partial | One each | Python only |
| Apache POI (Java) | Medium | Good | OOXML | JVM dependency |
| Aspose (Commercial) | Medium | Excellent | All | Expensive, closed |
| **office_oxide** | **<10ms** | **Excellent** | **All** | **Zero deps, 5MB** |

## Proof

We already proved this with `pdf_oxide`:

- **47.9x faster** than PyMuPDF4LLM
- **2.1ms mean latency** per document
- **100% pass rate** on 3,830 real-world PDFs
- **99.6% clean text extraction** rate
- Zero panics, zero crashes

The same Rust-native, zero-copy, SIMD-optimized approach applies to OOXML formats.

## What We Ship

### Crates

| Crate | Purpose | Status |
|-------|---------|--------|
| `pdf_oxide` | PDF read/write/convert/extract | Production (v0.3.7) |
| `office_core` | Shared OPC/ZIP/XML, themes, styles | Planned |
| `docx_oxide` | Word documents (.docx) | Planned |
| `xlsx_oxide` | Excel spreadsheets (.xlsx) | Planned |
| `pptx_oxide` | PowerPoint presentations (.pptx) | Planned |
| `office_oxide` | Unified API + language bindings | Planned |

### Capabilities Per Format

| Capability | PDF | DOCX | XLSX | PPTX |
|-----------|-----|------|------|------|
| **Read/Parse** | Done | Planned | Planned | Planned |
| **Extract Text** | Done | Planned | Planned | Planned |
| **Extract Images** | Done | Planned | Planned | Planned |
| **Create** | Done | Planned | Planned | Planned |
| **Edit** | Done | Planned | Planned | Planned |
| **Convert to PDF** | N/A | Planned | Planned | Planned |
| **Convert to Markdown** | Done | Planned | Planned | Planned |
| **Convert to HTML** | Done | Planned | Planned | Planned |
| **Convert to JSON** | — | — | Planned | — |

### Language Bindings

| Language | Mechanism | Status |
|----------|-----------|--------|
| Rust | Native | Done (pdf_oxide) |
| Python | PyO3 | Done (pdf_oxide) |
| WASM/JS | wasm-bindgen | Done (pdf_oxide) |
| Node.js | napi-rs | Planned |
| C/C++ | cbindgen FFI | Planned |
| Ruby | FFI | Planned |
| Go | cgo | Planned |
| Java | JNI | Planned |

## Design Principles

### 1. Speed Above All
- Zero-copy parsing where possible
- Memory-mapped I/O for large files
- Streaming processing — constant memory regardless of file size
- SIMD-accelerated text processing
- Parallel processing via Rayon where beneficial

### 2. Correctness
- Spec-compliant (ISO 29500, ECMA-376)
- Tested against thousands of real-world documents
- Fuzz-tested for robustness
- No panics — all errors are recoverable `Result` types

### 3. Minimal Dependencies
- No C/C++ dependencies in the core path
- No system library requirements (no libxml2, no zlib system dep)
- Pure Rust where possible
- Small binary size

### 4. Developer Experience
- Unified API across all formats: `office_oxide::extract_text("file.docx")`
- Builder pattern for document creation
- Serde integration for serialization
- Comprehensive documentation and examples

### 5. LLM-First
- Optimized text extraction that produces clean, structured output
- Markdown conversion as a first-class citizen
- Metadata extraction (author, dates, properties)
- Table extraction to structured data (JSON, CSV)
- Chunk-friendly output for RAG pipelines

## Target Users

1. **AI/LLM companies** — document ingestion pipelines (RAG, training data)
2. **SaaS platforms** — document preview, conversion, processing features
3. **Enterprise** — on-premise document processing, compliance, archival
4. **Developers** — anyone who needs to work with office documents programmatically

## Competitive Moat

1. **Rust performance** — 10-100x faster than alternatives, can't be matched by Python/Java
2. **Unified suite** — one dependency for all formats, not 4 separate libraries
3. **Universal bindings** — works in every language ecosystem
4. **WASM support** — runs in the browser, no server needed
5. **pdf_oxide track record** — proven we can deliver SOTA quality and performance

## Success Metrics

- **Latency**: <10ms for text extraction on typical documents across all formats
- **Correctness**: >99% fidelity on real-world document corpora
- **Compatibility**: Handle documents from MS Office, Google Docs, LibreOffice, Apple Pages/Numbers/Keynote exports
- **Adoption**: Become the default document processing library for Rust and Python AI ecosystems

## Roadmap

### Phase 1 — Foundation (Current)
- [x] pdf_oxide — production ready
- [ ] Collect format specifications
- [ ] Set up workspace and CI
- [ ] office_core — OPC, ZIP, XML shared primitives

### Phase 2 — DOCX
- [ ] docx_oxide read/parse
- [ ] Text extraction
- [ ] DOCX → PDF conversion (using pdf_oxide writer)
- [ ] DOCX → Markdown/HTML conversion
- [ ] DOCX creation from scratch
- [ ] Python bindings

### Phase 3 — XLSX
- [ ] xlsx_oxide read/parse
- [ ] Cell/formula extraction
- [ ] XLSX → CSV/JSON conversion
- [ ] XLSX → PDF conversion
- [ ] XLSX creation
- [ ] Python bindings

### Phase 4 — PPTX
- [ ] pptx_oxide read/parse
- [ ] Slide text/image extraction
- [ ] PPTX → PDF conversion
- [ ] PPTX → image export (per slide)
- [ ] Python bindings

### Phase 5 — Unified API & Polish
- [ ] office_oxide unified API
- [ ] Node.js bindings (napi-rs)
- [ ] C ABI
- [ ] Comprehensive benchmarks vs competitors
- [ ] Documentation site

## License

MIT OR Apache-2.0 (same as pdf_oxide)
