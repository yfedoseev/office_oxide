# Competitive Landscape

## The Document Processing Problem

Every AI company, SaaS platform, and enterprise needs to process office documents. The existing options all have critical weaknesses that office_oxide can exploit.

---

## Current Solutions

### LibreOffice Headless / UNO API
- **Languages**: C++, Python bindings
- **Formats**: All (best compatibility)
- **Strengths**: Most complete format support, battle-tested
- **Weaknesses**:
  - **Slow**: 200ms–2s per document conversion
  - **Heavy**: 500MB+ install, requires X11/headless display
  - **Fragile**: Crashes under concurrent load, memory leaks
  - **Not embeddable**: Can't ship as a library
  - **Process-based**: Must spawn a process per conversion
- **Used by**: Collabora, OnlyOffice (partially), many SaaS backends

### python-docx / openpyxl / python-pptx
- **Language**: Python
- **Formats**: One each (DOCX / XLSX / PPTX only)
- **Strengths**: Easy API, well-documented, large community
- **Weaknesses**:
  - **Slow**: Pure Python, 50–500ms per document
  - **Read-heavy**: Write support is limited/buggy
  - **No conversion**: Can't convert between formats
  - **No PDF output**: Need separate tool chain
  - **Three separate libraries**: No unified API
- **Used by**: Most Python projects, small-scale processing

### Apache POI (Java)
- **Language**: Java
- **Formats**: OOXML + legacy binary (doc/xls/ppt)
- **Strengths**: Mature, good spec compliance, legacy format support
- **Weaknesses**:
  - **JVM dependency**: 100MB+ runtime
  - **Memory hungry**: Loads entire documents into memory
  - **Medium speed**: Faster than Python, slower than native
  - **No conversion**: Read/write only, no format conversion
- **Used by**: Enterprise Java shops, Spring-based systems

### Aspose (Commercial, .NET/Java)
- **Language**: .NET, Java
- **Formats**: All formats, excellent coverage
- **Strengths**: Best-in-class format fidelity, active development
- **Weaknesses**:
  - **Expensive**: $1,000–15,000+ per developer/year
  - **Closed source**: No customization, vendor lock-in
  - **Runtime dependency**: .NET or JVM required
  - **Licensing complexity**: Per-developer, per-deployment
- **Used by**: Enterprise, regulated industries

### Pandoc
- **Language**: Haskell
- **Formats**: Many text formats (Markdown, HTML, LaTeX, DOCX)
- **Strengths**: Excellent text conversion, extensible via filters
- **Weaknesses**:
  - **Not a library**: Command-line tool, hard to embed
  - **No spreadsheets/presentations**: Text-focused only
  - **Haskell runtime**: Non-trivial deployment
  - **Slow for batch**: Process spawn overhead
- **Used by**: Academic writing, documentation pipelines

### docx-rs / calamine / umya-spreadsheet (Rust)
- **Language**: Rust
- **Formats**: Individual format support
- **Strengths**: Rust performance, some existing work
- **Weaknesses**:
  - **Incomplete**: Missing many features, edge cases
  - **No conversion**: Read/write only
  - **Fragmented**: Separate libraries, no unified vision
  - **Low activity**: Some are unmaintained
  - **No text extraction focus**: Not optimized for LLM use cases
- **Used by**: Niche Rust projects

### Unstructured.io
- **Language**: Python (wraps many tools)
- **Formats**: All (via LibreOffice, Tesseract, etc.)
- **Strengths**: LLM-focused, good chunking, metadata extraction
- **Weaknesses**:
  - **Slow**: Wraps slow tools, adds overhead
  - **Heavy dependencies**: Requires LibreOffice, Tesseract, etc.
  - **Cloud-first**: Self-hosting is complex
  - **Expensive**: Commercial pricing for scale
- **Used by**: AI/LLM companies for RAG pipelines

### Marker (Python)
- **Language**: Python
- **Formats**: PDF, DOCX, PPTX (via conversion)
- **Strengths**: Good PDF→Markdown, ML-powered layout
- **Weaknesses**:
  - **Slow**: Uses ML models, 1–10s per page
  - **GPU dependent**: Needs CUDA for reasonable speed
  - **PDF-focused**: Other formats are converted to PDF first (lossy)
- **Used by**: AI researchers, RAG pipelines

---

## Our Advantages

| Dimension | Competitors | office_oxide |
|-----------|------------|--------------|
| **Speed** | 50ms–2s per doc | **<10ms per doc** |
| **Memory** | 100MB–1GB | **<10MB typical** |
| **Dependencies** | Heavy (JVM, .NET, LibreOffice) | **Zero system deps** |
| **Deployment** | Complex | **Single binary / pip install** |
| **Formats** | Usually 1-2 each | **All major formats, one library** |
| **Conversion** | Limited or via LibreOffice | **Native, any-to-any** |
| **LLM-optimized** | Afterthought | **First-class** |
| **WASM** | None | **Runs in browser** |
| **Cost** | $0–$15K/yr | **Free, open source** |
| **Language support** | Usually 1-2 | **Rust, Python, JS, WASM, C, Go, Java** |

## Market Opportunity

### Primary: LLM/AI Document Ingestion
- Every RAG pipeline needs document processing
- Speed directly impacts cost (fewer servers) and latency (better UX)
- Market growing exponentially with LLM adoption

### Secondary: SaaS Document Features
- Preview, convert, extract — every SaaS eventually needs this
- Currently solved with LibreOffice (painful) or Aspose (expensive)
- WASM support enables client-side processing (privacy, speed)

### Tertiary: Developer Tools
- Replace fragmented Python libs with one fast dependency
- CLI tool for document conversion (like `ffmpeg` for documents)

## Positioning

**"The FFmpeg of documents"** — one tool, all formats, blazing fast, open source.

Just as FFmpeg became the universal standard for media processing, office_oxide aims to be the universal standard for document processing. Fast enough for real-time, complete enough for production, simple enough for a `pip install`.
