# Office Oxide — Benchmark Results

Benchmarked on **6,062 files** from **11 open-source test suites**.
Single-thread, release build with LTO, 60s timeout per file, warm disk
cache (one warm-up pass discarded). Measurements are the median of
three consecutive runs on an otherwise idle machine.

## Summary

| Metric | Value |
|--------|-------|
| Total files tested | 6,062 |
| Formats supported | .docx, .xlsx, .pptx, .doc, .xls, .ppt |
| Overall pass rate | **98.4%** (5,965 / 6,062) |
| All 97 failures | Invalid ZIP/CFB containers, missing required parts, malformed XML, CVE-exploit fixtures, non-Office files (WordPerfect, Excel 3/4, empty, mislabeled) |
| Real failures on valid documents | **0** |

## OOXML Formats (5,146 files)

Corpus from: LibreOffice Core, Apache POI, Open XML SDK, ClosedXML, Pandoc, python-docx, python-pptx.

### .docx (2,538 files)

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **0.8ms** | **3.9ms** | **98.9%** | **MIT** |
| python-docx | Python | 11.8ms | 98ms | 95.1% | MIT |

### .xlsx (1,802 files)

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **5.0ms** | **40ms** | **97.8%** | **MIT** |
| python-calamine | Rust (Python) | 13.9ms | 183ms | 96.6% | MIT |
| openpyxl | Python | 94.5ms | 698ms | 96.2% | MIT |

The `python-calamine` row includes pyo3 dispatch overhead. A
Rust-to-Rust comparison against the underlying `calamine` crate (0.26)
is available via `bench_rust`:

```bash
cd bench_rust
cargo run --release --bin bench_rust -- --json out.json --lib all CORPUS_DIR
```

This runs office_oxide, calamine, docx-rs, and dotext in a single
process (same rustc, same LTO, same RSS accounting) and emits
per-library mean/total/failure stats plus peak RSS to `out.json`. The
scripts/bench.sh entry point invokes it automatically. Rust-to-Rust
numbers will land in this table with the next full benchmark run.

### .pptx (806 files)

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **0.7ms** | **3.9ms** | **98.4%** | **MIT** |
| python-pptx | Python | 32.5ms | 174ms | 86.7% | MIT |

## Legacy Formats (916 files)

Corpus from: Apache POI, Apache Tika, calamine, openpreserve, oletools, LibreOffice Core.

### .doc (246 files)

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **0.3ms** | **3.4ms** | **94.7%** | **MIT** |
| catdoc | C | 4.3ms | 41ms | 90.2% | GPL-2.0 |
| antiword | C | 4.5ms | 66ms | 76.8% | GPL-2.0 |

### .xls (494 files)

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **2.8ms** | 75ms | **99.2%** | **MIT** |
| xls2csv (catdoc) | C | 6.9ms | **58ms** | 84.0% | GPL-2.0 |
| python-calamine | Rust (Python) | 9.0ms | 96ms | 90.7% | MIT |
| xlrd | Python | 36.6ms | 503ms | 93.1% | BSD-3 |

office_oxide leads on mean (2.4× faster than xls2csv) and pass rate (+15pp), but xls2csv has a tighter p99 — its output is truncated / lossy on complex sheets which keeps its tail short.

### .ppt (176 files)

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **0.7ms** | **6.6ms** | **100%** | **MIT** |
| catppt (catdoc) | C | 2.8ms | 8ms | 77.8% | GPL-2.0 |

## Test Corpus Sources

| Source | Files | License |
|--------|-------|---------|
| LibreOffice Core | 2,185 | MPL-2.0 |
| Apache POI | 1,298 | Apache-2.0 |
| Open XML SDK | 707 | MIT |
| ClosedXML | 371 | MIT |
| Pandoc | 224 | GPL-2.0 |
| python-docx / python-pptx | 111 | MIT |
| Apache Tika | 108 | Apache-2.0 |
| calamine | 28 | MIT |
| openpreserve | 20 | CC0 |
| oletools | 17 | BSD-2 |
| LibreOffice (legacy) | 12 | MPL-2.0 |

## Failure Analysis — 97 / 6,062 files (1.6%)

By category (from `analyze_failures` across the full corpus):

| Category | Count | Notes |
|----------|------:|-------|
| Invalid ZIP / CFB archive | 43 | Truncated, missing EOCD, bad CFB magic signature |
| Missing required part | 21 | Encrypted, password-protected, or stream absent |
| Malformed XML | 18 | XML bombs, ill-formed tags, fuzz-corrupted content |
| Invalid CFB header | 15 | WordPerfect / IBM DisplayWrite / Excel 3/4 misnamed as .doc/.xls, CVE-exploit fixtures |

By format:

| Format | Fail | Pass | % |
|--------|-----:|-----:|--:|
| .docx  |   27 | 2,511 | 98.9% |
| .xlsx  |   40 | 1,762 | 97.8% |
| .pptx  |   13 |   793 | 98.4% |
| .doc   |   13 |   233 | 94.7% |
| .xls   |    4 |   490 | 99.2% |
| .ppt   |    0 |   176 | 100%  |

**Zero failures on legitimate Word 97+ / Excel 97+ / PowerPoint 97+ files.** Every non-passing file is either a non-Office input misnamed with an Office extension, a fuzz/CVE corpus fixture, or a genuinely invalid archive.

## Methodology

- **Runner**: `office_oxide_bench` (recursive walker over the corpus;
  calls `Document::open()` + `plain_text()` per file)
- **Measured**: parse time (open + parse) + full text extraction
- **Environment**: Linux, single-threaded, release build with LTO
- **Cache state**: warm disk cache (one discard run beforehand),
  steady-state parser performance
- **Runs**: three consecutive runs on an idle system; figures are the
  typical / median across runs
- **Competitive libs**: installed via pip/apt, measured with Python
  `time.monotonic()` or subprocess timing
- **Date**: April 2026

## Reproducing these numbers

Run `scripts/bench.sh CORPUS_DIR` — it captures machine spec, installs
pinned competitor versions (`scripts/bench-requirements.txt`), builds
office_oxide with `--release` (LTO on), and produces
`rust.txt` / `python.txt` / `python.json` / `machine.json` under an
output directory. See that script for the exact invocation sequence.

### Environment (target reference)

These tables were produced on the following reference environment. Your
hardware will differ; relative differences between libraries are the
durable signal, not the absolute numbers.

| Field | Value |
|-------|-------|
| OS | Linux x86_64 |
| Kernel | 6.6.x |
| Rust toolchain | stable 1.85 with LTO (`lto = true`, `opt-level = 3`) |
| Python | 3.12 |
| Single-threaded | yes (`RAYON_NUM_THREADS=1` implied) |

`scripts/bench.sh` captures CPU model, RAM, kernel, rustc/cargo
versions, and exact package versions into `machine.json` +
`python-libs.txt` in the output directory so third parties can compare
against their own runs.

### Pinned competitor versions

Python libraries (from `scripts/bench-requirements.txt`):

| Library | Version |
|---------|---------|
| markitdown | 0.0.1a3 |
| python-docx | 1.1.2 |
| openpyxl | 3.1.5 |
| python-pptx | 1.0.2 |
| python-calamine | 0.2.3 |
| xlrd | 2.0.1 |

Native tools (Debian/Ubuntu `apt`):

| Tool | Package | Notes |
|------|---------|-------|
| catdoc | `catdoc` | provides `catdoc`, `xls2csv`, `catppt` |
| antiword | `antiword` | orphaned upstream; distro-packaged |

Rust comparison crates (from `bench_rust/Cargo.toml`):

| Crate | Version |
|-------|---------|
| calamine | 0.26 |
| dotext | 0.1 |
| docx-rs | 0.4 |

## Scope and non-goals

- **Apache POI / Apache Tika are not compared here.** They require a
  JVM and target a different deployment shape (long-running server
  processes, not per-invocation CLI/library calls). A JVM-inclusive
  benchmark may be added in a later release; this table is explicitly
  about native / no-runtime libraries.
- **PPTX has no mature native competitor other than python-pptx.** We
  report what exists; we do not compare against MarkItDown's PPTX
  pipeline separately because it wraps python-pptx internally.
- **Memory / peak-RSS numbers are not yet published.** The harnesses
  (`bench_python.py --json`, `bench_rust`) capture RSS but we have not
  yet executed a full corpus run with those numbers recorded. Expected
  in the next benchmark release.

