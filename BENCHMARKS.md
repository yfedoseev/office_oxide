# Office Oxide — Benchmark Results

Benchmarked on **10,386 files** (the project's local test collection:
3,519 .docx, 3,367 .xlsx, 1,123 .pptx, 780 .doc,
1,180 .xls, 417 .ppt), run on 2026-10-29. Single-thread, release
build with LTO, 30 s timeout per file for the Python and command-line
tools, warm disk cache.

The figures for the Rust libraries (office_oxide, calamine, docx-rs) are the
**median of three consecutive runs**; those of the Python libraries and the
command-line tools are **one run** each (see Methodology).

## Summary

| Metric | Value |
|--------|-------|
| Total files tested | 10,386 |
| Formats supported | .docx, .xlsx, .pptx, .doc, .xls, .ppt |
| office_oxide opened without error | **97.3%** (10,110 / 10,386) |
| Files office_oxide did not open | 276 (see Failure Analysis) |

## OOXML Formats

### .docx (3,519 files)

| Library | Language | Mean | Pass Rate | License |
|---------|----------|------|-----------|---------|
| **office_oxide** | **Rust** | **1.7ms** | **98.7%** | **MIT** |
| docx-rs (crate) | Rust | 5.0ms | 98.0% | MIT |
| python-docx | Python | 9.5ms | 95.7% | MIT |
| markitdown | Python | 147.5ms | 98.4% | MIT |

### .xlsx (3,367 files)

| Library | Language | Mean | Pass Rate | License |
|---------|----------|------|-----------|---------|
| calamine (crate) | Rust | 4.2ms | 98.2% | MIT |
| **office_oxide** | **Rust** | **6.2ms** | **98.6%** | **MIT** |
| python-calamine | Rust (Python) | 8.7ms | 97.8% | MIT |
| openpyxl | Python | 60.7ms | 96.0% | MIT |
| markitdown | Python | 307.5ms | 96.2% | MIT |

`calamine (crate)` is the Rust crate read directly; `python-calamine` is the
same library behind a Python binding, so it includes the binding's dispatch
overhead. Both are run to read the cell values of every sheet, while
office_oxide also extracts the text of comments, charts and the like, so the
means are not like-for-like.

### .pptx (1,123 files)

| Library | Language | Mean | Pass Rate | License |
|---------|----------|------|-----------|---------|
| **office_oxide** | **Rust** | **2.6ms** | **98.1%** | **MIT** |
| python-pptx | Python | 13.4ms | 90.0% | MIT |
| markitdown | Python | 74.1ms | 90.2% | MIT |

`markitdown` is measured with its optional converters installed
(`markitdown[docx,pptx,xlsx,xls]`); it wraps other libraries, so it is
slower than they are.

## Legacy Formats

### .doc (780 files)

| Library | Language | Mean | Pass Rate | License |
|---------|----------|------|-----------|---------|
| **office_oxide** | **Rust** | **4.1ms** | **90.0%** | **MIT** |
| catdoc | C | 5.6ms | 95.5% | GPL-2.0 |
| antiword | C | 5.6ms | 79.5% | GPL-2.0 |

### .xls (1,180 files)

| Library | Language | Mean | Pass Rate | License |
|---------|----------|------|-----------|---------|
| **office_oxide** | **Rust** | **5.1ms** | **95.3%** | **MIT** |
| calamine (crate) | Rust | 9.0ms | 93.0% | MIT |
| xlrd | Python | 26.9ms | 92.5% | BSD-3 |
| xls2csv (catdoc) | C | 36.5ms | 97.7% | GPL-2.0 |
| python-calamine | Rust (Python) | 44.5ms | 92.6% | MIT |

### .ppt (417 files)

| Library | Language | Mean | Pass Rate | License |
|---------|----------|------|-----------|---------|
| **office_oxide** | **Rust** | **6.7ms** | **92.3%** | **MIT** |
| catppt (catdoc) | C | 80.2ms | 99.8% | GPL-2.0 |

**Reading the pass rates.** "Pass" means the library opened the file and
returned without an error. The command-line tools (catdoc, antiword, xls2csv,
catppt) exit 0 on many files they cannot read and print little or nothing, so
their rate overstates what they extracted; office_oxide reports a file it
cannot read as an error, and so does most of the other libraries. The
collection also holds deliberately damaged and password-protected files,
which is where most of office_oxide's failures are.

## Failure Analysis — 276 / 10,386 files (2.7%)

What office_oxide reported for the files it did not open:

| Category | Count |
|----------|------:|
| Invalid ZIP / CFB container | 143 |
| Encrypted / password-protected | 87 |
| Missing or duplicate package part | 32 |
| Other | 13 |
| Malformed XML | 1 |

By format:

| Format | Fail | Pass | % |
|--------|-----:|-----:|--:|
| .docx |   44 | 3,475 | 98.7% |
| .xlsx |   46 | 3,321 | 98.6% |
| .pptx |   21 | 1,102 | 98.1% |
| .doc  |   78 |   702 | 90.0% |
| .xls  |   55 | 1,125 | 95.3% |
| .ppt  |   32 |   385 | 92.3% |

Password-protected files are refused with a clear error (the opt-in
`default-password` feature opens the ones Office protects with its fixed
default password). The remaining failures are invalid or truncated ZIP / CFB
containers, packages with a missing required part, and fuzz or exploit
fixtures kept in the collection on purpose.

## Methodology

- **Runner**: `bench_rust` for the Rust libraries (calls `Document::open()` +
  `plain_text()` per file for office_oxide), `bench_python.py` for the Python
  libraries, `bench_native.py` for the command-line tools. One file is read
  at a time.
- **Measured**: open + parse + full text extraction, per file. Means are
  total time over all files of the format (failures included). Per-file
  times are not kept, so there is no p99 column.
- **Environment**: Linux, single-threaded, release build with LTO, nothing
  else running during a measurement.
- **Runs**: three consecutive runs of the Rust harness, median reported; one
  run of the Python harness and of the command-line tools.
- **Crashes**: a file that ends the whole process of a competing library
  (an allocation it cannot satisfy aborts the process) is counted as that
  library's failure without being read, so one hostile file does not end the
  run for every other file. Two files do this to calamine (the crate and
  python-calamine): `xls/edge_cases/phpspreadsheet-xls_oleloop.xls` and
  `xlsx/basic/lo_too-many-cols-rows.xlsx`. The Python harness also stops a
  library call that passes 10 GiB of resident memory.
- **Command-line tools** are one subprocess per file, so their times include
  process start-up.

## Reproducing these numbers

Run `scripts/bench.sh CORPUS_DIR`; it captures the machine spec, installs the
pinned Python libraries (`scripts/bench-requirements.txt`), builds
office_oxide with `--release` (LTO on) and runs the Rust and Python
harnesses. The command-line tools run with
`python3 bench_native.py CORPUS_DIR --json OUT`. A file that kills a
competing library's process can be listed in `BENCH_CRASH_LIST` (one path per
line for `bench_rust`, `library<TAB>path` for `bench_python.py`).

### Environment

| Field | Value |
|-------|-------|
| CPU | Intel Core i5-10210U @ 1.60 GHz (8 threads) |
| Memory | 14 GiB |
| Kernel | Linux 6.6 x86_64 |
| Rust toolchain | stable 1.98 with LTO (`lto = true`, `opt-level = 3`) |
| Python | 3.12 |

Hardware differs between machines; the relative differences between
libraries are the durable signal, not the absolute numbers.

### Pinned competitor versions

Python libraries (from `scripts/bench-requirements.txt`):

| Library | Version |
|---------|---------|
| markitdown | 0.1.9 (`[docx,pptx,xlsx,xls]`) |
| python-docx | 1.2.0 |
| openpyxl | 3.1.5 |
| python-pptx | 1.0.2 |
| python-calamine | 0.9.0 |
| xlrd | 2.0.2 |

Native tools (Debian/Ubuntu `apt`):

| Tool | Package | Notes |
|------|---------|-------|
| catdoc | `catdoc` 0.95 | provides `catdoc`, `xls2csv`, `catppt` |
| antiword | `antiword` | orphaned upstream; distro-packaged |

Rust comparison crates (from `bench_rust/Cargo.lock`):

| Crate | Version |
|-------|---------|
| calamine | 0.37.0 |
| dotext | 0.1.1 |
| docx-rs | 0.4.22 |

`dotext` is run by the harness too, but it reads only a few of the files
(it opened 2 of 3,519 .docx and 1 of 3,367 .xlsx), so it is not
tabulated.

## Scope and non-goals

- **Apache POI / Apache Tika are not compared here.** They require a
  JVM and target a different deployment shape (long-running server
  processes, not per-invocation CLI/library calls). A JVM-inclusive
  benchmark may be added in a later release; this table is explicitly
  about native / no-runtime libraries.
- **Memory / peak-RSS numbers** are captured by the harnesses
  (`rust.json`, `python.json`) but not tabulated here.

---

# Earlier run — April 2026 (6,062-file corpus)

The figures below are the earlier run that the README quotes: a different
corpus (6,062 files from 11 public test suites), older pinned versions, and
a harness that also recorded p99. They are kept as they were published and
are **not** comparable with the run above (different files, different
library versions, different machine state).

Benchmarked on **6,062 files** from **11 open-source test suites**.
Single-thread, release build with LTO, 60s timeout per file, warm disk
cache (one warm-up pass discarded). Measurements are the median of
three consecutive runs on an otherwise idle machine.

### Summary

| Metric | Value |
|--------|-------|
| Total files tested | 6,062 |
| Formats supported | .docx, .xlsx, .pptx, .doc, .xls, .ppt |
| Overall pass rate | **98.4%** (5,965 / 6,062) |
| All 97 failures | Invalid ZIP/CFB containers, missing required parts, malformed XML, CVE-exploit fixtures, non-Office files (WordPerfect, Excel 3/4, empty, mislabeled) |
| Real failures on valid documents | **0** |

### OOXML Formats (5,146 files)

Corpus from: LibreOffice Core, Apache POI, Open XML SDK, ClosedXML, Pandoc, python-docx, python-pptx.

#### .docx (2,538 files)

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **0.8ms** | **3.9ms** | **98.9%** | **MIT** |
| python-docx | Python | 11.8ms | 98ms | 95.1% | MIT |

#### .xlsx (1,802 files)

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

#### .pptx (806 files)

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **0.7ms** | **3.9ms** | **98.4%** | **MIT** |
| python-pptx | Python | 32.5ms | 174ms | 86.7% | MIT |

### Legacy Formats (916 files)

Corpus from: Apache POI, Apache Tika, calamine, openpreserve, oletools, LibreOffice Core.

#### .doc (246 files)

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **0.3ms** | **3.4ms** | **94.7%** | **MIT** |
| catdoc | C | 4.3ms | 41ms | 90.2% | GPL-2.0 |
| antiword | C | 4.5ms | 66ms | 76.8% | GPL-2.0 |

#### .xls (494 files)

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **2.8ms** | 75ms | **99.2%** | **MIT** |
| xls2csv (catdoc) | C | 6.9ms | **58ms** | 84.0% | GPL-2.0 |
| python-calamine | Rust (Python) | 9.0ms | 96ms | 90.7% | MIT |
| xlrd | Python | 36.6ms | 503ms | 93.1% | BSD-3 |

office_oxide leads on mean (2.4× faster than xls2csv) and pass rate (+15pp), but xls2csv has a tighter p99 — its output is truncated / lossy on complex sheets which keeps its tail short.

#### .ppt (176 files)

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **0.7ms** | **6.6ms** | **100%** | **MIT** |
| catppt (catdoc) | C | 2.8ms | 8ms | 77.8% | GPL-2.0 |

### Test Corpus Sources

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

### Failure Analysis — 97 / 6,062 files (1.6%)

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

### Methodology

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

### Reproducing these numbers

Run `scripts/bench.sh CORPUS_DIR` — it captures machine spec, installs
pinned competitor versions (`scripts/bench-requirements.txt`), builds
office_oxide with `--release` (LTO on), and produces
`rust.txt` / `python.txt` / `python.json` / `machine.json` under an
output directory. See that script for the exact invocation sequence.

#### Environment (target reference)

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

#### Pinned competitor versions

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

### Scope and non-goals

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

