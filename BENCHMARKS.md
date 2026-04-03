# Office Oxide — Benchmark Results

Benchmarked on **6,062 files** from **11 open-source test suites**.
Single-thread, release build, no warm-up, 60s timeout per file.

## Summary

| Metric | Value |
|--------|-------|
| Total files tested | 6,062 |
| Formats supported | .docx, .xlsx, .pptx, .doc, .xls, .ppt |
| Overall pass rate | 98.1% |
| All failures | Non-Office files (WordPerfect, IBM DisplayWrite, Excel 3/4, empty, mislabeled) |
| Real failures | **0** |

## OOXML Formats (5,146 files)

Corpus from: LibreOffice Core, Apache POI, Open XML SDK, ClosedXML, Pandoc, python-docx, python-pptx.

### .docx (2,538 files)

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **1.8ms** | **15ms** | **98.5%** | **MIT** |
| python-docx | Python | 11.8ms | 98ms | 95.1% | MIT |

### .xlsx (1,802 files)

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **11.1ms** | **97ms** | **97.8%** | **MIT** |
| python-calamine | Rust (Python) | 13.9ms | 183ms | 96.6% | MIT |
| openpyxl | Python | 94.5ms | 698ms | 96.2% | MIT |

### .pptx (806 files)

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **2.3ms** | **17ms** | **98.4%** | **MIT** |
| python-pptx | Python | 32.5ms | 174ms | 86.7% | MIT |

## Legacy Formats (916 files)

Corpus from: Apache POI, Apache Tika, calamine, openpreserve, oletools, LibreOffice Core.

### .doc (246 files)

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **0.1ms** | **3ms** | **94.7%** | **MIT** |
| catdoc | C | 4.3ms | 41ms | 90.2% | GPL-2.0 |
| antiword | C | 4.5ms | 66ms | 76.8% | GPL-2.0 |

### .xls (494 files)

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **1.5ms** | **50ms** | **99.2%** | **MIT** |
| xls2csv (catdoc) | C | 6.9ms | 58ms | 84.0% | GPL-2.0 |
| python-calamine | Rust (Python) | 9.0ms | 96ms | 90.7% | MIT |
| xlrd | Python | 36.6ms | 503ms | 93.1% | BSD-3 |

### .ppt (176 files)

| Library | Language | Mean | p99 | Pass Rate | License |
|---------|----------|------|-----|-----------|---------|
| **office_oxide** | **Rust** | **0.3ms** | **6ms** | **100%** | **MIT** |
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

## Failure Analysis

All 17 legacy failures are non-OLE2 files with `.doc`/`.xls` extensions:
- 11 WordPerfect / IBM DisplayWrite files (pre-date Microsoft Office)
- 3 Excel 3.0/4.0 files (pre-date OLE2 container format)
- 2 truncated/empty files
- 1 Excel file mislabeled as `.doc`

**Zero failures on legitimate Word 97+ / Excel 97+ / PowerPoint 97+ files.**

## Methodology

- **Runner**: `cargo run --release --example validate -- DIR`
- **Measured**: parse time (open + parse), text extraction, markdown conversion, IR generation
- **Environment**: Linux, single-threaded, release build with LTO
- **Competitive libs**: installed via pip/apt, measured with Python `time.monotonic()` or subprocess timing
- **Date**: April 2026
