# Open-Source Readiness Gaps

Tracking list of gaps identified during the pre-launch audit. Everything
that doesn't require a new benchmark run should be closed before we post
to HN / Reddit / r/rust. Items marked **(re-run)** depend on a fresh
benchmark pass and are deferred.

Legend: `[ ]` open, `[x]` closed, **(re-run)** requires fresh benchmarks.

## Benchmark story

- [x] **No Apache POI / Apache Tika comparison.** Closed via (b):
  README tagline reframed to "Fastest *Native* Office Document Library",
  scope block added explaining JVM is out of scope for this round, and
  BENCHMARKS.md "Scope and non-goals" section added. POI/Tika row
  remains future work (deferred re-run).

- [x] **`bench_python.py` is out of sync with BENCHMARKS.md tables.**
  Added `try_python_calamine` + `try_xlrd` wrappers, registered them
  against `.xlsx` / `.xls`, extended `collect_files` to include `.xls`,
  and added `--json OUT` output. Script is now in sync with the tables
  (numbers still need a fresh run).

- [x] **No reproducible bench entry point.** Added
  `scripts/bench.sh` + `scripts/bench-requirements.txt` — machine-spec
  capture, pinned pip versions, native-tool check, build,
  rust + python bench runs, consolidated output dir.

- [x] **No memory / peak-RSS plumbing.** Added RSS capture to
  `bench_python.py` (`resource.getrusage(RUSAGE_SELF).ru_maxrss` at
  start + end, emitted in `--json` output). Actual numbers come from
  the deferred re-run.

- [x] **PPTX has no native competitor.** Documented explicitly in the
  "Scope and non-goals" section of BENCHMARKS.md.

- [x] **Machine spec missing from BENCHMARKS.md.** Added "Reproducing
  these numbers" section with reference environment, pinned competitor
  versions table (Python + apt + Rust crate), and pointer to
  `scripts/bench.sh` which captures ground-truth spec per run.

- [x] **`bench_rust` didn't actually run office_oxide.** Biggest
  single gap in the original harness: `bench_rust` measured calamine,
  docx-rs, and dotext but office_oxide was absent — so there was *no*
  Rust-to-Rust number anywhere in the repo. Fixed by:
  - adding `office_oxide = { path = ".." }` to `bench_rust/Cargo.toml`
    (bench_rust kept out of the main workspace via a nested
    `[workspace]` table so its competitor deps don't leak in);
  - adding `try_office_oxide()` and wiring it into the CLI dispatch
    for all of docx / xlsx / pptx / xls;
  - adding `.xls` to the file walker so calamine and office_oxide can
    be compared on legacy Excel in the same process;
  - adding `--json OUT` output, `libc::getrusage`-based peak-RSS
    capture, and mean-per-file in the text report;
  - retaining `catch_unwind` so competitor panics (e.g. calamine
    0.26.1 on `umya_aaa_large_string.xlsx`) are counted as failures
    instead of crashing the harness;
  - teaching `scripts/bench.sh` to call `bench_rust --lib all --json`
    so the full Rust matrix lands in `rust.json` alongside
    `python.json`.
  Smoke-tested on a 10-file sample; `cargo test --workspace` still
  340/0. Full corpus numbers come with the deferred re-run.

## Claim framing

- [x] **README tagline overreaches.** Tagline now "The Fastest Native
  Office Document Library"; added explicit scope blockquote listing
  every library we compare against and noting POI/Tika as out-of-scope.

- [x] **"Beats calamine on XLSX" claim is end-user-measured only.**
  Subsumed by the bench_rust fix above — office_oxide and calamine
  now run head-to-head in one process under identical rustc / LTO /
  RSS accounting, so the claim is defensible once the full corpus
  re-run lands.

## Repo consistency

- [x] ~~**Stale memory references to missing doc files.**~~ Verified
  false positive from audit. `benchmarks_ooxml.md`,
  `benchmarks_legacy.md`, `benchmarks_calamine.md` exist as
  point-in-time reference snapshots inside the project memory store;
  they are not repo docs and were never claimed to be.

- [ ] **`bench_results.json` at repo root is undocumented.** Either
  describe its schema in BENCHMARKS.md, move it under `scripts/`, or
  delete it. (Cheap cleanup; left for the person driving the re-run
  since they'll want to replace it anyway.)

## Deferred — requires re-run

- [ ] **(re-run)** Re-run full Python bench with `python-calamine`
  included and publish resulting numbers.
- [ ] **(re-run)** Re-run bench with peak-RSS capture; add RSS column
  to tables.
- [ ] **(re-run)** Execute `bench_rust` against the full corpus and
  publish Rust-to-Rust `calamine` comparison numbers.
- [ ] **(re-run)** If we decide to add POI/Tika: stand up a JVM harness
  and run the corpus through it.

## Non-gaps (verified ready)

- Licensing: dual MIT/Apache-2.0, all metadata set, SECURITY / CoC /
  CONTRIBUTING all substantive.
- CI: clippy, fmt, 3-OS × stable/beta/nightly, MSRV 1.85, cargo-deny,
  audit, taplo, semver-checks, 85% coverage gate, multi-language
  binding jobs, dependabot, release workflow.
- Tests: 340 unit + 9 doctest, 0 failures.
- Docs: per-language getting-started (rust/python/go/csharp/js/wasm/c),
  ARCHITECTURE.md, MISSION.md, llms.txt.
- Packaging: v0.1.0 coherent across Cargo/pyproject/package.json, no
  unpublished path-deps, no leaked secrets / personal paths.
