# Release regression sweep

Compares the last release tag against a branch over a corpus of **real**
Office files. The unit suite is built from synthetic in-code fixtures, so
it asserts what the author believed the format means; this asserts what
actually happens to documents other people made.

Run on the v0.1.9 → 0.1.10 release it found seven defects the 792-test
suite could not see, including a process-killing infinite loop and a
panic on 32 files.

## Running it

```sh
# 1. a corpus (public test data: Apache POI, pandoc)
GH_TOKEN=$(gh auth token) python3 scripts/regression-sweep/fetch_corpus.py

# 2. build both arms — the TAG, not main
git worktree add --detach /tmp/arm-prev v0.1.9
CARGO_TARGET_DIR=/tmp/target-prev cargo build --release -p office_oxide_cli \
  --manifest-path /tmp/arm-prev/Cargo.toml
cargo build --release -p office_oxide_cli

# 3. sweep each arm, then compare
python3 scripts/regression-sweep/sweep.py /tmp/target-prev/release/office-oxide corpus prev.jsonl
python3 scripts/regression-sweep/sweep.py target/release/office-oxide          corpus next.jsonl
python3 scripts/regression-sweep/compare.py   # status transitions + content changes
python3 scripts/regression-sweep/triage.py    # did an error replace real content?
REPO=$PWD python3 scripts/regression-sweep/content_diff.py text
```

**Verify the arm is at the tag** before trusting a number:
`[ "$(git rev-parse HEAD)" = "$(git rev-parse v0.1.9^{commit})" ]`.

## The two axes

**Axis 1 — status transitions.** Office parsers can return `Err` where
they used to return `Ok`, and that is the axis that decides shippability.
The question is never "how many files now error" but **"did the error
replace real extracted content, or silence?"** — `triage.py` answers it.
An error replacing a zero-byte success is an improvement; an error
replacing text is a regression. On the 0.1.10 sweep all 37 newly-erroring
files produced nothing under v0.1.9.

**Axis 2 — content change on files that succeed in both arms.** Compare
word **multisets**, not sets: a set cannot see a doubled table row, and
deduplication is often exactly what a release changes.

## Traps this sweep actually hit

- **A "timeout" that was load.** A fuzzer file sat either side of the 30 s
  cap depending on machine load. Timed quiescently, the new arm was 37%
  *faster* (20.7 s vs 33.1 s). The cap is now 120 s. Never file a timing
  regression measured under concurrent load.
- **Entity resolution reads as token loss.** `AT&T` extracted as `ATT`
  before the fix; afterwards a `\w+` tokeniser sees `AT` + `T` and reports
  `ATT` as lost. The content improved.
- **A stack cliff moves with struct size.** Nested-table documents
  overflowed at 5,000–10,000 levels before this release's fields were
  added and at 3,000–4,000 after. Worse, the cliff depends on the *caller's*
  stack: 512–1,024 in a debug build on a default 2 MiB thread. Calibrate a
  depth cap against the smallest stack a caller has, not the largest.
- **More text is not automatically better.** Broadening changes (text
  boxes, altChunk, AlternateContent) are flattering by construction — every
  extra thing collected reads as recovery. One of them was gluing a text
  box to the next run (`LinzANTRAG`).

## What it cannot see

A defect present in **both** arms. The sweep only shows what changed, so a
document that has always extracted wrong is invisible here by
construction. Adding an external panel — LibreOffice, Tika/POI, pandoc,
python-docx/openpyxl/python-pptx — is the next step; note that Tika wraps
POI, so those two are one opinion, and pandoc deliberately drops
headers/footers, which this release deliberately adds.
