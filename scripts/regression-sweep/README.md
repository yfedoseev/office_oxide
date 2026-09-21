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

## `vanished.py` — the check that actually decides content loss

Byte counts and hash equality answer the wrong question. A release that
restores a dropped `&` makes output *longer*; one that stops emitting
65,536 rows of grid padding makes it dramatically shorter. Neither is
content loss, and both dominate any byte total.

`vanished.py` compares word multisets and reports only words that are
absent from the new arm *and* cannot be explained as re-tokenisation —
the word is not contained in a new token (a join), and is not built from
new tokens (a split). Both directions are things a fix legitimately does:
restoring `&` turns `Ramp` back into `R` and `amp`; restoring an
apostrophe turns `its` into `it` and `s`.

    PREV_BIN=/path/to/old REPO=$PWD python3 vanished.py > vanished.txt

Every surviving entry still needs a human read. In the 0.1.10 sweep all
183 resolved to five deliberate classes (PowerPoint master placeholder
prompts, entity resolution, `General`/number-format rendering, the nesting
depth cap) or to the new arm recovering *more* text than the old one.

**Both scripts take their inputs on the command line.** An earlier version
of `compare.py` hardcoded `next.jsonl`, so re-running it after a fix
silently re-read the pre-fix sweep and reported the old numbers as though
nothing had changed.

## Whole-corpus sweep with an independent reference panel (0.1.12)

The scripts above take a flat directory of files and compare only the
two arms. The 0.1.12 release used a second generation that walks the
nested `~/projects/office_oxide_tests` tree, keeps every surface's
output on disk, and adds the external panel the section above says is
"the next step". It found the largest content-loss defect the crate has
had (`.xls` shared strings cut by a `CONTINUE` record: 80 files at
2–58 % of every reference) — a defect present in *both* arms, which the
two-arm diff is blind to by construction.

```sh
# both arms; outputs land in OUT/out/<relpath>.<surface>.gz + OUT/sweep.jsonl
python3 scripts/regression-sweep/sweep_tree.py /tmp/target-prev/release/office-oxide ~/projects/office_oxide_tests /tmp/sweep/prev --jobs 6
python3 scripts/regression-sweep/sweep_tree.py target/release/office-oxide          ~/projects/office_oxide_tests /tmp/sweep/next --jobs 6
python3 scripts/regression-sweep/compare_tree.py /tmp/sweep/prev /tmp/sweep/next /tmp/sweep/report

# the reference panel (once per corpus; results are reusable across releases)
uv venv /tmp/refvenv && uv pip install --python /tmp/refvenv/bin/python python-docx openpyxl python-pptx xlrd python-calamine
/tmp/refvenv/bin/python scripts/regression-sweep/refpanel.py ~/projects/office_oxide_tests /tmp/panel --jobs 5
java -Xmx2g -jar tika-app.jar -t -i <dir of symlinks to the six format dirs> -o /tmp/tika_out -numConsumers 2 -timeoutThresholdMillis 120000

# who is the odd one out?
python3 scripts/regression-sweep/quorum.py /tmp/sweep/prev /tmp/sweep/next /tmp/tika_out /tmp/panel ~/projects/office_oxide_tests/metadata/manifest.json /tmp/sweep/quorum.json
```

`quorum.py` flags a file when at least two references agree with each
other (≥ 0.85 recall both ways) while office_oxide recovers < 70 % of
their words, when office_oxide is empty or errors against real
references, or when it emits 3× more than every reference.

Traps this generation hit, so they don't get re-discovered:

- **OOM kills and timeouts under a 6-way sweep are not verdicts.**
  Re-measure the file serially with `measure.py` before calling it a
  regression; the kernel killed workers at 3–10 GB that finish alone.
- **`<span>`-per-run HTML tokenises as loss** if inline tags are
  replaced with a space: `xxxx` becomes `x x x x`. Strip inline tags with
  nothing. The `html` axis in `compare_tree.py` does this.
- **Revision-aware vs not.** catdoc, antiword, pandoc and Tika show
  `sprmCFRMarkDel` deletions and `w:vanish` hidden text; the crate hides
  both by policy. A "loss" made of whole sentences and `1.\t` fragments is
  a tracked-changes document — dump the deleted CP ranges before filing.
- **The `ir` surface is not comparable across a projection change**, and
  sheets of one-letter cells score badly under a ≥ 3-letter tokenizer.

## `harness.py` — one run, one verdict per file (0.1.12, pass 4)

The scripts above are the pieces; `harness.py` is the release check
built from them, with the decision model written down in its module
doc: which reference is "expected" (the consensus of references that
agree with each other), when a two-arm difference is a regression (the
consensus has the words and the new arm does not, and the new arm does
not simply side with a different reference than the old one), and what
the harness can explain by itself (`w:vanish` runs, field vocabulary).

```sh
# everything: both arms, the round trip, the pandoc structure comparison, then the report
python3 scripts/regression-sweep/harness.py run \
  --prev-bin /tmp/target-prev/release/office-oxide --next-bin target/release/office-oxide \
  --corpus ~/projects/office_oxide_tests --work /tmp/harness \
  --panel /tmp/panel --tika /tmp/tika_out \
  --validator target/release/examples/corpus_validate --jobs 6

# re-report over existing inputs (a re-sweep of one arm, a changed rule)
python3 scripts/regression-sweep/harness.py report --work /tmp/harness --corpus ~/projects/office_oxide_tests
```

Two axes here exist nowhere else. **Cross-surface** (`SURFACE_SPLIT`):
our own `text`, `markdown` and `html` of one file compared with each
other after removing what each surface legitimately adds — a word on one
surface and not another is a renderer that dropped or invented content,
the *dual-renderer* defect class this release kept finding (every
renderer bug from #381 on). **Round trip** (`ROUNDTRIP_LOSS`): the words
of the original IR against those of the IR read back from what we wrote
(`examples/corpus_validate` `rt_words`) — the only check that sees a
writer silently dropping content while reporting success (#383–#385).

Traps this pass hit:

- **A markdown tag filter must match real tags only.** `<[^>]+>` ate half
  a pharmaceutical document between `< 0.5 mg` and the next `>`; the
  filter now requires a tag name, and an escaped `\<` is text.
- **A timing flag on a file whose output grew is not a regression.**
  v0.1.11 returned nothing from four 7–29 MB `.xls` files (the old record
  cap); this branch returns two million words, matching Tika exactly, in
  twice the time. The perf rule skips files whose word count grew 1.5×.
- **The references have the bugs we fixed.** Tika/POI merge an embedded
  chart's cached series into the sheet (`gnumeric_chart-tests-excel.xls`
  gains `NaN` cells that the sheet does not have) — our #352 class. When
  the consensus disagrees with us on a sheet with an embedded chart, read
  the records before believing it.
- **Load.** Other builds on the same machine turned 7 ms into 600 ms in
  the sweep; every `PERF_REGRESSION` row says to re-measure serially.
