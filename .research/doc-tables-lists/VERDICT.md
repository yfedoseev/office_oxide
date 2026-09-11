# PR #116 — verdict after contaminated + PR-blind review

> **SUPERSEDED IN PART — see `round3-blind-verdict.md` (round 3, blind, run
> after the contributor's fixes landed).** B1/B2/B3 are fixed. Two claims in
> this document are corrected there: (a) the section "One thing you got right
> that I initially thought was wrong" about `vertMerge` is **wrong** —
> `vertMerge` is a 2-bit `VerticalMergeFlag` field (`fvmRestart = 0x03` →
> `0x0060`), not two independent flags; (b) B2's framing that field-instruction
> text is a `to_ir()`-only regression is **wrong** — `sanitize_text` strips only
> the delimiters, so the instruction text leaks on the flat path too. Round 3
> also found the SPRM opcode table is wrong for `sprmPIlvl`/`sprmPIlfo`/
> `sprmPChgTabs`, which this round read and passed.

Method: `claude_code_tricks/blind-agent-pr-review.md`, run properly.
Round 1 = my own review with the diff in hand (`round1-contaminated.md`).
Round 2 = three fresh non-fork agents given only the neutral problem, never
told a PR existed, then every claim diffed against the real code **by me**
(`round2-verified.md`). Blind agent reports in `blind/`.

## Does the research reject the approach, or only find defects?

**It confirms the approach.** Both reference implementations model this the way
the PR does: merges derived exclusively from `TC80` bits (LibreOffice ignores
`sprmTVertMerge`/`sprmTMerge` entirely; POI does too), and `col_span` from a
unified column-edge grid (POI's `buildTableCellEdgesArray`). The PR is not
doing something idiosyncratic. Every finding below is a fixable defect inside a
sound design.

Verdict: **request changes** — 3 blockers, all narrow.

## Blockers

### B1 — `sprmTDefTable` length prefix is 2 bytes, not 1
Silent structural corruption on any table with **12 or more columns**.
[MS-DOC] 2.2.5.1 names exactly two exceptions to the 1-byte rule, and
sprmTDefTable is one; `cb` is `u16` = remainder + 1.
Independently reached by **two of the three blind agents**, from separate
sources — POI `SprmOperation.initSize`/`SPRM_LONG_TABLE`, and LibreOffice's
`enum SprmType { L_FIX, L_VAR, L_VAR2 }` where `L_VAR2` is the 2-byte class.
Measured threshold and failure mode in `round2-verified.md` V7.
Fix: special-case 0xD608 and 0xD606, consume `cb + 1`, and shift
`parse_tdef_table`'s offsets by one (its "reserved byte" is `cb`'s high byte).

### B2 — IR text is sliced from a CP-misaligned, unsanitized string
Two symptoms, one root cause, in `papx::build_paragraphs`.
1. It slices `raw_text` (pre-`sanitize_text`), so field instruction text and raw
   0x01/0x13/0x14/0x15 control chars land in the IR. Confirmed by probe.
   Affects any `.doc` with a hyperlink, TOC, page number, cross-ref or image.
2. It indexes a `Vec<char>` by **character position**. One CP is one UTF-16
   code unit, but `from_utf16_lossy` collapses a surrogate pair to one `char`,
   so a single astral char desyncs the rest of the document. Because
   `terminator = chars[end - 1]` drives cell/row grouping, this corrupts table
   structure, not just text. `.min(chars.len())` clamps, so it fails silently.
   `extract_text` silently skipping unfittable pieces is a second drift source.
Fix (one change, both symptoms): stop slicing a flat string. Resolve each
paragraph's CP range against `&[Piece]` + `word_doc`, decode that range, and run
`sanitize_text` on the per-paragraph slice.

### B3 — `cargo fmt --all -- --check` fails, 30 diffs across 7 files
Contributor used stock rustfmt, not the repo's `rustfmt.toml`. Note the PR's
green checks are misleading: only 5 checks ran (CLA, label, artifact guard, PR
title, `check`) — the fmt/clippy/test matrix never ran because it is a fork PR.
Locally: build clean, clippy clean, 464 lib + all integration tests pass.

## Medium — should fix, need not block

- **M1** No tolerance when unifying column edges. LibreOffice snaps within
  `nTolerance = 4` twips because Word rounds boundaries per row; the PR uses
  exact `sort`/`dedup`. Every 1-3 twip difference injects a spurious grid edge
  and inflates `col_span`. The fixture cannot catch this — its rows align
  exactly.
- **M2** `build_paragraphs` materialises a `Vec<char>` of the whole document
  (4 bytes/char) purely to slice by CP. Goes away with the B2 fix.
- **M3** Dead parsed state: `PapProps::itap`, `PapProps::ilfo`, and
  `Fib::fc_plcf_lst`/`lcb_plcf_lst` are all parsed and never read. Consequence
  of the `itap` one: **nested tables are silently flattened** into the outer
  table's cells. Not claimed in the PR description; should be stated.
- **M4** Soft line breaks (0x0B) previously split a paragraph via
  `sanitize_text`; the structured path keeps them literal. Segmentation change.

## Low / follow-up

- **L1** `sprmPChgTabs` (0xC615) is the *other* documented exception — `cb == 255`
  is an escape and the length must be computed from the payload. Same bug class
  as B1, far rarer trigger. (POI gets this one wrong too, so don't copy POI.)
- **L2** Lists always emit as unordered. A numbered procedure silently renders
  as bullets — arguably worse than not extracting lists. Contributor is upfront
  about this; worth gating behind correct `ilfo` resolution.
  Landmine for that follow-up: `ilfo == 2047` is a trapdoor into the Word 6/95
  ANLD system outside `PlfLst`, with no bullet code at all; resolving
  `rgLfo[2046]` yields garbage.
- **L3** Zero-width cells (`centers[i] == centers[i+1]`) mean "cell does not
  exist"; `count_grid_edges` gives them `col_span` 1 via `.max(1)`.
- **L4** `examples/rust/inspect_doc.rs` + Cargo.toml entry is unrelated scope.
- **L5** The fuzz target never calls `to_ir()`, so the new IR assembly is
  unfuzzed. The new *binary* parsers (papx.rs, sprm.rs) **are** covered, since
  `DocDocument::parse` builds paragraphs eagerly.

## Credit where due — claims that did NOT survive verification

- Blind agent said headers/footers/footnotes would leak into the IR. **Wrong** —
  the PR filters `cp < text_len` against `ccpText`.
- Blind agent's clean spec reading of TCGRF `vertMerge` (2-bit field, 3=restart)
  implies the PR's "Word keeps the restart bit on continuation rows" comment is
  a rationalisation. **Wrong** — I dumped the fixture: rows 1 and 2 both carry
  `rgf=0x0060` in column 0 and POI treats them as one 2-row merge. The
  contributor's observation is correct and the tidy spec reading is not what
  Word writes. Known limitation: two *distinct* stacked merges in one column
  are indistinguishable and will collapse.
- Blind agent said the change has "zero fuzz coverage". **Overstated** — see L5.

## Separate issue, owned by no PR

`parse_plc_pcd` (piece_table.rs) reads `cp_start`/`cp_end` with no monotonicity
validation; `extract_text` then does `cp_end.min(max_chars) - cp_start`, an
unchecked `u32` subtraction. A crafted `.doc` panics in debug/fuzz and wraps
silently in release (`[profile.release]` sets no `overflow-checks`). Reachable
from `Document::from_reader`, which the fuzz target does call. Present on `main`
today. File independently of #116.
