# Round 1 (CONTAMINATED) — my own review of PR #116

Written BEFORE the blind round returned, so the later comparison is honest
rather than retrofitted. Contaminated by construction: I read the diff, the PR
description, and the contributor's own claims first.

PR: #116 `feat(doc): extract tables and lists from legacy binary .doc files`
Author: xugangqiang (external fork), 2 commits, +1680/-50, label `ai-assisted`.
Head: 621a8046. CLA signed. mergeStateStatus BLOCKED (awaiting review).

## Mechanical state (verified locally, not taken on trust)

- `cargo build --workspace` — clean.
- `cargo clippy --workspace --all-targets` — clean, zero warnings.
- `cargo test --workspace` — 464 lib tests + all integration tests pass, 0 failures.
- `cargo fmt --all -- --check` — **FAILS, 30 diffs** across 7 files
  (convert_doc.rs 3, document.rs 1, papx.rs 5, sprm.rs 9, and 4 in each of the
  3 new integration tests). Contributor formatted with default rustfmt, not the
  repo's rustfmt.toml (max_width=100, fn_call_width=80, fn_params_layout=Tall,
  match_block_trailing_comma=true).
- CI on the PR ran only 5 checks (CLA, AI-assisted label, binary-artifact
  guard, PR title, `check`). The fmt/clippy/test matrix did NOT run — fork PR.
  Green checks on the PR page therefore do NOT mean the quality gate passed.

## Findings

### F1 (HIGH, empirically confirmed) — field codes / control chars leak into IR

`doc_to_ir` now builds the IR from `papx::build_paragraphs`, which slices text
out of **`raw_text`** (src/doc/document.rs), the pre-sanitize string. The old
path built it from `sanitize_text(&raw_text)`. `sanitize_text`
(src/doc/piece_table.rs:205) strips 0x01, 0x08, 0x13, 0x14, 0x15 and maps
CR to LF, 0x07 to TAB, 0x0B/0x0C to LF.

Confirmed by probe: a prose paragraph whose raw text carries a field-code run
plus a picture char now serialises into the IR verbatim, field instruction text
and all — the emitted TextSpan contained the literal string `HYPERLINK
"http://x"` together with the raw 0x13/0x14/0x15/0x01 control characters.

Affects every `.doc` containing a hyperlink, TOC, page number, cross-reference,
date field, or inline image — i.e. most real Word documents. `plain_text()` is
unaffected (still sanitized); only `to_ir()` regresses, so every downstream
consumer of the IR (markdown, JSON, MCP, bindings, pdf_oxide) is affected.
Not caught by any test, because the three new fixtures contain no fields.

### F2 (OPEN QUESTION) — sprmTDefTable variable-length prefix

`sprm.rs` treats every `spra == 6` SPRM as `[1-byte length][operand]`, and the
module doc says this "was verified empirically against the table.doc fixture".

Probe on table-merges.doc: the walk does consume exactly every byte
(157/157, 245/245, 245/245, 113/113), and 0xD608 is NOT the last SPRM — nine
SPRMs follow it (d609 d612 d670 d634 x4 3403 3466) and still decode cleanly.
So the assumption holds for these fixtures and my off-by-one worry is disproved
at this size.

BUT all three fixtures have small tables. A TDefTable operand is roughly
24*ncols + 4 bytes, so it exceeds 254 at about 11 columns. Suspicion: there is
a length-escape mechanism for large variable operands that this walker does not
implement, which would desync the entire grpprl walk on wide tables.
UNRESOLVED — this is exactly what blind agent B is being asked, without being
told a PR exists.

### F3 (MEDIUM) — dead code in Fib

`fc_plcf_lst` / `lcb_plcf_lst` added to `Fib` and parsed, never read anywhere.

### F4 (MEDIUM) — memory

`build_paragraphs` materialises a `Vec<char>` of the entire document text
(4 bytes per char) purely to slice by character position. On large `.doc` files
this is a new allocation several times the size of the text.

### F5 (LOW) — acknowledged limitations

Contributor is upfront: lists always emit as unordered (bullet) regardless of
whether Word had a numbered list; adjacent distinct lists merge because `ilfo`
grouping is unimplemented; style-inherited list membership undetected.
Ordered-vs-bullet silently wrong (not absent) is arguably worse than not
extracting lists at all — a numbered procedure renders as bullets.

### F6 (LOW) — scope creep

Adds `examples/rust/inspect_doc.rs` plus a Cargo.toml `[[example]]` entry; a
debugging aid, unrelated to the feature.

## My round-1 leaning

The approach is right and matches how Apache POI models this. Blockers: F1,
plus F2 if it turns out to be real. Everything else is follow-up material.
