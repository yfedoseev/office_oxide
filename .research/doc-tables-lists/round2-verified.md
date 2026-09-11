# Round 2 — blind findings, verified by me against the real diff

Method: `claude_code_tricks/blind-agent-pr-review.md`. Three fresh non-fork
agents given only the neutral problem, never told PR #116 exists. Checkout held
on `main`; local PR branch renamed; `gh`/`git log`/`git branch`/`.research`
barred. Every claim below was checked by me against the actual code before
being elevated — the agents' claims are hypotheses until then.

## Agent C (regression risk) — landed

### V1 CONFIRMED + ESCALATED — CP-space vs `char`-index drift (NEW, round 1 missed)

Blind claim: "no string in this codebase is CP-indexed", partly because
`String::from_utf16_lossy` collapses a surrogate pair from 2 CPs into 1 `char`.

Verified against the diff. `papx::build_paragraphs` does:

    let chars: Vec<char> = raw_text.chars().collect();
    let start = (cp_start as usize).min(chars.len());
    let end   = (cp_end   as usize).min(chars.len());
    let terminator = chars[end - 1];
    let text: String = chars[start..end - 1].iter().collect();

It indexes a `Vec<char>` with **character positions**. In a Word Unicode piece
one CP is one UTF-16 code unit (the PR's own `piece_byte_base` uses stride 2 to
say so), but `extract_text` builds the string via `String::from_utf16_lossy`,
so one astral character (emoji, CJK Ext-B, math alphanumerics) is 2 CPs and
1 `char`.

Demonstrated: for `"Hi <U+1F600> there"` the CP count is 11 and
`chars().count()` is 10 — drift of 1, permanent for the rest of the document,
and cumulative with each further astral char.

Why this is worse than a text-offset bug: `terminator = chars[end - 1]` is what
drives cell/row/paragraph grouping (`\x07` = cell/row mark). Once drift starts,
the terminator is read from the wrong position, so **table structure itself
corrupts**, not just the text. `.min(chars.len())` clamps instead of failing,
so it degrades silently.

Second, independent drift source, same root cause: `extract_text`
(piece_table.rs:144,154) **silently skips any piece** whose byte range does not
fit the buffer, while CP space still counts it.

Correct fix (this also subsumes round-1 F1): do not slice a flat string at all.
Resolve each paragraph's CP range against `&[Piece]` + `word_doc`, decode that
range directly, and run `sanitize_text` on the per-paragraph slice. Sanitising
the shared string is what forced the PR to use unsanitized `raw_text` in the
first place — the two bugs are the same bug.

### V2 NOT CONFIRMED — headers/footers/footnotes leaking into the IR

Blind claim: a PAPX walk enumerates the whole CP space, so header/footer/
footnote paragraphs would leak into the main-text IR.

Checked the diff: **the PR already guards this.** `build_paragraphs` filters
`.filter(|(cp, _)| *cp < text_len)` with `text_len = fib.text_len` (ccpText),
and `extract_text` is capped the same way. Claim does not hold. Recorded as an
example of why the agent's output must be diffed, not trusted.

### V3 CONFIRMED — soft line breaks change paragraph segmentation

`sanitize_text` maps 0x0B to `\n`, so the old line heuristic split a Word
paragraph containing soft breaks into several IR paragraphs. The structured
path keeps 0x0B as a literal control char inside one paragraph. Behaviour
change plus raw control character in output. Same family as round-1 F1.

### V4 CONFIRMED — pre-existing panic, owned by no PR

`parse_plc_pcd` (piece_table.rs) reads `cp_start`/`cp_end` straight from the
file with **no monotonicity validation**. `extract_text` then computes
`piece.cp_end.min(max_chars) - piece.cp_start` — unchecked `u32` subtraction.
A crafted `.doc` with `cp_end < cp_start` underflows: panic in debug/fuzz,
silent wrap in release (`[profile.release]` sets no `overflow-checks`).
Reachable from `Document::from_reader`, which the fuzz target does call.
This is on `main` today and is not #116's fault — file separately.

### V5 CONFIRMED, with a correction to the agent

Blind claim: "the fuzz target never calls `to_ir()`, so this has zero fuzz
coverage." The first half is right — `fuzz/fuzz_targets/fuzz_parse.rs` only
calls `Document::from_reader`, so `convert_doc.rs` (the new `walk_paragraphs`,
`build_table_rows`, span resolution) is never fuzzed.

But the agent overstated it. `DocDocument::parse` builds paragraphs **eagerly**,
so the PR's new binary parsers — `papx.rs` and `sprm.rs`, the parts that
actually touch untrusted bytes — *are* reached by the existing fuzz target.
The uncovered part is the pure-Rust IR assembly, which is the lower-risk half.

### V6 CONFIRMED — no `.doc` integration test existed before this PR

`main` has none. #116 adds three. Point in the PR's favour.

## Agents A (SOTA design) and B (SPRM spec) — still running

## Agent B (SPRM spec) — landed

### V7 CONFIRMED BLOCKER — sprmTDefTable length prefix is 2 bytes, not 1

This is round-1 F2, resolved. The blind agent, never told a PR existed, quoted
[MS-DOC] 2.2.5.1 verbatim: *"The first byte of the operand indicates the size
of the rest of the operand, except in the cases of sprmTDefTable and
sprmPChgTabs."* Per 2.9.321, sprmTDefTable's `cb` is a **2-byte LE** value =
remainder-of-structure + 1, so operand_len = cb + 1. Both Apache POI
(`SprmOperation.initSize` / `SPRM_LONG_TABLE`) and LibreOffice (`L_VAR2` in
`wwSprmParser::GetSprmTailLen`) implement the 2-byte form.

The PR's `sprm.rs` module doc asserts the opposite — that spra==6 SPRMs
"including sprmTDefTable = 0xD608 use a uniform 1-byte length prefix", and that
this "was verified empirically against the table.doc fixture".

Why the fixtures pass anyway: when cb < 256 the low byte of cb equals cb, and
the PR's `1 + prefix` total happens to equal the true `cb + 1`. The two
formulas coincide below the threshold. The contributor verified on a 3-column
table and generalised — the exact trap.

I verified the failure directly with spec-correct synthesised operands
(`cb = 22n + 4`):

    cols  cb   true_operand  prefix_byte  sprms_decoded  sentinel_found  cols_parsed
      2    48       49            48            2            true            2
      5   114      115           114            2            true            5
     11   246      247           246            2            true           11
     12   268      269            12           86            FALSE          None
     20   444      445           188           87            FALSE          None
     40   884      885           116          258         (garbage)         None

**Threshold is exactly 12 columns.** At 12+, the walker reads cb's low byte,
under-consumes by 256*cb_high, and decodes the remainder of the row's grpprl as
dozens of fabricated SPRMs. Consequences, both silent:
- the TAP fails to parse, so all merged-cell spans are lost (1x1 fallback);
- the fabricated opcodes can collide with the ones that drive structure
  (0x2416 fInTable, 0x460B ilvl, 0x6649 itap, 0xD608 itself), so paragraphs can
  be misclassified as in-table or as list items. At 40 columns a bogus
  sprmPFInTable was fabricated out of random bytes.

12-column tables are entirely ordinary. This is a blocker.

Fix: special-case 0xD608 (and 0xD606 sprmTDefTable10, same rule) to read a
2-byte cb and consume cb + 1. Note `parse_tdef_table` then needs its offsets
shifted by one, since its current "reserved byte" at operand[0] is really cb's
high byte.

### V8 CONFIRMED (LOW) — sprmPChgTabs 0xC615 has a 255-escape

Same spec sentence names a second exception. 0xC615 is a PAP sprm (sgc=1) so it
does appear in PAPX grpprls. When cb == 255 the real length must be computed
from the payload (`2 + 4*cTabsDel + 3*cTabsAdd`). The PR consumes 1 + 255 and
desyncs. Requires a paragraph with a great many tab stops, so low severity —
but it is the same class of bug and worth fixing alongside V7.
Agent's useful aside: POI itself routes 0xC615 through the 2-byte path, which
is wrong per spec — so copying POI here would not have saved us.

### V9 NOT CONFIRMED — vertical-merge "quirk" is real, PR is right

Blind agent describes TCGRF vertMerge as a clean 2-bit field at 0x0060
(1 = continuation, 3 = restart), which would make the PR's treatment of 0x0020
and 0x0040 as two independent flags a misreading, and its documented "Word
copies the whole row TAP so continuation rows keep the restart bit" comment a
rationalisation.

I dumped the actual fixture. Rows 1 and 2 of table-merges.doc both carry
rgf = 0x0060 (field value 3 = "restart") in column 0, yet POI's own HTML test
treats them as a single 2-row vertical merge. **The PR's observation is
correct and the clean spec reading does not match what Word wrote.** The PR's
logic also still handles the spec-conformant 0x0020 continuation encoding.
Not a defect. Known limitation worth documenting: two *distinct* vertical
merges stacked in the same column are indistinguishable under this encoding
and will be collapsed into one.

### V10 CONFIRMED (MEDIUM) — no tolerance when unifying column edges

Agent notes LibreOffice snaps cell boundaries within a tolerance
(`nTolerance = 4` twips) when building the unified column-edge array, because
Word rounds boundary positions per row. The PR's `build_table_rows` unifies
`rgdxaCenter` values by exact `sort`/`dedup` on `i32`.

Verified the fixture cannot catch this: every row's boundaries align exactly
(rows 1 and 2 are both `[0, 1062, 5738, 8148, 9302]`, row 0 is
`[0, 6872, 9302]`, row 3 is `[0, 9302]`). In real documents 1-3 twip
differences between rows are common; each one injects a spurious grid edge and
inflates `col_span` for every cell spanning it. Should snap within ~4 twips.

### Confirmed correct — no action

- **PAPX FKP entry length quirk.** Agent's rule (cb != 0 -> grpprlInPapx is
  2*cb - 1 bytes; cb == 0 -> re-read cb', size 2*cb') is exactly what the PR's
  `extract_grpprl` computes for both branches. Verified by hand.
- **Column-span algorithm.** The PR's "count grid edges in [centers[col],
  centers[col+1])" reproduces POI's `buildTableCellEdgesArray` result; hand-
  checked against the fixture's row 3 (single cell, grid
  {0,1062,5738,6872,8148,9302}, 5 edges in [0,9302) -> col_span 5, matching the
  PR's asserted ground truth K(5,1)).
