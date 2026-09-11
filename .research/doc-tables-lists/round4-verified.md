# Round 4 — verified findings (my own checks, primary source only)

Method: `claude_code_tricks/blind-agent-pr-review.md`. Six agents:
A1 vertical merges, A2 table structure, A3 robustness, A4 lists/tabs (all
solution-blind on a clean `main` archive); B on `change-07.diff` (= #116) and
B on `change-19.diff` (= #120 delta over #116), both implementation-blind.
Baselines were `git archive` snapshots — no `.git`, no `.research`, no PR refs.
`Regression (PR #116)` strings were scrubbed from the diffs before dispatch
(they are a search key straight back to the source — and an AGENTS.md rule 7
violation in their own right).

Nothing below is recorded on an agent's say-so. Each line was re-checked
against the primary [MS-DOC] text or against reference-implementation source.

## Verified verbatim from [MS-DOC]

**Sprm, spra = 6** (Sprm page): *"Operand is of variable length. The first byte
of the operand indicates the size of the rest of the operand, except in the
cases of sprmTDefTable and sprmPChgTabs."*
→ The exception set is **exactly two**: `0xD608` and `0xC615`.
→ `0xD606` ("sprmTDefTable10"), special-cased in #116 as a 2-byte-`cb`
   exception, is **not** in it and appears in no property table. Unsourced.
→ `0xC60D` sprmPChgTabsPapx is **not** an exception — it follows the ordinary
   1-byte rule. Any code treating it as length-less is wrong.

**Paragraph Properties table** (verbatim rows):
- `sprmPIlvl` **(0x260A)**, ispmd 0x0A — *"An unsigned 8-bit integer that
  specifies the list level of the paragraph."* Values `0x0-0x8`; `0xC` = list
  skips this paragraph.
- `sprmPIlfo` **(0x460B)**, ispmd 0x0B — *"A 16-bit signed integer value that is
  used to determine which list contains the paragraph."*
  `0x0000` not in a list (and list formatting removed); `0x0001-0x07FE` 1-based
  index into `PlfLfo.rgLfo`; `0xF801` not in a list;
  **`0xF802-0xFFFF` = the negation of a 1-based index — still in a list**, with
  left/first-line indents preserved.
  → confirms round 3's N1. Adds what round 3 missed: the negated band is *in* a
    list. `2047`/`0x07FF` is not a sentinel in any band (round 1's error, already
    retracted in round 3).
- `sprmPChgTabsPapx` **(0xC60D)** → `PChgTabsPapxOperand`;
  `sprmPChgTabs` **(0xC615)** → `PChgTabsOperand`. Different operand types.

**TCGRF** field order: `A horzMerge (2 bits)`, `B textFlow (3 bits)`,
`C vertMerge (2 bits)` → vertMerge is bits 5-6, mask `0x0060`.

**VerticalMergeFlag**: `fvmClear` 0x00, `fvmMerge` 0x01, `fvmRestart` 0x03 —
*"The contents and formatting of this cell extend down into any consecutive
cells below it that are set to the fvmMerge value."*
→ Raw bits: `fvmMerge` = `0x0020`, `fvmRestart` = `0x0060`.
→ A run continues **only** on `fvmMerge`. Anything else — including another
  `fvmRestart` — terminates it.

**Overview of Tables (2.4.3)**: *"The properties of each row mark MUST define
the cells for that table row. SprmTDefTable and sprmTInsert are used to create
cell definitions... There is no requirement that each row of a table have the
same number of cells."* And: *"Cells can be vertically merged... The second and
subsequent cells in the merged group MUST NOT contain any content other than
their cell marks."*

**Appendix A note <9>** (verbatim): *"Word 97 and Word 2000 require that each
row have sprmTDefTable applied... Word 2002, Office Word 2003, Office Word 2007,
Word 2010, and later require sprmTDefTable or sprmTInsert."* Plus: later Word
versions *"only emit sprmTDefTable for versions that do not process
sprmPTableProps"*, and omitting it makes documents *"not compatible with Word 97
or Word 2000"*.
→ **Trims an agent claim.** A2 argued modern Word files would lose structure
  because properties move to `sprmPTableProps`. Not so: Word still emits
  `sprmTDefTable` for back-compat. The real exposure is narrower — files from
  **non-Word producers** that define cells with `sprmTInsert` alone are
  spec-legal and lose every row. Document as a limitation; not a blocker.

## Verified against the reference implementation

Apache POI `AbstractWordConverter.getNumberRowsSpanned`:
```java
if (!nextCell.isVerticallyMerged() || nextCell.isFirstVerticallyMerged()) break;
```
`isFirstVerticallyMerged()` → `TableCellDescriptor.isFVertRestart()` (the
`0x0040` bit), which a raw `0x0060` cell has set. **POI stops the run at a
second restart**, i.e. POI agrees with the spec reading.

## F1 — #116's `is_quirk_continuation` is wrong (BLOCKER)

`convert_doc.rs`: `resolve_row_spans` / `is_merge_continuation` /
`is_quirk_continuation` treat an `fvmRestart` cell as a *continuation* whenever
the cell directly above is also `fvmRestart`, on the stated grounds that *"Word
sometimes copies the entire row TAP onto continuation rows."*

1. Contradicts the VerticalMergeFlag sentence above.
2. Contradicts POI, the implementation the PR cites as its model.
3. **The evidence no longer exists.** The comment cites `table-merges.doc`,
   deleted in the fix round. The two tests embedding real captured bytes show
   row 0 `rgf = 0x0000` and row 1 `rgf = 0x0060` — clear, then restart: exactly
   what the plain spec model predicts. Nothing surviving shows two consecutive
   `0x0060` rows.
4. **Wrong on its own terms.** `[fvmRestart, fvmRestart, fvmMerge]` — a
   standalone cell above a genuine 2-row merge — chains transitively into a
   single `row_span = 3` on row 0; row 1's cell content is dropped. Correct
   output is span 1, then span 2.
5. A test pins the wrong behaviour: `[0x0060, 0x0060, 0x0060]` is asserted to
   produce one 3-row merge.

**Provenance: mine.** Round 1 told the contributor the two-flag reading was
right and the "tidy spec reading" was not what Word writes, citing a byte dump
of a fixture that has since been deleted. Round 3 retracted the reading but the
heuristic it spawned survived into the code. Fix = delete the heuristic and the
test that pins it; the run continues only on `fvmMerge`.

## F2 — on `main`, owned by no PR

- `src/cfb/reader.rs:203` `read_fat` — DIFAT chain loop has **no cycle guard**,
  unlike `read_chain` two functions below (which carries `visited`/`max_sectors`).
  `header.difat_sector_count` (parsed `header.rs:97`) is never read. A DIFAT
  sector whose trailing `u32` points at itself loops forever, allocating a
  sector buffer per pass. AGENTS.md rule 6 names this case verbatim: *"bogus CFB
  FAT chains ... must surface as `Err`, not a crash."*
- `src/doc/piece_table.rs:136` — `piece.cp_end.min(max_chars) - piece.cp_start`,
  unchecked `u32` subtraction over unvalidated `aCP`. `[profile.release]` sets no
  `overflow-checks`: panics under test, wraps in release. Round 3 filed this;
  still unfixed; independently rediscovered from a different symptom this round.
- `fuzz/` has no corpus, no dictionary, and **no CI job** (`grep -rn fuzz
  .github/` is empty). "Extend the fuzz target" has been landing into something
  that never runs.

## F3 — #116 merges more than the page shows

Branch = `main` + `b5819b1` ("chore(deps): bulk-refresh all dependencies (#99)",
unmerged) + the feature commit. GitHub renders the PR against `b5819b1`, so the
13-file view hides it; merging the branch lands 26 files including 6 workflow
files, `Cargo.lock`, `js/package-lock.json`, `bench_rust/Cargo.lock`, a `.csproj`.

## Open — awaiting the two implementation-blind reports

## F4 — reachable panic in #116 (BLOCKER, reproduced by me)

`vert_merge_state(rgf) = ((rgf >> 5) & 0x03)` yields **four** values; the match
in `resolve_row_spans` handles `FVM_CLEAR`/`FVM_MERGE`/`FVM_RESTART` and closes
with `_ => unreachable!("vert_merge_state masks to 2 bits")`. The comment states
the reason it is wrong: 2 bits is four values, not three. `vertMerge == 0x02`
(raw `rgf & 0x0060 == 0x0040`) is representable, is written by nothing in the
`VerticalMergeFlag` enumeration, and is therefore exactly what a damaged or
hostile file contains.

Reproduced (probe added to the change's own test module, its own `vmerge_row`
helper):
```
thread 'convert_doc::tests::probe_vertmerge_value_two_does_not_panic' panicked at
src/convert_doc.rs:309:18:
internal error: entered unreachable code: vert_merge_state masks to 2 bits
```
Reachable from `Document::from_reader` → `to_ir()`. Violates AGENTS.md rule 6.
Independently found by the implementation-blind round from the diff alone.
Fix: treat the undefined value as `fvmClear` (a total function over 2 bits), not
`unreachable!`.

## F5 — F1 confirmed by execution, and it loses content

Probe on the change's own helper, `[fvmRestart, fvmRestart, fvmMerge]`:
```
PROBE B: spans=[[3], [], []] cells_per_row=[1, 0, 0]
```
One 3-row span; rows 1 and 2 emit **no cells at all**. Spec-correct output is
row 0 span 1, row 1 span 2, row 2 absorbed. The middle cell's content is dropped
from the IR while `plain_text()` still contains it — invisible to a
golden-output suite. Reached independently by the implementation-blind round.

## F6 — test hygiene, verified by reverting hunks

`decode_cp_range_unbacked_range_is_empty` is the **only** test covering the
**only** allocation bound the change adds. Reverting just the clamp
(`seg_end = seg_end.min(backed_cp);`) and re-running it:
```
test doc::piece_table::tests::decode_cp_range_unbacked_range_is_empty ... ok
```
It passes without the fix, in 0.76s — it iterates all 20M declared CPs and
asserts only that the result is `""`, which is true either way. The assertion
cannot see the cost it exists to bound.

The implementation-blind round reports 3 of 23 added tests green without the fix
they name; the other two are `extract_grpprl_handles_word8_reread` (uses `cw = 6`,
so it never enters the `cw == 0` re-read branch it is named for — already noted
in round 3 as N2) and a `fib` test that writes and reads the same offsets.

## F7 — FKP-walk amplification in #116 (measured by me)

`parse_papx_paragraphs`: `n = (lcb - 4) / 8` BTEs, each may name the **same**
page, and the page is re-parsed per BTE. No cap on `n`.
```
16 KiB table stream (  2047 BTEs) ->  40940 FkpParagraphs   7.5ms
64 KiB table stream (  8191 BTEs) -> 163820 FkpParagraphs  31.1ms
256 KiB table stream ( 32767 BTEs) -> 655340 FkpParagraphs 124.6ms
```
Linear, ~2.5 `FkpParagraph` per input byte (20 per BTE / 8 bytes per BTE), each
carrying a heap `Vec` grpprl. A 4 MiB table stream extrapolates to ~10.5M
paragraphs — arithmetic on the table above, not a measurement. Matches the
implementation-blind round's independently measured count.

## G — #120 (lists/tabs), all reproduced by me

**G1 (BLOCKER) — `0xC615`/`0xC60D` do have a `cb` byte.**
`is_pchg_tabs_no_prefix` declares both prefix-less, citing "[MS-DOC] §2.4.1".
Primary source:
- *PChgTabsOperand*: *"cb (1 byte): An unsigned integer that specifies the size
  of the operand... A value that is less than 255 specifies the size of the
  operand in bytes, not including cb. A value of 255 specifies... `4 ×
  PChgTabsDelClose.cTabs + 3 × PChgTabsAdd.cTabs`."*
- *PChgTabsPapxOperand*: *"cb (1 byte): ...the size of the operand in bytes, not
  including cb."*
So `sprmPChgTabs`'s exception is the `cb == 255` **escape**, not the absence of a
prefix; `sprmPChgTabsPapx` is not an exception at all (the Sprm page names only
`sprmTDefTable` and `sprmPChgTabs`).

Probe — spec-correct grpprl `[sprmPChgTabs cb=5 <operand>][sprmPFInTable=1]`:
```
PROBE: decoded 1 sprm(s):
   opcode=0xC615 operand=[05, 00, 01, D0, 02, 00, 16, 24, 01]
PROBE: tabs=[]
PROBE: f_in_table=false
```
The walk swallows the trailing SPRM into the tab operand: no tab stops, and
**the paragraph loses its table membership**. Word emits Prls in ascending sprm
order, so `0xC60D`/`0xC615` precede `0xD608` — this silently destroys the row
definition #116 exists to recover. Independently found by the
implementation-blind round, which measured `tap = Some(..)` → `tap = None`.

**G2 (BLOCKER) — list membership keyed on the wrong field/values.** Probe:
```
PROBE ilfo= 0x0000 -> emitted as LIST     (spec: NOT in a list)
PROBE ilfo= 0xF801 -> emitted as LIST     (spec: NOT in a list)
PROBE ilfo=   2047 -> emitted as prose    (not a spec value at all)
```
The only excluded value is the invented one. **Provenance: mine** — round 1's
`ilfo == 2047` note, retracted in round 3, is still in the tree. Also `ilfo` is
read as `u16` though the spec says *"A 16-bit signed integer value"*, and the
`0xF802-0xFFFF` negated-index band (still in a list) is unhandled.

**G3 — `0xC60D` decoded with the wrong entry width.** It carries
`PChgTabsDel` = *"cTabs (1 byte)"* + `rgdxaDel` (cTabs × 2-byte XAS) = `1 + 2n`;
the code uses the `PChgTabsDelClose` `1 + 4n` form for both opcodes.

**G4 — two added tests pass without the change.** Lifted verbatim into the
unpatched #116 baseline:
```
test doc::sprm::tests::sprm_tcell_padding_opcode_0xD632_does_not_populate_tabs ... ok
test convert_doc::tests::pchg_tabs_surfaced_on_paragraph ... ok
```
(the second hand-builds `PapProps { tabs: .. }` and never reaches the decoder).

**G5 — a comment regression.** #120 replaces a correct comment
(*"`parse_plc_pcd` validates CP ordering (piece_table.rs:119)"*) with a false one
(*"currently does not validate CP ordering"*) — the code does validate, at
`piece_table.rs:116-121`.

**G6 — fixture/parser collusion recurs.** The new `0xC615` fixtures omit the
mandatory `cb` byte, matching the parser's misreading: green by construction.
This is the same class round 3 named, in the change that fixes round 3.

## H — separate issue, owned by no PR: `ir::build_nested_list` drops items

`src/ir.rs` `build_nested_list`: when every item's level is greater than
`base_level`, the first item consumes the whole run as its "nested" range but
`*level <= base_level` is false, so `nested` is `None` and the rest are skipped.
Both callers pass a hard-coded `base_level = 0`
(`convert_docx.rs:644`, `convert_pptx.rs:287`).

Measured on **`main`**, unmodified:
```
PROBE: 3 items in -> 1 out
```
End-to-end through the PPTX path — a text body whose bullets all sit at outline
level 1, which is ordinary PowerPoint:
```
PROBE PPTX: 3 bullets in -> 1 list item(s) out
```
Two of three bullets silently lost from shipped PPTX extraction. Found by an
agent reviewing a `.doc` PR — code no PR touches. File independently; **not**
contributor review output.

# Round 4b — re-verification of the contributor's fixes (2026-08-27)

#116 head `3d7f14b0`, #120 head `04bd9392` (correctly rebased on #116).
Every item below re-run against the new heads, not read.

## #116 — the three blockers are fixed

| item | probe result |
|---|---|
| F4 panic on `vertMerge == 0x02` | `RECHECK A: no panic; rows=1 cells=[1]` — undefined value treated as clear |
| F1/F5 quirk heuristic | `[R,R,M]` → `spans=[[1], [2], []]`, `cells_per_row=[1,1,0]` — spec-correct, content preserved. `[R,R,R]` → three independent cells. `is_quirk_continuation` deleted |
| F7 FKP amplification | 256 KiB / 32767 BTEs: **655340 → 20** FkpParagraphs; time flat ~12µs |
| rule 7 PR numbers | clean |
| gate | `cargo fmt` clean; 475/475 lib tests pass |

New test `restart_after_restart_is_distinct_merge` asserts the exact pattern
that was broken, with strong assertions — it would catch a regression.
New test `papx_fkp_walk_is_bounded` **is** falsifiable: reverting the fix (both
the `n.min(max_pages)` clamp and the visited-page set — one logical change)
gives `left: 1000, right: 1`. (Reverting only the dedupe leaves it green because
the fixture's 512-byte `word_doc` makes `max_pages = 1`; that is a fixture
artefact, not a hollow test.)

## #116 — two items still open

**R1 — `is_two_byte_len_prefix` now over-corrects.** It reads
`0xD608 || 0xC615`. But the two are exceptions in *different* ways:
`TDefTableOperand.cb` is 2 bytes (§2.9.321), while `PChgTabsOperand.cb` is
**1 byte** — *"cb (1 byte): ...A value that is less than 255 specifies the size
of the operand in bytes, not including cb"*. The doc comment cites §2.9.321 for
both. Harmless inside #116 (for `cb < 255` both readings consume `cb + 1`, so the
walk stays in sync — verified: 2 sprms, `f_in_table=true`), but the operand slice
starts one byte late, which is what kills tab decoding in #120, and the
`cb == 255` escape still desyncs.

**R2 — `decode_cp_range_unbacked_range_is_clamped` still cannot fail.** Renamed
and given a second (partially-backed) assertion, but reverting
`seg_end = seg_end.min(backed_cp)` leaves it green (0.60s — still iterating all
20M declared CPs). Both assertions hold with or without the clamp. Suggest
extracting the backed-CP computation into a pure helper and asserting on it.

Cosmetic: `src/doc/papx.rs:334` still cites `table.doc`, which is not in the tree.

## #120 — rebased, comment fixed, blockers still open

| item | probe result |
|---|---|
| G5 CP-ordering comment | **fixed** — correct wording restored |
| hollow `0xD632` test | removed |
| G1 `0xC615` | walk no longer desyncs (`sprms=2`, `f_in_table=true`) **but `tabs=[]`** — the spec-correct tab stop at 720 twips is not decoded (operand off by one, per R1) |
| G1 `0xC60D` | **unchanged**: `sprms=1`, `tabs=[]`, `f_in_table=false` — still swallows the following SPRM |
| G2 `ilfo` sentinels | **unchanged**: `0x0000 → LIST`, `0xF801 → LIST`, `2047 → prose` |
| G3 `0xC60D` entry width | **unchanged** — still the `1+4n` DelClose form; spec says `PChgTabsDel` is `1+2n` |

Expected: he was asked to land #116 first and rebase #120, which is exactly what
he did. #120's substance is still to come.
