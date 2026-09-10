# PR #116 — round 3 verdict (blind review of the post-review state)

Method: `claude_code_tricks/blind-agent-pr-review.md`.
Ran after the contributor pushed `281268d` + `0a0631c` in response to
`VERDICT.md` (round 1 + 2). Four fresh non-fork agents, none told a PR
existed, none able to read each other or `.research/`.

| round | what it saw | baseline |
|---|---|---|
| A-widetable | neutral symptom: "wide tables lose merges, paragraphs gain properties that aren't in the document" | `621a804` (pre-fix) |
| A-textfidelity | neutral symptom: "field text and control chars in IR; emoji garbles table structure" | `621a804` (pre-fix) |
| B-change11 | `change-11.diff` = whole feature, no title/description | `b5819b1` |
| B-change24 | `change-24.diff` = the fix round only, no title/description | `621a804` |

Round C (this document) = me, comparing, verifying every load-bearing
citation against the primary source before recording it.

## The three round-1 blockers are genuinely fixed

- **B1 `sprmTDefTable` 2-byte `cb`** — fixed, and fixed *upstream* (in the
  walker, not by making `parse_tdef_table` tolerant). Ships a ≥12-column
  regression test and a round-trip walker test. B-change24 independently
  confirmed this hunk is spec-correct and called it the part worth keeping.
- **B2 CP-misaligned/unsanitized IR text** — fixed structurally via a new
  `decode_cp_range` that decodes per paragraph off the piece table. Both
  new tests genuinely fail on baseline (B-change24 verified by reverting the
  hunk). **But it introduced a new panic — see N3.**
- **B3 `cargo fmt`** — clean (verified).
- Mediums M1 (edge tolerance) and M3 (nested-table notice) were addressed
  unprompted.

That is a good-faith, competent response to review. The problem is what the
blind rounds found *underneath* it.

## New blockers

### N1 — the SPRM opcode table is wrong, and the fixtures encode the same error

Verified by me against [MS-DOC] "Paragraph Properties" and "Table Properties":

| code says | opcode | actually is |
|---|---|---|
| `sprmPIlvl` | `0x460B` | **`sprmPIlfo`** (16-bit *signed*, which list) |
| `sprmPIlfo` | `0x460D` | **not a paragraph SPRM at all** |
| — | `0x260A` | the real `sprmPIlvl` (1-byte level) — never read |
| `sprmPChgTabs` | `0xD632` | **`sprmTCellPadding`** (cell margins, `CSSAOperand`) |
| `sprmPChgTabs` old | `0xD634` | **`sprmTCellPaddingDefault`** |
| — | `0xC615` / `0xC60D` | the real `sprmPChgTabs` / `sprmPChgTabsPapx` — never read |

Consequences:
1. List *nesting* is driven by the LFO index, not the level. Real
   multi-level lists flatten; distinct lists look like distinct levels.
   B-change11 proved both directions empirically: a genuine `sprmPIlvl`
   (`0x260A`) produces **no list at all**, and `ilfo = 0xF801` ("not in a
   list" per spec) produces **a 2-item bulleted list**.
2. `props.ilfo` can never be set — `0x460D` has `spra = 2` (2-byte operand)
   but the arm requires `len() >= 4`. So the new `ilfo != Some(2047)` guard
   is unreachable dead code, and its unit test passes only because it
   hand-builds `PapProps` instead of going through `extract_pap_props`.
   (Confirmed independently by B-change24's probe.)
3. The `2047` sentinel is invented. Spec: "not in a list" is `0x0000` and
   `0xF801`; `0x0001–0x07FE` are valid indices. **This one is my fault** —
   the round-1 follow-up note described `ilfo == 2047` as an ANLD trapdoor
   and the contributor reasonably turned that into a "not in a list"
   exclusion. The note was about Word 6/95 ANLD numbering (implementation
   practice), not a spec sentinel, and I did not say so clearly.
4. Tab-stop extraction decodes cell-margin operands as tab stops.
   B-change24's probe: `TabStop { position_twips: 783, alignment: Bar,
   leader: MiddleDot }` fabricated out of a `CSSAOperand`. The operand
   *body* parsing is faithful; only the dispatch is wrong.
5. **The tests cannot catch any of it.** `tests/common/mod.rs::list_grpprl`
   emits `0x0B 0x46` — the same wrong opcode the parser reads — and its
   doc comment openly documents the `0x460D` inconsistency and works
   around it by omitting the field. Generator and parser share one
   misreading, so the suite is green by construction.

The PR description's "bug fix surfaced by the fixture: Word writes this
list's `ilvl = 1`, so `flush_list` now uses the run's minimum" is that
misreading's fingerprint: the `1` was the 1-based `rgLfo` index.

### N2 — `PapxInFkp`'s `cb == 0` branch is one byte short

Both round-A agents reached this independently, from different starting
symptoms. Verified verbatim: *"After cb', there are 2×cb' more bytes in
grpprlInPapx. The bytes after cb' form a GrpPrlAndIstd."* → grpprl ends at
`p + 1 + 2cb'`; `papx.rs` uses `p + 2cb'`. The `cb != 0` path is correct.
Effect: the last SPRM byte is dropped, so when `sprmTDefTable` is last in
the grpprl the TAP fails to parse and merges vanish — a width-independent
second route to the same 1×1 symptom B1 fixed.

The test named `extract_grpprl_handles_word8_reread` uses `cb = 6`, so it
takes the **other** branch and never enters the re-read path it is named
for.

### N3 — new robustness defects in the fix round (AGENTS.md rule 6)

- `decode_cp_range`: `Vec::with_capacity((seg_end - seg_start) as usize)`
  underflows on a non-monotonic piece (`parse_plc_pcd` still validates
  nothing). **Reproduced by me:** `attempt to subtract with overflow` at
  `piece_table.rs:204`. Release has no `overflow-checks`, so it wraps to a
  ~8 GB reservation instead.
- It also reserves and loops per *declared* CP before any check against
  `word_doc.len()`. B-change24 measured 42.1 s / ~8.6 GB from a 64-byte
  stream, 2.47 s reachable through `build_paragraphs`; the baseline was
  O(1) on the same input. *(agent-measured, not reproduced by me)*
- B-change11: multiply-overflow at `papx.rs:197`, and a 1 MB crafted
  `0Table` yielding 3.77 M paragraphs / ~1.38 GB. *(agent-measured)*
- `fuzz/` was not extended, though the PR adds two new parsers of untrusted
  input. Rule 6 says to extend it.

### N4 — `vertMerge` is a 2-bit enum, not two flags — and round 1 got this wrong

**Correction to `VERDICT.md` / my round-1 comment, recorded in place.**

Round 1 contained a section titled "One thing you got right that I initially
thought was wrong", telling the contributor that the tidy spec reading of
`vertMerge` as a 2-bit field was *not* what Word writes, and that treating
`0x0020`/`0x0040` as independent flags was correct.

That was wrong. Verified verbatim:
- TCGRF: *"C - vertMerge (2 bits): A value from the VerticalMergeFlag
  enumeration"*, at bits 5–6 → mask `0x0060`.
- VerticalMergeFlag: *"This MUST be one of the following values"* —
  `fvmClear 0x00`, `fvmMerge 0x01`, `fvmRestart 0x03`.

So `fvmRestart` **is** `0x0060` — it sets both bits the code reads as
separate flags. My round-1 observation ("continuation rows carry the restart
bit too") was just `fvmRestart` seen through the wrong model; it needed no
Word-quirk explanation. `0x0040` alone is value `0x02`, which the
enumeration does not define.

B-change11 and B-change24 reached this independently of each other and of
me. B-change11 further reports that two stacked vertical merges cause **cell
content to be silently deleted** *(agent-measured, not reproduced by me)*.

I cannot now re-check my round-1 byte dump: the fix round deleted
`table-merges.doc`, while `sprm.rs` and `convert_doc.rs` still carry
comments citing it as evidence. That is its own problem — the justification
outlived the evidence.

## Smaller, verified

- **2 new clippy warnings** from the PR's own tests (`convert_doc.rs:539`,
  `:637`). Caveat the agent didn't have: on this toolchain the baseline also
  warns (`src/cfb/reader.rs:43`, `src/ppt/text.rs:379` — files this PR never
  touches), so "fails the gate" is toolchain drift, not purely the PR's doing.
- **Review identifiers in code** — `"Medium #1"`, `"Medium #2"`, `"Medium #4"`
  in `convert_doc.rs` test names/comments. AGENTS.md rule 7 says name tests
  by defect *class*.
- **Comments cite deleted fixtures** — `table-merges.doc`, `simple-list.doc`,
  `table.doc` referenced in `src/` after the fix round removed them.
- **`sprmPFTtp` (`0x2417`) is decorative** — present in every fixture, read by
  nothing; row marks are keyed on `0xD608` instead. Rows defined by
  `sprmTInsert` alone (spec-legal: *"MUST specify at least one cell using
  sprmTInsert or sprmTDefTable, or a combination thereof"*) collapse silently.
  Reached by both round-A agents and B-change11.
- **Bundled scope** — the fix round carries seven logical changes plus a
  repo-wide rustfmt normalization; tab stops and soft-line-break handling
  were not requested by review.

## What the pairing bought that one round could not

- **A-widetable and A-textfidelity, from opposite symptoms, both landed on
  the `PapxInFkp` off-by-one and the unread `sprmPFTtp`** — neither could
  see the other, and neither was looking for it.
- **A-textfidelity corrected round 1's own B2 framing**: `sanitize_text`
  strips only field *delimiters*, never the instruction text, and its own
  test pins that. So `HYPERLINK "…"` leaks on **both** paths — the flat one
  too. Parity with the flat path was never a full fix, and I said it was.
- **Both B rounds independently reversed round 1's "you got this right"** on
  vertical merge.
- The opcode-table defect (N1) sits in code round 1 read closely and passed.

## Verdict

**Request changes.** Not mergeable as it stands, and the reason is one level
up from any individual defect: **the opcode table was never checked against
the specification, and every fixture is generated from that same table**, so
the suite is structurally incapable of falsifying it. Three separate features
(lists, the `ilfo` exclusion, tab stops) rest on it.

Recommended shape:
1. **Land the table path** once N2 and N3 are fixed — B1's fix is correct and
   the reconstruction is genuinely good work.
2. **Cut lists and tab stops from this PR entirely.** They are not narrow;
   they are decoding fields that aren't the fields they name.
3. **Add a gate that retires the class**: assert every opcode constant
   against a spec-derived table, and require one test built from *captured
   real bytes* rather than the in-code generator.
4. Fix N3 and extend `fuzz/` before any of it lands.

---

# Round 3b — verification pass (all measurements my own)

Ran after the write-up above, to replace agent-reported numbers with
reproduced ones and to audit the whole opcode table rather than the three
opcodes that happened to surface.

## Agent claims, re-tested

| claim | source | result |
|---|---|---|
| `fc_to_cp` underflows on a non-monotonic piece | B-change11 | **CONFIRMED** — `attempt to subtract with overflow`, `papx.rs:197` |
| `fc_to_cp` overflows on a large declared range | B-change11 | **CONFIRMED** — `attempt to multiply with overflow`, same line |
| `decode_cp_range` underflows | round C | **CONFIRMED** — `piece_table.rs:204` |
| `decode_cp_range` cost is driven by declared CPs, not stream size | B-change24 | **CONFIRMED** — see below |
| stacked vertical merges lose cell content | B-change11 | **CONFIRMED, and worse than reported** |

### `decode_cp_range` amplification — measured

64-byte stream throughout; only the *declared* CP count varies:

| declared CPs | output | elapsed |
|---:|---:|---:|
| 5,000,000 | 0 chars | 55.6 ms |
| 10,000,000 | 0 chars | 100.7 ms |
| 20,000,000 | 0 chars | 222.1 ms |

Linear in the declared count, ~11 ns/CP, **and it produces nothing** — every
offset is past the end of the 64-byte stream, so the whole cost is waste.
Extrapolated to the `u32` ceiling (4.295 e9 CPs): ~47 s and an 8.6 GB
reservation (2 bytes/CP), which corroborates B-change24's independently
measured 42.1 s / 8.6 GB. I have not run that extreme myself — the
extrapolation is arithmetic on the table above, not a measurement.

### Stacked vertical merges — measured

Four single-column rows, spec encoding: `A` `fvmRestart`(0x0060),
`B` `fvmMerge`(0x0020), `C` `fvmRestart`(0x0060), `D` `fvmMerge`(0x0020) —
i.e. two independent 2-row merges. Correct output: two cells, each
`row_span = 2`, both `A` and `C` rendered.

Actual: `emitted rows=4  spans=[[4], [], [], []]` — **one cell spanning all
four rows**; `A` survives, `C` is gone.

Precision, because it matters for the report: `B` and `D` disappearing is
*correct* — [MS-DOC] VerticalMergeFlag says an `fvmMerge` cell "contributes
its layout region to the set and its own contents are not rendered". The
defect is `C`, an `fvmRestart` cell whose contents MUST render, silently
absorbed because the code's two-flag model cannot tell `0x0060` (restart)
from "merge bit set".

## Full opcode audit

The three known-bad opcodes were not the whole story, so I checked every
opcode the decoder acts on, plus the size machinery.

**Correct, verified verbatim:**
- The `spra` → operand-size table (`0|1`→1, `2|4|5`→2, `3`→4, `6`→variable,
  `7`→3) matches [MS-DOC] "Sprm" exactly.
- `sgc` (bits 10–12) and `spra` (bits 13–15) extraction are right.
- `0x2416` sprmPFInTable, `0x6649` sprmPItap, `0xD608` sprmTDefTable,
  `0x563A` sprmTIstd, `0x2417` sprmPFTtp — all correct identities.

So the decoding *machinery* is sound. The defect is confined to opcode
identity — which is why the fixtures never caught it: the machinery they
exercise works.

**Wrong (4 of the 7 dispatched opcodes):**

| opcode | code calls it | MS-DOC says |
|---|---|---|
| `0x460D` | `sprmPIlfo` | not defined |
| `0x460B` | `sprmPIlvl` | `sprmPIlfo` |
| `0xD632` | `sprmPChgTabs` | `sprmTCellPadding` |
| `0xD634` | `sprmPChgTabs` old | `sprmTCellPaddingDefault` |

**Never decoded, though features depend on them:** `0x260A` sprmPIlvl,
`0xC615` sprmPChgTabs, `0xC60D` sprmPChgTabsPapx.

**Unsourced:** `0xD606` ("sprmTDefTable10") is special-cased as a 2-byte-`cb`
exception, but MS-DOC's exception sentence names exactly two properties —
*"except in the cases of sprmTDefTable and sprmPChgTabs"* — and `0xD606`
appears in no MS-DOC property table. Harmless if Word 97+ never emits it,
but it is a constant with no cited source.

Note the irony: `0xC615` sprmPChgTabs is one of only **two** documented
exceptions to the variable-length rule — the sibling of the B1 bug that
started this review — and it is still unhandled while the code special-cases
an undocumented opcode instead.

## The gate

`proposed-opcode-gate.rs` — two tests, ~80 lines, no I/O:

1. `dispatched_opcodes_match_their_claimed_spec_property` — every opcode the
   decoder acts on, checked against a transcribed [MS-DOC] name table.
2. `required_properties_are_decoded` — spec properties the features need but
   nothing reads.

Verified against the current head: fails with exactly the four mismatches and
the three missing properties, each named in one line. This is the piece that
*retires*: it costs nothing per run and catches every future instance,
including in code paths no fixture exercises.
