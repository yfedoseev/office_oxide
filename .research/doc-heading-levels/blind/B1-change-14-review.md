# Blind review — `change-14.diff` against `office_oxide` @ `5007d9c` (main)

Every correctness claim below is tagged **[specified]** (read in [MS-DOC] on
learn.microsoft.com, or in Apache POI / LibreOffice source fetched during this
review) or **[inferred]** (reasoned from the code in this checkout).

Read-only review. Nothing in the repository was modified; no build was run.

> **Leak note.** `grep` for `min(6)` across the working tree returned hits under
> `./.claude/worktrees/agent-a15412f1b9130f80a/src/...` — a sibling git worktree
> that already contains this diff applied. I did not open it and did not use it.
> Flagging it because it is a provenance leak sitting in the same filesystem.

---

## What defect I infer this is fixing

The `.doc` → IR path classifies headings with a *guess*. `convert_doc::emit_prose`
(baseline, `src/convert_doc.rs:444-481`) decides "heading" from the text itself —
shorter than 100 chars, doesn't end in `.` or `,`, and either all-uppercase or
"first element and shorter than 60 chars" — and then assigns level `1` if it is
the first element in the document and `2` otherwise. So a legacy `.doc` can never
produce an H3–H6, and a genuinely-styled `Heading 3` that happens to end in a
period, or to be long, becomes an ordinary paragraph.

The change wires up the real signal: parse the document style sheet (`STSH`) out
of the Table stream, map a paragraph's `istd` to a built-in `sti` 1–9 (or a
user-defined style named `Heading N`), and also read what it calls
`sprmPOutlineLvl` out of the paragraph's `grpprl`. The resolved level is carried
on `PapProps.heading_level` and consumed by a new first arm in `emit_prose`,
clamped to a new `ir::MAX_HEADING_DEPTH = 6`.

That is the right diagnosis and, structurally, the right place. The style-sheet
half is largely correct. **The SPRM half decodes the wrong opcode entirely**, and
that half is the one the flagship end-to-end test and five unit tests are built
around.

---

## Q1 — Would each added test fail on the baseline?

Mechanically, *none* of the added tests compile against `5007d9c`: they reference
`PapProps::heading_level`, `PapProps::style_istd`, `PapProps::outline_lvl_explicit`,
`FkpParagraph::istd`, `doc::styles::{StyleDef, parse_style_sheet, heading_level_for_istd}`,
`doc::MAX_OUTLINE_LEVEL`, `ir::MAX_HEADING_DEPTH`, and a 5-argument
`build_paragraphs`. So the literal answer is "they don't build".

The useful question is the counterfactual one: **with the change applied, would
the test still pass if the behaviour it names were reverted?** That is what the
"discriminating?" column answers. **[inferred]** throughout, from hand-execution
of the baseline and patched code.

### Group A — `src/convert_doc.rs` (3 tests)

| Test | What the code actually does | Discriminating? | Evidence |
|---|---|---|---|
| `styled_heading_inside_table_stays_a_cell` | `walk_paragraphs` dispatches on `is_table_trailing_mark` → `f_in_table` → `is_doc_list_item` **before** reaching `emit_prose` (`src/convert_doc.rs:377-395`). `add_cell_paragraph` (`:105-120`) unconditionally pushes `Element::Paragraph` and never looks at `heading_level`. | **No.** Passes with the `emit_prose` heading arm deleted. It also passes with the whole `styles.rs` module deleted. | The cell's `heading_level: Some(3)` is *never read* on this path — no code exists that could make it a `Heading`. |
| `styled_heading_on_row_terminator_is_not_a_heading` | `is_table_trailing_mark` is the first arm; the paragraph's `text` is dropped on the floor by `end_row` (`:123-130`), which never touches `text`. | **No.** Same reason. The comment claiming non-empty text "is what makes this test meaningful" is wrong — `end_row` ignores `text` either way. | `end_row` takes only `tap` and `itap`. |
| `styled_heading_list_item_stays_a_list_item` | `ilfo: Some(1)` → `is_doc_list_item` true → list arm, again before `emit_prose`. | **No.** | `:384-390`. |

All three are *guards against a change nobody made*. They are cheap and harmless,
but they contribute zero evidence that the fix works, and two of them (`inside_table`,
`row_terminator`) canonize a fidelity **loss** as desired behaviour — see Q6.

### Group B — `src/doc/document.rs` (4 tests)

| Test | What baseline produces | Discriminating? | Evidence |
|---|---|---|---|
| `ir_styled_heading_uses_real_level` | Paragraph 1 = `"Intro paragraph."`, paragraph 2 = `"Subsection"`. Baseline heuristic on `"Subsection"`: not all-uppercase, and `elements` is non-empty → `is_heading = false` → `Element::Paragraph`. Assertion `Heading{level:3}` **fails**. | **Yes.** This is the honest core regression test. | `emit_prose` `:449-457` baseline. |
| `ir_deep_outline_level_clamps_to_ir_max_depth` | Single paragraph `"Deep section"`, `elements.is_empty()` and `len < 60` → baseline emits `Heading{level: 1}`. Assertion `level == 6` **fails**. | **Yes**, for the clamp. | Same. |
| `synthetic_doc_styled_heading_uses_style_sheet_level` (**flagship**) | I traced the synthetic CFB by hand end-to-end (see Q5). It does parse: `elements = [Paragraph("Introduction."), <second>]`. Baseline: no style sheet is read, so `"Subsection Three"` hits the heuristic — not all-caps, `elements` non-empty → `Paragraph`. `match { Heading => .., other => panic! }` **fails**. | **Yes** — the style-sheet half genuinely rides through CFB → FIB → CLX → PlcfBtePapx → FKP → `parse_style_sheet` → IR. | Byte-level trace in Q5. |
| `synthetic_doc_outline_sprm_heading` | Same pipeline, but the "heading" comes from a `grpprl` of `12 64 05 00 00 00`. That opcode is **`sprmPDyaLine`**, not an outline sprm (Q3). The test passes only because the decoder and the fixture share the same wrong belief. | **Yes** against the baseline — and **worthless** as evidence, because it tests a fiction. | See Q3/Q5. |

### Group C — `src/doc/papx.rs` (10 tests)

| Test | Discriminating? | Note |
|---|---|---|
| `build_paragraphs_resolves_heading_from_style_istd` | Yes | Hand-built `StyleDef{sti:3}`; exercises `heading_level_for_istd` + plumbing. Sound. |
| `build_paragraphs_resolves_heading_from_parsed_style_sheet` | Yes | Best test in the diff: real spec-shaped STSH bytes → `parse_style_sheet` → `build_paragraphs`. |
| `build_paragraphs_resolves_user_heading_name_case_insensitive` | Yes | Tests the name heuristic, which is *not* a spec rule (Q3). |
| `build_paragraphs_leaves_body_text_unheaded_with_styles_present` | Yes | Genuine negative test. |
| `build_paragraphs_prefers_sprm_style_over_papx_istd` | **Vacuous in production.** | grpprl `0A 64 03 00 00 00`. `0x640A` is not a defined MS-DOC sprm (Q3); the real one is `sprmPIstd = 0x4600`. This arm can never fire on a real file, so the test only proves the code matches its own invented opcode. |
| `build_paragraphs_sprm_style_override_can_remove_a_heading` | Same — vacuous | Same fabricated `0x640A`. |
| `build_paragraphs_resolves_heading_from_sprm_outline_lvl` | **Vacuous / actively wrong** | grpprl `12 64 05 00 00 00` = `sprmPDyaLine` with `dyaLine = 5`. |
| `build_paragraphs_sprm_outline_lvl_overrides_style` | Same, and encodes a precedence the spec inverts (Q6) | |
| `build_paragraphs_sprm_outline_lvl_zero_suppresses_styled_heading` | Same, and encodes a sentinel the spec puts at `9`, not `0` (Q3) | |
| `build_paragraphs_sprm_outline_lvl_accepts_deepest_level` | Same | `MAX_OUTLINE_LEVEL` as an operand of `sprmPDyaLine` is `dyaLine = 9` (0.45 pt line spacing). |

### Group D — `src/doc/sprm.rs` (5 tests)

| Test | Discriminating? | Note |
|---|---|---|
| `sprm_p_outline_lvl_sets_heading_level` | Yes vs. baseline, meaningless vs. the format | `12 64 02 00 00 00` = `sprmPDyaLine`, `dyaLine = 2`. |
| `sprm_p_outline_lvl_body_is_none` | Yes / meaningless | `dyaLine = 0`. |
| `sprm_p_outline_lvl_rejects_invalid_byte` | Yes / **misnamed** | The input `12 64 68 01 01 00` is a *perfectly valid* `sprmPDyaLine`: `dyaLine = 0x0168 = 360`, `fMultLinespace = 1` — i.e. **1.5× line spacing**. The test calls it a "coincidental byte sequence"; it is the single most common non-default line-spacing value in Word. This is the author brushing past the actual bug. |
| `pap_props_read_sprm_p_style_override` | Yes / vacuous | `0x640A` again. |
| `pap_props_style_istd_absent_when_no_sprm` | Trivially true | Would pass on any implementation. |

### Group E — `src/doc/styles.rs` (10 tests, all new module)

| Test | Discriminating? | Note |
|---|---|---|
| `parse_stsh_reads_sti_and_names` | n/a (new module) | The fixture **is** spec-shaped — I verified `cbStshi + Stshif(18) + rglpstd`, `cstd ≥ 0x000F`, `cbSTDBaseInFile = 0x000A`, fixed-index istd 0–9 → sti 0–9, istd 13/14 empty, `LPStd = cbStd(u16) + STD`, `STD = stdf + xstzName`, `Xstz = Xst + chTerm`. All **[specified]**. Good test. |
| `parse_stsh_rejects_spurious_name_table_layout` | Weak | Asserts `styles.iter().all(name.is_empty())`. In fact the parser bails at the first bogus `cbStd` (`u16::from_le_bytes([1,4]) = 0x0401 > remaining`) and returns an **empty** vec, so the assertion is *vacuously true*. It would also pass for a parser that always returns `vec![]`. |
| `parse_stsh_rejects_unexpected_cb_std_base_in_file` | Yes | Real, and the rule is **[specified]** (§2.9.274: MUST be `0x000A` or `0x0012`). |
| `parse_stsh_skips_odd_sized_lpstd_padding` | Yes | I traced it: removing the `pos += 1` padding step makes the next `cbStd` read `0x2000`, which overruns and breaks the loop → `len() == 1` → fails. Rule **[specified]** (§2.9.135). Strongest test in the diff. |
| `heading_level_resolves_builtin_and_user`, `heading_level_uses_primary_name_before_aliases` | Yes | Alias/comma rule **[specified]** (§2.9.258). |
| `truncated_style_sheet_is_empty`, `parse_style_sheet_absent_is_empty`, `parse_style_sheet_out_of_bounds_is_empty`, `parse_style_sheet_reads_stsh_at_fib_offset` | Yes | Sound bounds tests. |

**Summary of Q1.** Three of the four `convert_doc.rs`/routing tests are
non-discriminating. The flagship (`synthetic_doc_styled_heading_uses_style_sheet_level`)
**does** fail on the baseline and does exercise the style-sheet path end to end —
credit where due. But **eight tests** (`build_paragraphs_*_sprm_*`,
`pap_props_read_sprm_p_style_override`, `sprm_p_outline_lvl_*`,
`synthetic_doc_outline_sprm_heading`) validate a decoder against a fixture that
encodes the same wrong opcode; they cannot fail no matter how wrong the opcode is.

---

## Q2 — Count the arms

**`emit_prose`.** Baseline is a 2-arm conditional (`is_heading` → `Heading{1 or 2}`,
else `Paragraph`), where `is_heading` is itself a 5-clause heuristic with an
`||`-branch. The change prepends a 3rd arm. Crucially it **does not remove or gate
any of the five heuristic clauses**. So after the change a `.doc` that *has* a
parsed style sheet still runs the "short line, no trailing period ⇒ this is an H1"
guess on every paragraph the style sheet says is body text. That is the actual
upstream mistake — the parser guessing when it now has ground truth — and it
survives. The natural shape of the correct fix is: *when the style sheet parsed
non-empty, trust it and stop guessing* (fall back to the heuristic only when
`styles.is_empty()`). The change instead adds a bypass in front of the guess.

**`walk_paragraphs`.** 4 arms, untouched. Fine.

**`extract_pap_props`.** Baseline has 6 opcode arms; the change adds 2 more. Two of
the new arms (`0x640A`, `0x6412`) are for opcodes that do not mean what the arms
say, so this is 2 new dead-or-harmful arms rather than 2 new correct ones.

**Two arms compensating for the same upstream mistake?** Yes, in a mild form:
`resolve_heading_level` in `papx.rs` and the heuristic in `emit_prose` are now two
independent heading classifiers in the same pipeline, plus a third
(`convert_docx::resolve_heading_level`) for the OOXML path with a *different*
level convention (0-based there, 1-based here). The change documents the mismatch
in a comment instead of unifying it.

---

## Q3 — Every constant, opcode and offset

| Value | Claimed meaning in the diff | Verified meaning | Verdict |
|---|---|---|---|
| `0x6412` | "`sprmPOutlineLvl` … 4-byte operand's low byte is the outline level" | **`sprmPDyaLine`**, `ispmd 0x12`, `sgc 1`, `spra 3`. Operand is an **`LSPD`** (§2.9.146): `dyaLine` (i16) + `fMultLinespace` (u16). **[specified]** — [MS-DOC] §2.6.2 table row `sprmPDyaLine (0x6412)`; LibreOffice `sprmids.hxx:395` `using PDyaLine = sprmPar<0x12, 0, SPRA::operand_4b_3>; // 0x6412`; POI `ParagraphSprmUncompressor.java:161-162` `case 0x12: newPAP.setLspd(new LineSpacingDescriptor(...))`. | **WRONG.** Blind to: line spacing. The code reads the **low byte of `dyaLine`** and calls it an outline level. |
| the real outline sprm | — | **`sprmPOutLvl = 0x2640`**, `ispmd 0x40`, **`spra 1` ⇒ a 1-byte operand**, values `0x0–0x8` = **zero-based** outline level, `0x9` = **body text**. **[specified]** — [MS-DOC] §2.6.2; LibreOffice `sprmids.hxx:428` `POutLvl = sprmPar<0x40, 1, SPRA::operand_1b_1>; // 0x2640`; POI `ParagraphSprmUncompressor.java:314-321` `case 0x40: newPAP.setLvl(...)`. | Not implemented. Note the operand is **1 byte, not 4**, and `0` means *Heading 1*, not body text — the change's sentinel is at the wrong end of the range. |
| `0x640A` | "`sprmPStyle` … the 4-byte operand's low word is the style index" | **No such sprm.** `ispmd 0x0A` with `fSpec = 0` and `spra = 3` has no entry in the [MS-DOC] §2.6.2 paragraph table (`ispmd 0x0A` is `sprmPIlvl = 0x260A`, `fSpec = 1`, 1-byte). The real style sprm is **`sprmPIstd = 0x4600`**, `spra 2` ⇒ **2-byte** operand. **[specified]** — §2.6.2; LibreOffice `sprmids.hxx` `PIstd = sprmPar<0x00, 1, SPRA::operand_2b_2>; // 0x4600`. | **WRONG.** Dead arm on real files. Harmless in practice, but its tests and doc comments are fiction. |
| `MAX_OUTLINE_LEVEL = 9` | "The deepest outline level MS-DOC stores… The `sprmPOutlineLvl` operand, `StdfBase.sti`, and a user-defined `Heading N` style name all use this range" | Three different ranges are being conflated. `sprmPOutLvl` operand: `0x0–0x9`, where `9` = body text **[specified]**. `StdfBase.sti`: 12-bit, `0x0000–0x0FFD` plus `0x0FFE` = user-defined **[specified], §2.9.260**. `Heading N` name: `1–9` **[inferred]**. | Value `9` happens to be right for heading depth, doc comment is not. |
| `MAX_HEADING_DEPTH = 6` | markdown depth bound | Correct for markdown/HTML `h1..h6` **[specified, CommonMark/HTML]**. But `src/convert_docx.rs:201` still hardcodes `(level + 1).min(6)`, so the constant's claim that "format readers … clamp to this" is only half-true. | OK, incompletely applied. |
| FIB `0x00A2` / `0x00A6` (`fcStshf`/`lcbStshf`) | style sheet in the Table stream | `FibRgFcLcb97` starts at `0x9A` (FibBase 32 + csw 2 + fibRgW97 28 + cslw 2 + fibRgLw97 88 + cbRgFcLcb 2). `fcStshf` is pair index **1** ⇒ `0x9A + 8 = 0xA2`, `lcbStshf = 0xA6`. Consistent with the repo's existing `fcPlcfBtePapx` at pair 13 ⇒ `0x102`. **[specified]** — [MS-DOC] §2.5.6 field order `fcStshfOrig, lcbStshfOrig, fcStshf, lcbStshf, …`. | **Correct.** Guard `data.len() > 0x00A9` is also exactly right (needs `≥ 0xAA`). |
| `Stshif`: `cstd` @0 (u16), `cbSTDBaseInFile` @2 (u16), must be `0x000A` or `0x0012`, size exactly 18 | as claimed | **[specified]** — §2.9.274, verbatim: "This value MUST be 0x000A when the Stdf structure does not contain an StdfPost2000 … and MUST be 0x0012 when [it] does." | **Correct.** |
| `LPStshi = cbStshi(u16) + Stshif…`, `rglpstd` follows immediately | as claimed | **[specified]** — §2.9.271 `STSH = lpstshi + rglpstd`; §2.9.272 `STSHI = stshif(18) + ftcBi + StshiLsd + StshiB`. Skipping `cbStshi` bytes correctly skips the optional trailing members too. The "there is no style-name STTB" claim is right for Word 97+. | **Correct**, and the accompanying regression test is a real one. |
| `StdfBase.sti` = low 12 bits of first `u16`; `0x0FFE` = user-defined | as claimed | **[specified]** — §2.9.260. | **Correct.** |
| Fixed-index `istd 0–9 → sti 0–9`, `sti 1..9` = Heading 1..9 | as claimed | **[specified]** — §2.9.271 fixed-index table; and §2.6.2 `sprmPIstd`: "An **istd** value in the range of 1 to 9, inclusive, also specifies the outline level of the paragraph … equal to the value of the **istd** minus 1." | **Correct.** Note the spec ties it to *istd*, the change keys off *sti*; identical for fixed-index entries. |
| `LPStd` even-byte padding, `cbStd` excludes it | as claimed | **[specified]** — §2.9.135. | **Correct**, and tested. |
| `Xstz = cch(u16) + cch×u16 + chTerm(u16)`; name may be `"primary,alias"` | as claimed | **[specified]** — §2.9.354 (Xstz), §2.9.258 (STD.xstzName aliases). Section numbers cited in the diff are right. | **Correct.** |
| cited "MS-DOC §2.9.138" for the outline level | — | §2.9.138 is not the outline level; the outline-level operand is documented inline in §2.6.2 and there is no separate structure. §2.9.146 is `LSPD`. | **Bogus citation.** |
| PAPX header `istd` at `p+1 .. p+3` (post-reread) | as claimed | §2.9.175: `cb != 0` ⇒ `grpprlInPapx` is `2·cb − 1` bytes and *is* a `GrpPrlAndIstd`; `cb == 0` ⇒ `cb'` at +1, then `2·cb'` bytes forming a `GrpPrlAndIstd`. §2.9.114: `GrpPrlAndIstd = istd(u16) + grpprl`. Both forms put `istd` at `p+1` after the re-read adjustment. **[specified]** | **Correct.** Little-endian, unsigned, 2 bytes — all right. |
| CFB sentinels `0xFFFFFFFE/FF/FD`, sector shift 9, mini shift 6, DIFAT@`0x4C` | test fixture | **[specified]** — [MS-CFB]. | Correct as far as the reader consumes them; see Q5 for what it doesn't consume. |

### Real-world consequence of the `0x6412` error

`sprmPDyaLine` is written by Word into a paragraph's PAPX whenever the user sets
non-default line spacing. The change reads `operand[0]` = the **low byte of a
16-bit twip count** and accepts it as an outline level when `≤ 9`. Concretely
(**[inferred]** from LSPD semantics, **[specified]** for the encoding):

* "At least **13 pt**" ⇒ `dyaLine = 260 = 0x0104` ⇒ `operand[0] = 0x04` ⇒ the body
  paragraph becomes **`#### **text****`**.
* "At least **26 pt**" ⇒ `520 = 0x0208` ⇒ `0x08` ⇒ Heading 8, clamped to `######`.
* Multiple **1.07** ⇒ `257 = 0x0101` ⇒ Heading 1. Multiple **1.08** (the Word
  2013+ Normal default) ⇒ `259 = 0x0103` ⇒ Heading 3.
* "At least **12.8 pt**" ⇒ `256 = 0x0100` ⇒ `operand[0] = 0` ⇒ `outline_lvl_explicit = true`
  with `heading_level = None` ⇒ `resolve_heading_level` returns `None` **and never
  consults the style**. A paragraph genuinely styled `Heading 2` silently loses
  its heading.

So the SPRM half both fabricates headings from line spacing and suppresses real
ones. The `≤ 9` range check is not a safety net — it is precisely the window in
which `dyaLine mod 256 ≤ 9`.

---

## Q4 — Could a regression suite detect a regression here?

Largely **no**, and the change is close to unfalsifiable by the project's own
process. Reasons:

1. **The change is net-additive to output.** Paragraphs become Headings; markdown
   gains `###` prefixes and `**bold**`. A golden-output or text-diff corpus run
   scores "the document now has headings where it had flat prose" as the
   *intended improvement* — that is literally the PR's thesis. There is no way to
   tell a correct new heading from a line-spacing artefact by diff shape.
2. **AGENTS.md rule 5 asks for a corpus run, and the corpus is not distributed**
   ("bring your own"). The PR's own comments assert the corpus cannot reach the
   new code at all.
3. **The one detectable direction — suppression — is rare and quiet.** The
   `dyaLine ≡ 0 (mod 256)` case removes a heading; in a diff that reads as one
   fewer `##` line, which is exactly what a reviewer would expect from "we
   stopped over-firing the heuristic".
4. `fuzz/` is not extended, although the diff adds an entire new binary parser
   (`styles.rs`) fed from an attacker-controlled FIB offset. AGENTS.md rule 6
   explicitly asks for this.

**Adversarial input that exposes the failure.** Take any `.doc` with two
paragraphs; give the second one direct paragraph formatting "Line spacing: At
least 13 pt" and leave its style as `Normal`. Word writes
`12 64 04 01 00 00` into that paragraph's PAPX `grpprl`. Expected output: two
paragraphs. Output after this change: `Introduction.` followed by
`#### **Second paragraph**`. A one-line assertion — *"a Normal-styled paragraph
with non-default line spacing is not a heading"* — would fail, and no test in the
diff makes it.

A second, sharper one: a paragraph styled `Heading 2` **and** given "At least
12.8 pt" spacing. Correct output `## …`; actual output a plain paragraph, because
`outline_lvl_explicit` is set from `dyaLine`'s low byte and blocks the style
lookup. That case is the exact inverse of the change's stated goal, and the diff
contains a test (`build_paragraphs_sprm_outline_lvl_zero_suppresses_styled_heading`)
that *asserts this behaviour as correct*.

---

## Q5 — Does any fixture contain data the code under test never reads?

Yes, in three separable ways.

**(a) The `sprmPOutlineLvl` fixtures are pure self-confirmation.** Every one of
`12 64 xx 00 00 00`, `0A 64 xx 00 00 00`, and the `build_synthetic_outline_doc()`
grpprl `[0x12, 0x64, 0x05, 0x00, 0x00, 0x00, 0x00]` was written by the same mental
model as the decoder. Grepping the distinctive bytes against the format: `12 64`
is `sprmPDyaLine`; the fixture's `05 00 00 00` is `LSPD{ dyaLine: 5, fMultLinespace: 0 }`.
The comment in `build_synthetic_outline_doc` even reasons "the style-sheet path
would resolve istd 0 (Normal) to no heading at all … Level 5 therefore proves the
`sprmPOutlineLvl` path reached the IR end to end" — it proves the *plumbing*
reaches the IR, and nothing at all about the opcode. This is exactly the failure
mode of "a fixture built by the author's model of the format re-encoding the
decoder's assumptions."

**(b) The synthetic CFB carries decorative structure the reader never consumes.**
I traced `build_synthetic_styled_doc()` byte by byte against this repo's parsers
and it does work — header ✓, DIFAT[0]=1 ✓, FAT chains 2→3→4→5 and 6→7 ✓, FKP page
at `pn=2` ⇒ `wd[0x400]` with `crun` at `wd[0x5FF]` ✓, `rgfc` at page +0, `rgbx`
(13-byte `BxPap`) at page +12 with `bOffset` 19/21 ⇒ PAPX at page +38/+42 ✓, CLX
of 21 bytes yielding one Unicode piece at `fc 0x300` ✓, STSH of 364 bytes at
table offset `0x120` ✓. But:

* `write_dir_entry` writes `child`, `left sibling`, `right sibling` and `color`
  into every entry, and sets the Root's `child` to entry 1 while both stream
  entries have `FREE` siblings. In a real CFB that makes `1Table` **unreachable**
  from the root's red-black tree. It works here only because
  `CfbReader::find_entry` (`src/cfb/reader.rs:79-84`) is a **linear scan over all
  directory entries** and ignores the tree entirely. Those 16 bytes per entry are
  decorative.
* Both streams (2048 and 1024 bytes) are **below the 4096-byte mini-stream
  cutoff** the fixture itself writes at `0x38`. [MS-CFB] puts such user streams in
  the mini stream. It works only because `read_stream_by_index`
  (`src/cfb/reader.rs:152-162`) falls back to the regular FAT when
  `self.mini_stream.is_empty()`, and the fixture sets the Root Entry's size to 0.

So the flagship fixture "validates a structural rule" (real `.doc` bytes → IR)
considerably more weakly than it looks: it validates this repo's *tolerances*, not
the on-disk layout Word writes. That said, the parts that matter for *this change*
— FIB `fcStshf`, `STSH`/`Stshif`/`LPStd`/`STD`/`Xstz`, `GrpPrlAndIstd.istd` — are
genuinely spec-shaped, which is why the style-sheet half really is exercised.

**(c) `build_synthetic_styled_doc` and `build_synthetic_outline_doc` are ~90 lines
of duplicated bytes** differing only in one PAPX; `lpstd`/`build_stsh_heading3` are
duplicated a third time in `styles.rs` as `lpstd`/`synthetic_stsh`. Not a
correctness issue, but it triples the surface where a fixture and the decoder can
drift together.

---

## Q6 — Does any name, comment or constant assert a property the code lacks?

1. **`sprmPOutlineLvl` / `sprm_p_outline_lvl_*` / `outline_lvl_explicit`.** The
   name asserts the code reads MS-DOC's outline-level sprm. It reads
   `sprmPDyaLine`. Every test name, doc comment and field name in this cluster is
   false. **[specified]**

2. **`sprmPStyle` / `style_istd` / `pap_props_read_sprm_p_style_override`.** There
   is no `sprmPStyle`; `0x640A` is undefined. **[specified]**

3. **The documented precedence is inverted relative to the spec.**
   `resolve_heading_level`'s doc comment: *"Direct formatting overrides the style
   in Word, so `sprmPOutlineLvl` … settles the question."* [MS-DOC] §2.6.2,
   `sprmPOutLvl`, verbatim: *"This MUST be ignored if the paragraph has an **istd**
   that is greater than or equal to 0x1 and less than or equal to 0x9."*
   **[specified]** — the style wins. (In fairness, POI deliberately disables that
   condition: `ParagraphSprmUncompressor.java:314-318`, "Word seems to set outline
   levels even for paragraph with other styles than Heading 1..9 … See bug 49820".
   So the *behaviour* is defensible; the *comment presenting it as the spec rule*
   is not, and the comment cites no source.)

4. **`sprmPOutlineLvl` level `0` = "explicit body-text marker".** Spec: operand
   `0x0–0x8` is the **zero-based** outline level (so `0` = Heading 1) and `0x9` is
   body text. The change's mapping is shifted by one and its sentinel is at the
   wrong end. **[specified]**

5. **`MAX_OUTLINE_LEVEL`'s doc comment** claims the sprm operand, `StdfBase.sti`
   and a `Heading N` name "all use this range". `sti` is a 12-bit field spanning
   `0x0000–0x0FFE`. **[specified], §2.9.260**

6. **"Every real `.doc` in the POI / Tika / LibreOffice corpora reports
   `fc_stshf == 0` (no style sheet)."** [MS-DOC] §2.9.271: *"Each FIB MUST contain
   a stylesheet."* **[specified]**. If the author's corpus really reported zero
   there, the likely explanation is a Word 6/95 (`wIdent 0xA5DC`) subset — for
   which this FIB layout does not apply at all — or a measurement error, not that
   Word omits style sheets. This claim is load-bearing: it is the stated
   justification for skipping the corpus check that AGENTS.md rule 5 requires, and
   for building the whole thing on synthetic bytes. Since the FIB offsets `0xA2/0xA6`
   *are* right (Q3), the claim should have been re-checked before it was used to
   excuse the corpus.

7. **`styled_heading_inside_table_stays_a_cell` / `..._on_row_terminator_is_not_a_heading`.**
   The names and assertions frame "a styled heading inside a table never becomes a
   Heading" as the correct outcome. It is a fidelity loss: `add_cell_paragraph`
   flattens every in-cell paragraph, so a genuine `Heading 3` in a table cell is
   emitted as plain text. Pinning that with a test makes it look validated.

8. **`ir::MAX_HEADING_DEPTH`'s comment** — "Format readers whose native heading
   range is deeper … clamp to this when they build the IR, so every consumer can
   rely on the bound" — while `src/convert_docx.rs:201` still writes `.min(6)`
   literally and nothing enforces the invariant on `Heading` construction or
   deserialization (`#[serde(default = "default_heading_level")]` only supplies a
   default; a hostile serialized IR can still carry `level: 200`).

9. **Cosmetic:** the new `emit_prose` arm copies the heuristic's `t.bold = true`,
   so a real styled heading renders as `### **Subsection Three**`. The DOCX path
   does not do this. Not asserted anywhere, but it means "the styled heading keeps
   its real level" comes bundled with an unrelated formatting fabrication.

---

## Robustness (untrusted input)

I found **no panic or OOB read** on the change's own paths.

* `Fib::parse`: `data.len() > 0x00A9` before reading `0xA2..0xAA`; `read_u32` is
  itself bounds-checked. ✓
* `parse_style_sheet`: `saturating_add`, `.min(len)`, `start >= len || end <= start`
  → the slice is always in range. ✓
* `parse_stsh`: `cb_stshi < 18` rejects; `pos + cb_stshi > len` rejects; per-entry
  `pos + 2 > len` and `pos + cb_std > len` break. `pos` can exceed `len` after the
  odd-padding `+= 1`, but the next iteration's guard catches it. `cap = cstd.min(4096)`
  bounds the allocation to ~128 KB. Total work is bounded by `lcbStshf`. ✓
* `parse_xstz`: `saturating_mul`, `end > data.len()` → `None`. ✓
* `extract_grpprl`'s new `istd` read: guarded by `p + 3 <= page.len()`, `p ≤ 511`. ✓
* `emit_prose`: `level.min(MAX_HEADING_DEPTH)`. There is **no lower clamp**; today
  no producer can reach `Some(0)` (`resolve_heading_level` only yields `1..=9`), but
  if one ever did, `"#".repeat(0)` renders a heading as bare text with a leading
  space. Latent, unreached. Worth a `.clamp(1, MAX_HEADING_DEPTH)`.
* `Heading` deserialization is unbounded — `ir_render` clamps at render time, so no
  crash, but `MAX_HEADING_DEPTH` is a convention, not an invariant.

**Silently swallowed signals.** `parse_style_sheet`/`parse_stsh` return `Vec::new()`
on every malformed condition, and `document.rs` has no way to distinguish "no style
sheet" from "malformed style sheet" — both degrade to the line heuristic with no
diagnostic. That is consistent with the existing `.doc` code and arguably right for
this parser, but AGENTS.md rule 7 says "fail loudly, never fall back to a silent
plausible-but-wrong result", and the fallback here *is* the plausible-but-wrong
heuristic. A `debug!`/warning would cost nothing.

**Missing:** `fuzz/` is untouched despite a new parser reading attacker-controlled
offsets (AGENTS.md rule 6). `CHANGELOG.md` is untouched.

**Small pre-existing edge the change inherits:** `extract_grpprl` returns
`istd = 0` whenever `cb < 3`. In the `cw == 0` re-read form, `cb == 2` means the
`GrpPrlAndIstd` is exactly the 2-byte `istd` with no `grpprl` — a legitimate
"styled, no direct formatting" PAPX whose `istd` is now discarded. Rare, but it is
the one case where a real `Heading 3` would be dropped by the new code.

---

## Verdict

**right symptom, wrong layer**

The diagnosis is correct and overdue: `.doc` heading detection was a text-shape
guess, and `STSH` + `istd` is the right ground truth. The style-sheet half —
`fib.rs` offsets, `styles.rs`, the `istd` plumbing through `FkpParagraph`, the
`MAX_HEADING_DEPTH` clamp — is spec-accurate where I could check it against
[MS-DOC], well bounds-checked, and genuinely exercised end to end by the flagship
test.

But the change ships a second, co-equal mechanism built on a **misidentified
opcode**: `0x6412` is `sprmPDyaLine` (line spacing), not an outline level, and
`0x640A` does not exist. That mechanism is given *precedence over* the correct one
("the SPRM is authoritative and wins"), so on real documents it will both fabricate
headings from line-spacing values and suppress genuine styled headings — the exact
defect the PR set out to fix, reintroduced from a different direction. Eight tests
and one of the two synthetic `.doc` fixtures exist to confirm the misreading rather
than to test it. And the underlying cause — a text-shape heuristic still running on
every paragraph the style sheet has already classified — is untouched; the change
adds a bypass in front of it instead of retiring it where ground truth now exists.

The remedy is small and local: delete the `0x640A` arm; change `0x6412` to `0x2640`
with a **1-byte** operand; map `0..=8` to level `operand + 1` and `9` to body text;
decide the style-vs-sprm precedence deliberately (spec says style wins for
`istd ∈ 1..=9`, POI deliberately deviates — either is defensible, but cite it); and
gate the `emit_prose` heuristic on `styles.is_empty()`.

### Strongest argument against my own verdict

"Wrong layer" is arguably too harsh, and here is the best case that this is really
**correct but narrow** with one fixable bug:

The style-sheet path is the load-bearing half and it is right. It is also the half
that fires on the overwhelming majority of real documents, because Word encodes
`Heading N` through the paragraph *style* (`istd`), not through direct outline-level
formatting — `sprmPOutLvl` in a PAPX is comparatively rare, and per [MS-DOC] it is
supposed to be *ignored* whenever `istd ∈ 1..=9`, i.e. in exactly the cases this
change cares about. So deleting the two bogus arms outright would lose almost
nothing, and the remaining change would be a clean, correct, well-tested
improvement that lands the architecture (STSH parsing, `MAX_HEADING_DEPTH`, `istd`
on `FkpParagraph`) the right way round for the first time. On that reading the
opcode error is a removable wart on a sound change, not evidence of the fix sitting
at the wrong layer.

I still land on "right symptom, wrong layer" for two reasons the counter-argument
does not answer. First, the bogus arm is not inert: it is given *priority* over the
correct one and its `≤ 9` accept-window sits squarely on real `dyaLine` low bytes,
so it actively degrades output on documents the correct half would have handled.
Second, the heuristic that caused the original defect is still running unguarded
behind the new code, so even with the opcodes fixed the parser keeps guessing where
it now has ground truth — which is what makes the placement, not just the opcode,
the thing to change.
