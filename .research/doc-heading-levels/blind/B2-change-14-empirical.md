# B2 — change-14: empirical audit of the shipped tests

Worktree: `/home/yfedoseev/projects/office_oxide/.claude/worktrees/agent-a15412f1b9130f80a`
Baseline: `5007d9c`. Change applied as local commit `488fd3d` ("applied change-14") so every
probe could be reverted with `git checkout -- .`. The worktree was left in the
applied-change state with a clean `git status`.

Everything under **Measured** is a command I ran with its output pasted. Everything
under **Inferred** is reasoning on top of a measurement and is labelled as such.

---

## 1. Build / lint / test results as measured

### Baseline (before applying the diff)

```
$ cargo test --lib
test result: ok. 504 passed; 0 failed; 0 ignored; 0 measured; 0 filtered out; finished in 0.14s
[exited with code 0]
```

No pre-existing failures.

### After `git apply change-14.diff`

```
$ cargo build
    Finished `dev` profile [unoptimized + debuginfo] target(s) in 18.88s

$ cargo fmt --check
fmt exit: 0                     (no output, exit 0)

$ cargo test --all-targets
running 536 tests
test result: ok. 536 passed; 0 failed; ...          (lib)
running 17 tests   ... ok
running 3 tests    ... ok      (x4 doc integration targets)
running 19 tests   ... ok
running 4 tests    ... ok
running 23 tests   ... ok
running 20 tests   ... ok
running 38 tests   ... ok
running 16 tests   ... ok
running 2 tests    ... ok
[exited with code 0]

$ cargo clippy --all-targets --workspace -- -D warnings
    Finished `dev` profile [unoptimized + debuginfo] target(s) in 43.53s
clippy=0
```

**Everything the change claims about build/lint/test is true as measured.**
Lib tests 504 → 536 = **+32 new tests**, all green; no other target changed count.

---

## 2. Every test the diff adds or modifies

**32 new tests, 4 modified (signature-only).**

### New — `src/convert_doc.rs` (3)
| # | test |
|---|---|
| 1 | `convert_doc::tests::styled_heading_inside_table_stays_a_cell` |
| 2 | `convert_doc::tests::styled_heading_on_row_terminator_is_not_a_heading` |
| 3 | `convert_doc::tests::styled_heading_list_item_stays_a_list_item` |

### New — `src/doc/document.rs` (4)
| # | test |
|---|---|
| 4 | `doc::document::tests::ir_styled_heading_uses_real_level` |
| 5 | `doc::document::tests::ir_deep_outline_level_clamps_to_ir_max_depth` |
| 6 | `doc::document::tests::synthetic_doc_styled_heading_uses_style_sheet_level` |
| 7 | `doc::document::tests::synthetic_doc_outline_sprm_heading` |

### New — `src/doc/papx.rs` (10)
| # | test |
|---|---|
| 8 | `build_paragraphs_resolves_heading_from_style_istd` |
| 9 | `build_paragraphs_prefers_sprm_style_over_papx_istd` |
| 10 | `build_paragraphs_resolves_user_heading_name_case_insensitive` |
| 11 | `build_paragraphs_resolves_heading_from_sprm_outline_lvl` |
| 12 | `build_paragraphs_sprm_outline_lvl_overrides_style` |
| 13 | `build_paragraphs_leaves_body_text_unheaded_with_styles_present` |
| 14 | `build_paragraphs_sprm_style_override_can_remove_a_heading` |
| 15 | `build_paragraphs_sprm_outline_lvl_zero_suppresses_styled_heading` |
| 16 | `build_paragraphs_sprm_outline_lvl_accepts_deepest_level` |
| 17 | `build_paragraphs_resolves_heading_from_parsed_style_sheet` |

### New — `src/doc/sprm.rs` (5)
| # | test |
|---|---|
| 18 | `sprm_p_outline_lvl_sets_heading_level` |
| 19 | `sprm_p_outline_lvl_body_is_none` |
| 20 | `sprm_p_outline_lvl_rejects_invalid_byte` |
| 21 | `pap_props_read_sprm_p_style_override` |
| 22 | `pap_props_style_istd_absent_when_no_sprm` |

### New — `src/doc/styles.rs` (10, new file)
| # | test |
|---|---|
| 23 | `parse_stsh_reads_sti_and_names` |
| 24 | `parse_stsh_rejects_spurious_name_table_layout` |
| 25 | `parse_stsh_rejects_unexpected_cb_std_base_in_file` |
| 26 | `parse_stsh_skips_odd_sized_lpstd_padding` |
| 27 | `heading_level_resolves_builtin_and_user` |
| 28 | `heading_level_uses_primary_name_before_aliases` |
| 29 | `truncated_style_sheet_is_empty` |
| 30 | `parse_style_sheet_reads_stsh_at_fib_offset` |
| 31 | `parse_style_sheet_absent_is_empty` |
| 32 | `parse_style_sheet_out_of_bounds_is_empty` |

### Modified (call-signature updates only — no new assertions)
`doc::papx::tests::build_paragraphs_slices_text_and_flags`,
`build_paragraphs_keeps_astral_alignment`,
`build_paragraphs_strips_field_codes` (all gained `, &[]` for the new `styles`
parameter), and `papx_cw_zero_reread_extracts_trailing_tdef_table`
(`extract_grpprl(...)` → `.grpprl`).

### Production hunks (labels used below)
| id | file | what |
|---|---|---|
| H-IR | `src/ir.rs` | `pub const MAX_HEADING_DEPTH: u8 = 6` |
| H-REND | `src/ir_render.rs` | `min(6)` → `min(MAX_HEADING_DEPTH)` (no behaviour change) |
| H-FIB | `src/doc/fib.rs` | read `fc_stshf`@0x00A2 / `lcb_stshf`@0x00A6 |
| H-MOD | `src/doc/mod.rs` | `pub mod styles;` + `MAX_OUTLINE_LEVEL = 9` |
| H-STY | `src/doc/styles.rs` | new: `parse_style_sheet`, `parse_stsh`, `parse_xstz`, `heading_level_for_istd`, `heading_level_from_name` |
| H-SPRM-A | `src/doc/sprm.rs` | `0x640A` arm → `style_istd` |
| H-SPRM-B | `src/doc/sprm.rs` | `0x6412` arm → `heading_level` + `outline_lvl_explicit` |
| H-PAPX-I | `src/doc/papx.rs` | `FkpParagraph.istd` / `PapxData` / `extract_grpprl` reads istd |
| H-PAPX-R | `src/doc/papx.rs` | `resolve_heading_level` + `build_paragraphs(styles)` |
| H-DOC | `src/doc/document.rs` | call `parse_style_sheet` and pass it down |
| H-CONV-A | `src/convert_doc.rs` | `emit_prose` styled-heading branch |
| H-CONV-B | `src/convert_doc.rs` | `level.min(MAX_HEADING_DEPTH)` clamp |

---

## 3. Per-test revert table

Method: for each production predicate I made the minimal semantic revert **in place**
(so types still compile and the whole suite still runs), ran `cargo test --lib`, and
recorded which tests flipped to FAILED. Restored with `git checkout -- <file>` between
probes. Because the whole suite was run each time, the table is exact: a test not listed
as failing for a given revert **passes with that hunk reverted**.

| test | depends on hunk(s) | result with those hunks reverted | verdict |
|---|---|---|---|
| `sprm_p_outline_lvl_sets_heading_level` | H-SPRM-B (arm body; `outline_lvl_explicit=true`) | FAILED | real |
| `sprm_p_outline_lvl_body_is_none` | H-SPRM-B (arm body; `b>=1` guard; explicit flag) | FAILED | real |
| `sprm_p_outline_lvl_rejects_invalid_byte` | H-SPRM-B range guard **only at `b<=104`** | **PASSES** with arm deleted; PASSES with range widened to `<=99`; FAILED only at `<=104` | partly decorative — see §5 |
| `pap_props_read_sprm_p_style_override` | H-SPRM-A | FAILED | real |
| `pap_props_style_istd_absent_when_no_sprm` | — | **PASSES** with H-SPRM-A deleted | decorative (negative test) |
| `build_paragraphs_resolves_heading_from_sprm_outline_lvl` | H-SPRM-B + H-PAPX-R | FAILED | real |
| `build_paragraphs_sprm_outline_lvl_overrides_style` | H-SPRM-B + H-PAPX-R precedence | FAILED (both on arm-delete and on precedence inversion) | real |
| `build_paragraphs_sprm_outline_lvl_zero_suppresses_styled_heading` | `outline_lvl_explicit`, `b>=1`, precedence | FAILED (4 separate reverts) | real |
| `build_paragraphs_sprm_outline_lvl_accepts_deepest_level` | H-SPRM-B | FAILED | real |
| `build_paragraphs_resolves_heading_from_style_istd` | H-STY + H-PAPX-R style fallback | FAILED | real |
| `build_paragraphs_prefers_sprm_style_over_papx_istd` | H-SPRM-A + `style_istd.unwrap_or(fp.istd)` | FAILED | real |
| `build_paragraphs_sprm_style_override_can_remove_a_heading` | H-SPRM-A + `style_istd.unwrap_or(fp.istd)` | FAILED | real |
| `build_paragraphs_resolves_user_heading_name_case_insensitive` | H-STY name path | FAILED (style-fallback revert) — but **PASSES** with lowercasing removed | see §5 |
| `build_paragraphs_leaves_body_text_unheaded_with_styles_present` | H-STY `sti != 0` | FAILED only when the sti range is widened to include 0 | thin but real |
| `build_paragraphs_resolves_heading_from_parsed_style_sheet` | H-STY + H-PAPX-R + name offset | FAILED | real |
| `parse_stsh_reads_sti_and_names` | H-STY name offset `std[cb_std_base..]` | FAILED | real |
| `parse_stsh_rejects_spurious_name_table_layout` | — | **PASSES under every revert I made.** Instrumented: `parse_stsh` returns `[]`, so `styles.iter().all(...)` is **vacuously true** (`AUDIT len=0 styles=[]`) | **fully decorative** |
| `parse_stsh_rejects_unexpected_cb_std_base_in_file` | H-STY `cb_std_base != 0x000A && != 0x0012` | FAILED | real |
| `parse_stsh_skips_odd_sized_lpstd_padding` | H-STY odd-padding skip; also name offset | FAILED | real |
| `heading_level_resolves_builtin_and_user` | H-STY sti-range lower bound + name offset | FAILED on `0..=9` widening and name offset; **PASSES** on `1..=200` widening | partly real |
| `heading_level_uses_primary_name_before_aliases` | H-STY lowercase + comma split | FAILED on both | real (and it is the *only* test that covers capital-`H` names — see §5) |
| `truncated_style_sheet_is_empty` | H-STY `pos + cb_stshi > data.len()` | FAILED (panics: index OOB) | real |
| `parse_style_sheet_reads_stsh_at_fib_offset` | H-STY + name offset | FAILED | real |
| `parse_style_sheet_absent_is_empty` | — | **PASSES** with the `fc_stshf==0 \|\| lcb_stshf==0` guard deleted (the `end<=start` clause satisfies it) | decorative |
| `parse_style_sheet_out_of_bounds_is_empty` | — | **PASSES** with the `start >= table_stream.len()` clause deleted | decorative (the `end<=start` clause is what fires) |
| `ir_styled_heading_uses_real_level` | H-CONV-A | FAILED | real |
| `ir_deep_outline_level_clamps_to_ir_max_depth` | H-CONV-B (and H-CONV-A) | FAILED on clamp removal **and** on `min(7)` | real, boundary pinned |
| `synthetic_doc_styled_heading_uses_style_sheet_level` | H-FIB, H-DOC, H-PAPX-I, H-PAPX-R, H-STY, H-CONV-A | FAILED under *each* of those six reverts individually | real — the strongest test in the change, and the **only** test covering H-FIB, H-DOC and H-PAPX-I |
| `synthetic_doc_outline_sprm_heading` | H-SPRM-B, H-PAPX-R, H-CONV-A | FAILED | real |
| `styled_heading_inside_table_stays_a_cell` | — | **PASSES** with H-CONV-A deleted entirely | **decorative** |
| `styled_heading_on_row_terminator_is_not_a_heading` | — | **PASSES** with H-CONV-A deleted entirely | **decorative** |
| `styled_heading_list_item_stays_a_list_item` | — | **PASSES** with H-CONV-A deleted entirely | **decorative** |
| the 4 modified papx tests | — | signature-only; assert nothing new | n/a |

Key reverts, verbatim:

```
# H-CONV-A: styled-heading branch neutered (`heading_level.filter(|_| false)`)
test doc::document::tests::ir_deep_outline_level_clamps_to_ir_max_depth ... FAILED
test doc::document::tests::ir_styled_heading_uses_real_level ... FAILED
test doc::document::tests::synthetic_doc_outline_sprm_heading ... FAILED
test doc::document::tests::synthetic_doc_styled_heading_uses_style_sheet_level ... FAILED
test result: FAILED. 532 passed; 4 failed; ...
```

Note what is *absent* from that list: all three `convert_doc::tests::styled_heading_*`
tests. **Inferred:** they assert only that `walk_paragraphs` routes table / row-terminator /
list paragraphs before ever reaching `emit_prose`. That routing is pre-existing baseline
behaviour; the new `heading_level` field simply never reaches those branches. The
load-bearing assertion in each is the negative `!els.iter().any(|e| matches!(e, Element::Heading(_)))`
and it is satisfied by code the change did not touch. The positive assertions
(`any(Table)`, `any(List)`) are equally pre-existing. Nothing in these three tests
constrains the new production code.

```
# H-PAPX-I: extract_grpprl no longer reads istd from the page
test doc::document::tests::synthetic_doc_styled_heading_uses_style_sheet_level ... FAILED
test result: FAILED. 535 passed; 1 failed; ...

# H-FIB: fcStshf/lcbStshf read from 0x00A6/0x00AA instead of 0x00A2/0x00A6
test doc::document::tests::synthetic_doc_styled_heading_uses_style_sheet_level ... FAILED
test result: FAILED. 535 passed; 1 failed; ...

# H-DOC: document.rs passes an empty style list
test doc::document::tests::synthetic_doc_styled_heading_uses_style_sheet_level ... FAILED
test result: FAILED. 535 passed; 1 failed; ...
```

Three separate production hunks each have **exactly one** covering test, and it is
the same test. **Inferred:** if `synthetic_doc_styled_heading_uses_style_sheet_level`
is ever deleted or weakened, H-FIB, H-DOC and H-PAPX-I become entirely uncovered.

---

## 4. Fixture decoration analysis

I checked the hand-built binary fixtures against the code that consumes them, and
then against **[MS-DOC]** on learn.microsoft.com (fetched, not recalled).

### 4a. What the spec actually says — and the decisive finding

Fetched `[MS-DOC] 2.6.2 Paragraph Properties`
(<https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/484822ee-a9d9-4af4-8423-29fda67a6a58>).
Verbatim rows:

> | sprmPDyaLine<br>(0x6412) | 0x12 | An **LSPD** value that specifies the spacing between lines in this paragraph. By default, paragraphs use single spacing. |

> | sprmPOutLvl<br>(0x2640) | 0x40 | An unsigned 8-bit integer value that specifies the outline level of the paragraph. This value MUST be one of the following. **0x0 - 0x8** The value is the zero-based outline level that this paragraph is in. **0x9** The paragraph at any outline level; instead, the paragraph is body text. This MUST be ignored if the paragraph has an **istd** that is greater than or equal to 0x1 and less than or equal to 0x9. |

> | sprmPIstd<br>(0x4600) | 0x00 | An unsigned integer that specifies the **istd** of a paragraph style to apply. … An **istd** value in the range of 1 to 9, inclusive, also specifies the outline level of the paragraph …, where the new outline level is equal to the value of the **istd** minus 1. |

Consequences, **measured against the spec text above**:

1. **`0x6412` is `sprmPDyaLine` (line spacing), not an outline level.** The change
   names it `sprmPOutlineLvl` and decodes its first operand byte as a heading level.
2. **`0x640A` is not a defined PAP sprm at all.** `ispmd` 0x0A is `sprmPIlvl` (0x260A);
   the style sprm is `sprmPIstd` (0x4600) with a **2-byte** operand. The change calls
   0x640A "sprmPStyle" and reads a 4-byte operand.
3. Even taking the intended sprm, the **polarity is inverted**: the real `sprmPOutLvl`
   uses `0x0–0x8` for outline levels and `0x9` for *body text*. The change treats `0` as
   body text and `1–9` as Heading 1–9. `MAX_OUTLINE_LEVEL = 9` is therefore the
   body-text sentinel in the spec, not the deepest heading.
4. The style→level mapping *is* spec-supported: `sprmPIstd` says istd 1..9 ⇒ outline
   level istd−1, and STSH's fixed-index table (fetched §2.9.271) maps istd 0–9 → sti 0–9.
   So `sti ∈ 1..=9 ⇒ Heading 1..=9` is correct.

**The fixtures and the decoder agree with each other and disagree with the spec.**
Every grpprl fixture in the change is hand-written as `[0x12, 0x64, level, 0, 0, 0]` or
`[0x0A, 0x64, istd, 0, 0, 0]` — i.e. the fixture encodes exactly what the decoder is
looking for. No test would fail if both opcodes were changed together.

**Measured on a real corpus** (see §7 for the harness): `0x6412` occurs **758 times**
in 246 real `.doc` files, with these first-operand-byte values:

```
count  byte
  393   240   (dyaLine = 240 twips  → single spacing)
  208    56
   75    16
   67   104   (dyaLine = 360 twips  → 1.5-line spacing)
    7    14
    4   224
    3   112
    1    20
```

`0x640A` occurs **0** times (it is not a real opcode).

Note the 67 hits at byte **104 = 0x68** — the exact value
`sprm_p_outline_lvl_rejects_invalid_byte` calls "a coincidental byte pattern".
It is not coincidental: it is the low byte of 1.5-line spacing. The test's own
comment misdiagnoses its fixture.

**Inferred:** in this corpus no `dyaLine` low byte lands in `0..=9`, so the misdecoding
is currently latent. But any document with an exactly-specified line spacing congruent
to 0..9 mod 256 (256, 512, 768, 1024, 1280 twips …) would have its low byte read as an
outline level: low byte 0 ⇒ `outline_lvl_explicit = true` and the paragraph's *real*
heading style is suppressed; low bytes 1–9 ⇒ a fabricated Heading 1–9.

### 4b. Bytes written by the fixtures that the code never reads

I wrote perturbation probes (temporary `audit_probes` module in `src/doc/styles.rs`,
since removed) that rebuild the change's own `lpstd`/`stsh` helpers with individual
fields flipped and assert the parse result is unchanged. All passed:

```
test doc::styles::audit_probes::audit_stshif_tail_is_never_read ... ok
test doc::styles::audit_probes::audit_stdfbase_tail_is_never_read ... ok
test doc::styles::audit_probes::audit_xstz_chterm_is_never_read ... ok
test doc::styles::audit_probes::audit_sti_upper_bits_are_masked_but_untested ... ok
```

| fixture bytes | consumed by the code? | note |
|---|---|---|
| `Stshif` bytes 4..18 — `fStdStylenamesWritten`, `stiMaxWhenSaved`, `istdMaxFixedWhenSaved`, `nVerBuiltInNamesWhenSaved`, `ftcAsci/FE/Other` | **never read** (flipping all 14 to `0xFF` changes nothing) | decorative. Note `build_stsh_heading3()` / `stsh_with_heading_3()` leave them all zero, so `istdMaxFixedWhenSaved = 0`, which §2.9.274 says **MUST be 0x000F**. `synthetic_stsh()` sets it correctly — but nothing reads it either way. |
| `StdfBase` bytes 2..10 — `stk`, `istdBase`, `cupx`, `istdNext`, `bchUpe`, `grfstd` | **never read** | decorative. `stk = 0` violates §2.9.260 ("MUST be one of 1..4"); `bchUpe` must equal `cbStd` and is 0. |
| `Xstz.chTerm` (the 2-byte null) | **never read** — `parse_xstz` computes `end = 2 + cch*2` and returns | decorative |
| `StdfBase` first u16 **upper 4 bits** (`fScratch`/`fInvalHeight`/`fHasUpe`/`fMassCopy`) | read, then masked with `& 0x0FFF` | **the mask is never exercised**: every fixture writes `sti.to_le_bytes()` with the top bits zero. Removing `& 0x0FFF` leaves the suite green (see §5). A real `StdfBase` for a paragraph style has `stk = 1` in bits 12..15, so the mask matters in production and is untested. |
| `grLPUpxSw` | **omitted entirely** from every synthetic `STD` | §2.9.258 requires it; the fixture `cbStd` covers only `stdf` + `xstzName`. Invisible to this parser. |
| grpprl operand bytes 2..4 (`0x00,0x00,0x00` after the level/istd) | **never read** — the arms use `operand.first()` / `operand[0..2]` | decorative |
| `build_synthetic_outline_doc`'s 7-byte grpprl `[0x12,0x64,0x05,0,0,0,0]` "+ 1 pad byte" | only byte 2 read | decorative |
| `parse_stsh_rejects_spurious_name_table_layout`'s whole fixture | `parse_stsh` returns `[]` before reading any of it (`cb_std` decodes to 1025 from the phantom STTB header and the loop breaks) | the assertion is vacuous over an empty vector |

### 4c. Things the fixtures get right, confirmed against the spec

* `LPStd` = `cbStd(u16)` + `STD`, entries on even-byte boundaries with `cbStd` excluding
  the pad — §2.9.135, quoted: *"LPStd structures are stored on even-byte boundaries, but
  this length MUST NOT include this padding."* The `parse_stsh_skips_odd_sized_lpstd_padding`
  test is genuinely spec-derived and load-bearing.
* `Stshif.cbSTDBaseInFile` MUST be `0x000A` or `0x0012` — §2.9.274, quoted. Guard and test
  both correct.
* `STD` = `stdf` + `xstzName` + `grLPUpxSw`, name may be `"primary,alias,alias"` — §2.9.258.
  The comma-split is spec-derived.
* `StdfBase.sti` is 12 bits — §2.9.260. Mask correct.
* `Xstz` = `Xst` (`cch` + code units) + 2-byte `chTerm` — §2.9.354.
* PAPX header layout `[cw][istd:2][grpprl]`: `GrpPrlAndIstd` §2.9.114 has `istd` as the
  first 2 bytes. Correct.
* FIB offsets: `FibRgFcLcb97` (§2.5.5 page fetched) places `fcStshf` at index 1, i.e.
  `0x9A + 8 = 0xA2`, `lcbStshf` at `0xA6`. **Correct.**

---

## 5. Mutation probes — caught vs not caught

Each mutation was applied in isolation, the **whole** `cargo test --lib` suite run, and
then reverted.

### Caught

| mutation | tests that failed |
|---|---|
| delete the `0x6412` arm body | 7 |
| never set `outline_lvl_explicit` | 7 |
| drop the `b >= 1` guard (level 0 becomes `Some(0)`) | 2 |
| widen outline range to `b <= 104` | 1 |
| delete the `0x640A` arm body | 3 |
| `props.style_istd.unwrap_or(fp.istd)` → `fp.istd` | 2 |
| `resolve_heading_level` style fallback → `None` | 5 |
| invert precedence (style wins over SPRM) | 2 |
| `extract_grpprl` stops reading `istd` | 1 |
| FIB offsets shifted to 0x00A6/0x00AA | 1 |
| `document.rs` passes an empty style list | 1 |
| delete the `MAX_HEADING_DEPTH` clamp | 1 |
| clamp widened to `min(MAX_HEADING_DEPTH + 1)` | 1 |
| delete the styled-heading branch in `emit_prose` | 4 |
| delete the `cbSTDBaseInFile ∈ {0x0A, 0x12}` guard | 1 |
| delete the odd-`cbStd` padding skip | 1 |
| sti range `1..=9` → `0..=9` | 5 |
| remove `to_ascii_lowercase` | 1 |
| remove the primary-name comma split | 1 |
| name read at `std[2..]` instead of `std[cb_std_base..]` | 5 |
| delete `pos + cb_stshi > data.len()` | 1 (panics) |

### **Not caught** — coverage gaps

| mutation | full-suite result |
|---|---|
| outline range `b <= MAX_OUTLINE_LEVEL` → `b <= 10` (widen by one) | `536 passed; 0 failed` |
| … → `b <= 99` | `536 passed; 0 failed` |
| **built-in sti heading range `1..=9` → `1..=10`** | `536 passed; 0 failed` |
| **… → `1..=200`** | `536 passed; 0 failed` |
| **name-derived level range `1..=9` → `0..=9`** | `536 passed; 0 failed` |
| **… → `1..=49`** | `536 passed; 0 failed` |
| **drop the `& 0x0FFF` mask on `sti`** | `536 passed; 0 failed` |
| drop `cap = cstd.min(4096)` (unbounded `with_capacity`) | `536 passed; 0 failed` |
| drop `start >= table_stream.len()` bound | `536 passed; 0 failed` |
| drop the `fc_stshf == 0 \|\| lcb_stshf == 0` early return | `536 passed; 0 failed` |
| `cb_stshi < 18` → `cb_stshi < 2` (removes a **panic guard**) | `536 passed; 0 failed` |
| **`t.bold = true` → `false` on styled headings** (and on heuristic headings) | `536 passed; 0 failed` |

Notes on the sharpest gaps:

* **Upper bounds are essentially unconstrained.** `sti ∈ 1..=200` is green, so nothing
  stops the code from calling `sti = 65 / 105 / 107` (the fixed-index styles the change's
  own `synthetic_stsh()` *does* build at istd 10/11/12) a heading. Those fixture entries
  exist purely as decoration — no assertion ever asks whether
  `heading_level_for_istd(&styles, 10)` is `None`.
* The outline-level bound is pinned only at 104 — anything from 10 to 103 is accepted
  silently. **Inferred:** the "reject invalid byte" test picked one value rather than the
  boundary; asserting `b = 10` is rejected would pin the actual contract.
* `cb_stshi < 18` is a real panic guard: with it relaxed, `parse_stsh(&[2,0,0,0])`
  indexes `data[4]` out of bounds. The suite does not notice.
* **"Case-insensitive" is mis-tested.** Removing `to_ascii_lowercase()` leaves
  `build_paragraphs_resolves_user_heading_name_case_insensitive` and
  `heading_level_resolves_builtin_and_user` **green** — both use the already-lowercase
  name `"heading N"`, which is the matcher's own normalized form. The only test that
  exercises a capital-`H` name is `heading_level_uses_primary_name_before_aliases`
  ("Heading 4,Title 4"), whose stated purpose is aliases, not case.

---

## 6. Robustness probes (no panics / hangs / unbounded allocation)

Temporary probe modules added to `src/doc/styles.rs` and `src/doc/document.rs`
(since removed). All run in **debug** (overflow checks on).

```
test doc::styles::audit_probes::audit_no_panic_on_hostile_stsh ... ok
test doc::styles::audit_probes::audit_no_panic_parse_style_sheet_wrapper ... ok
test doc::document::audit_doc_probes::audit_truncated_synthetic_doc_never_panics ... ok
test doc::document::audit_doc_probes::audit_corrupted_synthetic_doc_never_panics ... ok
test doc::document::audit_doc_probes::audit_self_inconsistent_stshf_never_panics ... ok
```

Inputs exercised:

* `parse_stsh` on every all-zero and all-`0xFF` buffer of length 0..64; on `cstd = 65535`
  with a truncated `rglpstd`; on `cbStd = 0xFFFF` overrunning the buffer; on `cch = 0xFFFF`
  overrunning the `STD`; on **every prefix** of a well-formed sheet; and on **every
  single-byte corruption** (to `0x00` and `0xFF`) of a well-formed sheet.
* `parse_style_sheet` with `(fc, lcb)` ∈ `{(0,0), (0,MAX), (MAX,MAX), (MAX,1), (1,MAX),
  (len,10), (len-1,MAX)}` against both a real stream and an empty one.
* `DocDocument::from_reader` on every 7-byte-stride prefix of the synthetic `.doc`, on the
  empty input, and on **every** single-byte corruption (3 values × ~4600 offsets ≈ 14k parses)
  past the CFB header.
* `DocDocument::from_reader` with `fcStshf`/`lcbStshf` overwritten to
  `(0xFFFFFFFF,0xFFFFFFFF)`, `(1,0xFFFFFFFF)`, `(0x120,0xFFFFFFFF)`, `(4,0xFFFFFFF0)`, …

**No panic, hang or unbounded allocation in any code the change adds.**

**One panic found, in pre-existing code the change does not touch.** Corrupting a byte in
the CFB header region (offsets 0..512) reaches:

```
thread '...' panicked at src/cfb/header.rs:76:27:
attempt to shift left with overflow
```

`src/cfb/header.rs:75-76`: `let sector_power = u16::from_le_bytes([buf[0x1E], buf[0x1F]]);
let sector_size = 1usize << sector_power;` — an attacker-controlled `sector_power` ≥ 64
panics in debug / wraps in release. This is a baseline AGENTS.md rule-6 violation,
**not introduced by this change**, but reported because it is what the new synthetic
`.doc` fixture happens to surface first.

**Fuzz coverage:** `fuzz/fuzz_targets/fuzz_parse.rs` does call
`Document::from_reader(Cursor::new(data), DocumentFormat::Doc)`, so the new
`parse_style_sheet` path is *reachable* from the existing target and the change did not
need a new one. **Inferred:** reaching it requires a byte string that is a valid CFB with a
valid FIB and a `PlcfBtePapx`; there is no `fuzz/corpus/` directory in the repo, so
un-seeded fuzzing is very unlikely to get there. The change adds no seed and does not
extend the target.

---

## 7. Corpus reach — the headline measurement

The repo ships no `.doc` fixtures, but a real corpus exists at
`/home/yfedoseev/projects/office_oxide_tests/doc` (**246 files**, POI/Tika-derived).
I instrumented `DocDocument::from_reader` with an `AUDIT_TRACE`-gated `eprintln!` and
walked it (probe since removed; corpus read-only, untouched).

```
$ AUDIT_TRACE=1 AUDIT_DOC_DIR=.../doc cargo test --lib audit_corpus_walk -- --ignored --nocapture
AUDITTOTAL files=246 parsed_ok=232
```

200 of those reach `build_paragraphs` (the rest have no `PlcfBtePapx`). Aggregated:

| metric (change as shipped) | count |
|---|---|
| documents reaching `build_paragraphs` | 200 |
| `fc_stshf == 0` | **199 / 200** |
| `lcb_stshf == 0` | **0 / 200** |
| documents where `parse_style_sheet` returned ≥1 style | **1** |
| documents with ≥1 built-in heading style | **1** |
| documents with ≥1 styled-heading paragraph | **0** |
| total styled-heading paragraphs | **0** |
| documents where the `0x6412` arm set a level | **0** |
| documents where the `0x640A` arm fired | **0** |

**Zero real documents reach the new heading path.** The change's own comment asserts
this is because *"Every real `.doc` … reports `fc_stshf == 0` (no style sheet)"*.
That reading is wrong, and the measurement shows why: **`lcb_stshf` is non-zero in all
200 files.** `fcStshf == 0` means the STSH starts at **offset 0** of the Table stream,
which is where Word normally puts it.

`[MS-DOC] 2.5.5 FibRgFcLcb97`
(<https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/0c9df81f-98d0-454e-ad84-b612cd05b1a4>),
verbatim:

> **fcStshf (4 bytes):** An unsigned integer that specifies an offset in the Table Stream. An **STSH** that specifies the style sheet for this document begins at this offset.
>
> **lcbStshf (4 bytes):** An unsigned integer that specifies the size, in bytes, of the **STSH** that begins at offset **fcStshf** in the Table Stream. **This MUST be a nonzero value.**

Unlike every neighbouring Fc/Lcb pair on that page (e.g. *"If lcbPlcffndRef is zero,
fcPlcffndRef is undefined and MUST be ignored"*), `fcStshf` carries **no** "zero means
absent" clause — because §2.9.271 states *"Each FIB MUST contain a stylesheet."*
So `fib.fc_stshf == 0` is a valid offset, and `parse_style_sheet`'s
`if fib.fc_stshf == 0 || fib.lcb_stshf == 0 { return Vec::new(); }` discards it.

**Measured proof.** I deleted only the `fib.fc_stshf == 0` clause (leaving `lcb_stshf == 0`
as the absence signal) and re-ran the same corpus walk:

| metric | as shipped | with `fc_stshf==0` treated as offset 0 |
|---|---|---|
| documents with ≥1 parsed style | **1 / 200** | **200 / 200** |
| documents with ≥1 built-in heading style | **1** | **73** |
| documents with ≥1 styled-heading paragraph | **0** | **47** |
| total styled-heading paragraphs | **0** | **270** |

Example rows from the "after" run:

```
AUDITROW fc_stshf=0 lcb_stshf=1354 styles=39 heading_styles=9 paras=550 heading_paras=31 sprm6412=0 sprm640A=0
AUDITROW fc_stshf=0 lcb_stshf=6180 styles=97 heading_styles=9 paras=17  heading_paras=3  sprm6412=0 sprm640A=0
AUDITROW fc_stshf=0 lcb_stshf=1906 styles=18 heading_styles=3 paras=30  heading_paras=2  sprm6412=0 sprm640A=0
```

So the feature is not merely "carried by synthetic tests" — as shipped it is **inert on
every real document in a 246-file corpus**, and 270 heading paragraphs that it was
written to recover are being missed by a single spec-misreading in a guard clause. The
one-line comment that rationalised the synthetic-only test strategy is itself the bug,
and no test in the change can see it, because every fixture that exercises the style path
places the STSH at a **non-zero** offset (64, 0x120, …) — the exact case that never occurs
in practice.

---

## 8. What these tests actually prove, in one paragraph

They prove, genuinely and with per-test evidence, that the *internal wiring* the change
introduces is connected end to end: the STSH byte layout is decoded to the spec's
`LPStd`/`STD`/`Xstz` shapes (§2.9.135 / §2.9.258 / §2.9.354), a `cbSTDBaseInFile` outside
`{0x0A, 0x12}` is rejected, odd-sized `LPStd` padding does not desynchronise the array,
built-in `sti` and user `"Heading N"` names both reach a level, the SPRM path beats the
style path and an explicit level-0 marker suppresses a styled heading, the PAPX header
`istd` flows from real FKP bytes through `build_paragraphs` into the IR, and the
9→6 depth clamp fires at exactly the right boundary — twenty-three of the thirty-two new
tests fail when the hunk they name is reverted, and one of them
(`synthetic_doc_styled_heading_uses_style_sheet_level`) is single-handedly the only cover
for three separate production hunks. What they do **not** prove is anything about the
outside world: nine tests are decorative (the three `convert_doc` routing tests pass with
the entire new `emit_prose` branch deleted; `parse_stsh_rejects_spurious_name_table_layout`
asserts `all()` over a vector that is measurably empty; three guard-clause tests are
satisfied by a different clause; two negative tests pass with their arm deleted), every
upper bound in the change is unpinned (`sti ∈ 1..=200` and a name-derived level of 49 both
leave the suite green, as does dropping the `& 0x0FFF` mask that every real `StdfBase`
needs), the two SPRM opcodes are simply wrong against `[MS-DOC] 2.6.2` — `0x6412` is
`sprmPDyaLine` (line spacing, 758 occurrences in the real corpus, 67 of them at the very
byte the tests call "coincidental") and `0x640A` is not a paragraph sprm at all, the real
pair being `sprmPOutLvl (0x2640)` with inverted polarity and `sprmPIstd (0x4600)` — and
the fixtures cannot detect any of that because they were written from the same mental
model as the decoder. The measurement that settles it: across 246 real `.doc` files the
new code path produces **zero** styled headings, because `fcStshf == 0` (true for 199 of
200 files, all with non-zero `lcbStshf`) means "the style sheet starts at offset 0", not
"there is no style sheet" — remove that one clause and the same corpus yields 270 styled
headings in 47 documents. The suite is thorough about the code it was written alongside
and silent about the format it is supposed to read.
