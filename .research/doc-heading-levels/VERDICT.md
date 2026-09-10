# PR #127 — verdict after blind review

Method: `claude_code_tricks/blind-agent-pr-review.md`.
Baseline `5007d9c` (main), held for the whole run. `.research/` was moved out of
the tree for the duration so no round could read prior verdicts.

| round | what it saw | never told |
|---|---|---|
| A1 | neutral symptom: "`.doc` headings come out as `#`/`##` regardless of level, and most aren't detected at all" | that a change exists |
| A2 | neutral symptoms: "body-text paragraphs render as headings; some headings get >6 `#`; `to_markdown()` ≠ `to_ir().to_markdown()`" | that a change exists |
| B1 | `change-14.diff`, no title/description/issue | what it was for |
| B2 | `change-14.diff`, own worktree, empirical | what it was for |

Round C (this document) = me, comparing, verifying every load-bearing citation
against the primary source before recording it. `round1-contaminated.md` is my
own review with the diff in hand, written before any blind round reported, and
carries two corrections in place.

Verdict: **request changes** — the feature does not work on real documents, and
the two opcodes it is built on are the wrong opcodes.

## Does the research reject the approach?

**No — the approach is right and is what every reference reader does.** A1,
reasoning forward from the symptom with no knowledge a change existed,
independently proposed the same design: resolve the real outline level from the
paragraph's `istd` and from the outline SPRM, keep the heuristic only as a
fallback. LibreOffice, Antiword, wvWare and POI all key on the same data. The
contributor diagnosed the defect correctly and picked the right layer to fix it.

Every blocker below is a fixable defect inside a sound design.

---

## Blockers

### B1 — the feature is inert on every real `.doc`

`parse_style_sheet` (`src/doc/styles.rs`) opens with:

```rust
if fib.fc_stshf == 0 || fib.lcb_stshf == 0 { return Vec::new(); }
```

`fcStshf` is an *offset* into the Table stream, and the style sheet is normally
the first thing written there — so `0` is the overwhelmingly common **valid**
value, not a sentinel. [MS-DOC] FibRgFcLcb97 puts the "MUST be nonzero"
requirement on the length, not the offset: `lcbStshf` "specifies the size, in
bytes, of the STSH that begins at offset `fcStshf` in the Table Stream, and this
MUST be a nonzero value." Nothing says `fcStshf` may not be 0. STSH adds "Each
FIB MUST contain a stylesheet."

**Proved three independent ways (round C, by me — not taken from B2).**

*1. The corpus is found, not generated.* `office_oxide_tests/` is its own git repo
(`bd84f08`), sourced per its `SOURCES.md` from LibreOffice QA data, Apache POI
`test-data/document/`, Tika and Pandoc. All 246 files are dated 2026-04-03,
months before this review. `file(1)` reports genuine OLE property sets —
"Name of Creating Application: Microsoft Word 12.1.0 … Last Saved By: Jelmer
Kuperus", "Microsoft Word 9.0 … Author: brian davis … Last Printed: 2002".

*2. A CFB/FIB parser I wrote from scratch* (not office_oxide's, not B2's), over
the 199 files it could parse: **191 have `fcStshf == 0`, and all 191 have
`lcbStshf != 0`.** Following the offset into the correct Table stream
(`fWhichTblStm`), **163 of 184 carry a spec-valid STSH at offset 0** —
`cbSTDBaseInFile` exactly `0x000A` (40) or `0x0012` (123), the only two values
Stshif permits. So `0` is an offset, not an absence marker.

*3. An independent third-party reader agrees the data is there.* `antiword`
(which derives headings from `istd ∈ 1..=9` and never reads the outline SPRM)
reports headings in **108 of the 246** files, with real nesting — `53379.doc`
yields `<sect1>` through `<sect4>`.

*4. The end-to-end counterfactual, run by me on this corpus:*

| build | docs with headings | headings emitted |
|---|---|---|
| baseline `main` (no PR) | 128 | **393** |
| PR #127 as submitted | 128 | **393** |
| PR #127 + `fc_stshf == 0` removed from the guard | 131 | **499** |

Baseline and the PR are **identical**. The change adds **zero** headings across
246 real Word documents. Deleting one clause from one guard adds 106 headings in
3 more documents. (B2 measured the same shape with a different instrument and
different skip rules: 1/200 → 200/200 style sheets, 0 → 270 headings. The counts
differ because we counted different things; the conclusion is the same.)

So the PR's own observation — "every one of the 160 poi files reports
`fcStshf == 0`, so no real file reaches the new code" — is not a property of the
corpus. It is this bug, reported as if it were evidence, and it is what justified
shipping the feature on synthetic tests alone. No fixture can catch it: every
hand-built fixture places the STSH at a non-zero offset.

**Fix:** guard on `lcb_stshf` only.

### B2 — both opcodes are wrong

Verified verbatim against [MS-DOC] §2.6.2 Paragraph Properties, and
independently by B1 (LibreOffice `sprmids.hxx`) and B2 (POI source):

| the change says | what that opcode actually is | the real sprm |
|---|---|---|
| `0x6412` = `sprmPOutlineLvl` | **`sprmPDyaLine`** — "An LSPD value that specifies the spacing between lines in this paragraph" | **`sprmPOutLvl` (`0x2640`)**, 1-byte operand |
| `0x640A` = `sprmPStyle` | **no such paragraph sprm** (ispmd `0x0A` is `sprmPIlvl`, `0x260A`) | **`sprmPIstd` (`0x4600`)**, 2-byte operand |

POI's `ParagraphSprmUncompressor.java` confirms both from the other side:
`case 0x12` is `newPAP.setLspd(new LineSpacingDescriptor(...))`; `case 0x40` is
`newPAP.setLvl(...)`.

So `src/doc/sprm.rs` decodes **line spacing** and reads its low byte as an
outline level. `LSPD` is `dyaLine` (16-bit) then `fMultLinespace` (16-bit),
little-endian, so the byte read is `dyaLine & 0xFF`.

B2 measured `0x6412` occurring **758×** across the corpus — all of it line
spacing, 67 of those with operand byte `0x68`. That is the exact value the new
test `sprm_p_outline_lvl_rejects_invalid_byte` calls "a coincidental byte
pattern": it is 1.5-line spacing, and the test asserts the decoder rejects real
data it should be reading as line spacing.

Reachable false positives (arithmetic over a specified structure; not yet seen
in the corpus, which is luck rather than protection):

- `Exactly 12.5 pt` → 250 twips, encoded `0x10000-250` = `0xFF06` → low byte
  **6** → the body paragraph becomes an **H6**.
- `At least 13 pt` → 260 twips = `0x0104` → low byte **4** → **H4**. 26 pt → **H8**.
- Low byte **0** (e.g. 25.6 pt = `0x0200`) sets `outline_lvl_explicit`, which
  *suppresses* a genuinely styled heading. The failure runs both ways.

### B3 — the level value space is off by one with an inverted sentinel

[MS-DOC] `sprmPOutLvl`, verbatim:

> **0x0 - 0x8** The value is the zero-based outline level that this paragraph is in.
> **0x9** The paragraph at any outline level; instead, the paragraph is body text.
> … By default, paragraphs are body text, and are therefore not in any outline level.

The change implements "`0` = body text, `1`–`9` = Heading 1–9": `0` is Heading 1,
not body text, and `9` is body text, not Heading 9. This is the same inverted
model that `src/ir.rs:743` states in a doc comment (see N1) — the contributor
implemented what the crate's own public type says.

### B4 — nine of the thirty-two new tests pass with the fix reverted

Measured per test by B2, not inferred. Gates themselves are genuinely green
(fmt, clippy, 536 tests) and 23 of 32 do fail on baseline as claimed. The nine:

- all three `convert_doc::tests::styled_heading_*` tests pass with the entire new
  `emit_prose` styled-heading branch deleted — they assert pre-existing
  `walk_paragraphs` routing, not the new code;
- `parse_stsh_rejects_spurious_name_table_layout` asserts
  `styles.iter().all(|s| s.name.is_empty())` over a vector B2 instrumented and
  found **empty** — vacuously true under any implementation, including one that
  does nothing;
- `parse_style_sheet_absent_is_empty` and `parse_style_sheet_out_of_bounds_is_empty`
  are satisfied by a different clause than the one they name;
- two negative SPRM tests pass with their arm deleted.

Three production hunks (FIB offsets, `document.rs` wiring, PAPX `istd`) have
exactly one covering test between them — the same one.

Mutations the suite does not catch: widening `sti` to `1..=200`; widening the
name-derived level to `1..=49`; **dropping the `& 0x0FFF` `sti` mask**, which
every real `StdfBase` needs because `stk = 1` sets bit 12; relaxing the
`cb_stshi < 18` panic guard; and removing the forced `bold = true`.

### B5 — the fixture is validated by this repo's tolerances, not by the format

The synthetic CFB `.doc` parses only because `CfbReader::find_entry`
(`src/cfb/reader.rs:79-84`) is a flat linear scan that ignores the red-black
directory tree, so the fixture's broken sibling links never matter, and because
both streams sit under the 4096-byte mini-stream cutoff. A fixture that agrees
with the decoder is not evidence when both were written from the same model of
the format.

---

## Medium

- **Forced `bold = true` on headings** — the IR asserts formatting the document
  may not contain, and markdown renders `# **Section One**`; the DOCX path does
  not do this. **Not a defect of this PR:** `convert_doc.rs:467-471` on `main`
  already does exactly this, and the change mirrors the function it edits. Ours
  to fix, not the contributor's. (Correction: I first filed this against the PR.)
- **`outline_lvl_explicit` is the wrong shape.** SPRM application is an ordered
  fold — later Prl wins, and `sprmPIncLvl` (`0x2602`) *offsets* the level. A
  boolean "was it present" cannot express that.
- **The heuristic is bypassed, not retired.** The guess still runs unguarded
  behind the new arm, so a document can emit half its headings from real outline
  data and half from ALL-CAPS shape, at disagreeing levels.

  **Both blind rounds proposed a gate, and they disagree — the measurement picks
  A1.** B1 (Q2): fall back to the heuristic "only when `styles.is_empty()`".
  A1 (§5.5): only when "the document yielded no style-derived heading at all",
  warning that `metadata.title`/`Section.title` derive from the first
  `Element::Heading` (`convert_doc.rs:21-33`), so losing headings silently nulls
  the title for the CLI, every binding, and pdf_oxide.

  Measured (round C, guard fixed, 233 parsed files):

  | | value |
  |---|---|
  | docs with ≥1 heading today | 131 (all 131 have a title) |
  | docs with ≥1 **style-derived** heading | **47** |
  | docs B1's gate would strip of headings **and** title | **88** |
  | docs A1's gate would strip | **0** |

  A parsed style sheet is not evidence the document uses headings; a resolved
  heading is. The 88 in the gap are letters/memos/forms with a valid STSH and no
  heading styles in it. **Gate on the outcome, not the input.** B1 could not have
  known — it never ran a corpus; A1 named the mechanism without the number.
  Neither round alone lands this.
- **`heading_level_from_name` requires the literal prefix `"heading "`** — misses
  `Heading1` and every localized name. `sti` covers built-ins, so this only bites
  user-defined styles, which is exactly the case the name path exists for.
- **The style's own outline level is never read.** A user style based on
  `Heading 3` carries `sti = 0x0FFE` and an arbitrary name; its level lives in its
  `UpxPapx.grpprlPapx` and up the `StdfBase.istdBase` chain. Neither is consulted.

## Right, and worth saying

- `cbSTDBaseInFile` validated against `0x000A` / `0x0012` matches Stshif verbatim.
- `LPStd` even-byte padding and `cbStd == 0` = empty style match LPStd verbatim.
- STSH layout (`LPStshi` = `cbStshi` + `STSHI`, then `rglpstd`) is correct; the
  description records that an earlier revision had it wrong and self-corrected.
- `StdfBase.sti` as the low 12 bits is correct.
- **Precedence: the contributor is right and I was wrong** — see the correction
  in `round1-contaminated.md`. Making direct `sprmPOutLvl` beat the style
  contradicts the letter of [MS-DOC] ("MUST be ignored if the paragraph has an
  **istd** … between 0x1 and 0x9") but matches POI, which deleted that guard on
  purpose: *"Word seems to set outline levels even for paragraph with other
  styles than Heading 1..9, even though specification does not say so. See bug
  49820."* A1 reports LibreOffice never had the guard (inferred from A1; I did
  not verify the LibreOffice half at source).
- No panic, hang or unbounded allocation in the new code across ~14k hostile
  inputs (B2).

---

## Findings in code this PR does not touch

Filed separately — they are not review output and do not belong in a comment to
the contributor. Tracking: **N1 → #134, N2 → #135, N3 → #136, N4 → #137,
N5 → #138**; the heuristic-retirement follow-up is **#139**, and the mechanical
gates that would retire the recurring defect classes are **#133**. The review
itself was posted on #127.

### N1 — a public doc comment states the outline-level value space backwards

`src/ir.rs:743` and `src/docx/write.rs:264`, verbatim:
`/// Outline level (0 = body text, 1–9 = heading levels).`
`src/docx/formatting.rs:44`, verbatim:
`/// Outline level (0 = Heading 1, 1 = Heading 2, …).`

ECMA-376 §17.3.1.20: the value runs 0–9, "9 specifically indicates that there is
no outline level specifically applied to this paragraph", and omission defaults
to 9. The parser is right; the comment on the **public IR type** is wrong.

This is the origin of B3. Both blind rounds reached it from opposite directions:
A2 from the DOCX symptom, and the change under review implements it verbatim in
a completely different format's reader. `docx/write.rs:1604` emits the value
unchanged, so a downstream caller that trusts the comment (pdf_oxide depends on
this crate) writes `outlineLvl=0` on body paragraphs and every one becomes
Heading 1 in Word.

### N2 — DOCX renders explicit body text as a heading

`w:outlineLvl val="9"` is what Word writes for "Outline level: Body Text", and
what the built-in `TOCHeading` style uses to cancel the level it inherits from
`Heading1`. `src/convert_docx.rs:200-201` computes `(level + 1).min(6)` → **H6**.
The other renderer, `src/docx/text.rs:189`, does `"#".repeat(level.min(9))` and
emits up to **nine** hashes, which no markdown renderer treats as a heading.
Also `(level + 1)` is `u8` arithmetic on a parser-supplied value up to 255.
Found independently by A1 and A2.

### N3 — five renderers, two different heading clamps

`ir_render.rs:272` (markdown) uses `h.level.min(6)` with **no lower clamp**;
`ir_render.rs:441` (HTML), `docx/write.rs:525`, `docx/write.rs:1253` and
`create.rs:191` all use `clamp(1, 6)`. A level-0 heading silently becomes body
text in markdown only.

### N4 — two DOCX markdown renderers that disagree

A2 found seven heading-relevant divergences between `docx/text.rs` and
`convert_docx.rs` + `ir_render.rs`, the largest being that `ir_render.rs:253-258`
emits `Section.title` as a synthetic `## …` while `convert_docx.rs:70-84` sets
that title to the section's first heading — so the IR path duplicates every
section's first heading at the wrong level (same for PPTX). Existing tests use
`contains()` and are structurally blind to it.

### N5 — shift overflow on a corrupted CFB header

`src/cfb/header.rs:76`: `let sector_size = 1usize << sector_power;` where
`sector_power` is a `u16` read straight from offset `0x1E` of an untrusted file.
Panics in debug, wraps in release. Pre-existing, untouched by this PR, reachable
from every format the crate opens.

---

## SOTA design — what "fixed at the right level" looks like

1. **One value space, defined once.** Model outline level exactly as OOXML does —
   0-based, with 9 meaning body text — in a single type with a parsing
   constructor, and convert to `Heading::level` at one IR boundary. Both format
   readers then share one definition instead of two prose comments that
   contradict each other (N1).
2. **Make `Heading::level` enforce its own invariant.** A `HeadingLevel` newtype
   whose constructor clamps to 1..=6 retires N3 by construction; five renderers
   cannot diverge on a value they cannot construct out of range.
3. **Ship the built-in case with no style sheet at all.** [MS-DOC] STSH fixes
   `istd n → sti n` for n ≤ 9, and `sprmPIstd` says an `istd` of 1–9 "also
   specifies the outline level … equal to the value of the **istd** minus 1". So
   `istd ∈ 1..=9` alone gives Heading 1–9 in every conforming file. That is
   Antiword's entire implementation, it is ~60 lines, and it needs neither
   `fcStshf` nor a name table — which would have made B1 unreachable.
4. **Then resolve styles by their properties, not their identity.** STSH → LPStd
   → STD → `UpxPapx.grpprlPapx`, following `StdfBase.istdBase` up the chain with
   a visited set. A user style based on `Heading 3` and a localized `Überschrift 3`
   both work, with no name matching. `sti ∈ 1..=9` stays as a fast path and name
   matching becomes a last-resort fallback rather than the mechanism.
5. **Apply the grpprl as an ordered fold**, so `sprmPIstd` (0x4600),
   `sprmPOutLvl` (0x2640) and `sprmPIncLvl` (0x2602) compose in file order —
   last wins, `IncLvl` offsets — instead of a set of independent flags.
6. **Decide the heuristic per document, not per paragraph.** If any usable
   outline data exists, turn the shape guess off for the whole document.
7. **Guard on lengths, never on offsets.** `lcbStshf` is the field the spec makes
   nonzero; `fcStshf == 0` is ordinary.

### Testing, and the gate worth keeping

- Table-drive the SPRM operand over all 256 byte values and assert exactly
  `0..=8` → levels and `9` → body text. B2's mutation probes show the current
  suite pins the range at exactly one point (104).
- Assert `to_markdown() == to_ir().to_markdown()` for DOCX as a property (N4).
- **Assert the new path is *reached*, not merely that output is unchanged.** A
  corpus check that counts style sheets parsed and headings emitted would have
  turned B1 from a shipped no-op into a one-line failure. "Byte-identical output"
  is not evidence when the feature is switched off.
- **Convertible shape → permanent gate:** a CI job that reverts each new test's
  production hunk and requires red. That retires B4's class — nine decorative
  tests here, six flagship-passes-on-baseline in the previous run — for near-zero
  cost. This is the one finding class here that a mechanical check can replace.

## Method notes for next time

- **Worktree isolation leaks.** B2's `isolation: "worktree"` materialised the diff
  at `.claude/worktrees/agent-…/` *inside the repo*, and A2's repo-wide grep found
  it. A2 reported the leak and stopped; its findings are all in `main` code the
  diff never touches, so nothing was contaminated. Next time, put round-B
  worktrees outside the tree the round-A agents read.
- Parking `.research/` out of the tree for the run worked and is worth repeating.
