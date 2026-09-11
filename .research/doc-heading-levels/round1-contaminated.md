# PR #127 — round 1 (contaminated: my own review, diff in hand)

Written before any blind round reported. Every claim below carries its
verification status. Primary source fetched live from learn.microsoft.com.

## The opcode table is wrong — twice

Verified verbatim against [MS-DOC] 2.6.2 "Paragraph Properties"
(https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/484822ee-a9d9-4af4-8423-29fda67a6a58):

| the change says | actual sprm at that opcode | the real sprm it wanted |
|---|---|---|
| `0x6412` = `sprmPOutlineLvl` | **`sprmPDyaLine`** — "An LSPD value that specifies the spacing between lines in this paragraph." | **`sprmPOutLvl` (`0x2640`)**, 1-byte operand |
| `0x640A` = `sprmPStyle` | **no such paragraph sprm**; ispmd `0x0A` is `sprmPIlvl` (`0x260A`) | **`sprmPIstd` (`0x4600`)**, 2-byte operand |

So `src/doc/sprm.rs`'s new `0x6412` arm reads **line spacing** and interprets
its low byte as an outline level, and the `0x640A` arm is dead on real files.

## The level semantics are inverted and off by one

[MS-DOC] `sprmPOutLvl` (0x2640), verbatim:

> An unsigned 8-bit integer value that specifies the outline level of the
> paragraph. This value MUST be one of the following.
> **0x0 - 0x8** The value is the zero-based outline level that this paragraph is in.
> **0x9** The paragraph at any outline level; instead, the paragraph is body text.
> This MUST be ignored if the paragraph has an **istd** that is greater than or
> equal to 0x1 and less than or equal to 0x9. By default, paragraphs are body
> text, and are therefore not in any outline level.

The change implements "`0` = body text, `1`–`9` = Heading 1–9". Three errors:
1. **Off by one.** `0` is Heading 1, not body text.
2. **Sentinel inverted.** `9` is body text, not Heading 9.
3. **Precedence backwards.** The spec says the SPRM MUST be *ignored* when
   `istd` is 1–9; the change makes the SPRM override the style.

Also from `sprmPIstd` (0x4600), verbatim: "An **istd** value in the range of 1
to 9, inclusive, also specifies the outline level of the paragraph …, where the
new outline level is equal to the value of the **istd** minus 1." Combined with
[MS-DOC] STSH's fixed-index table (istd 0→sti 0, istd 1..9→sti 1..9), the
built-in heading level is available from `istd` alone — the STSH round trip the
change performs to recover `sti` from `istd` returns `istd` for exactly the
range it cares about.

The comment cites "MS-DOC §2.9.138" for the outline level. sprm operands are
enumerated in §2.6.2; I did not check what §2.9.138 actually is, only that it is
not where this is defined.

## Consequence: reachable false headings on ordinary documents

`LSPD` ([MS-DOC], verified): `dyaLine` (16-bit) then `fMultLinespace` (16-bit),
little-endian, so the operand's low byte is `dyaLine & 0xFF`.

- "Exactly" spacing is encoded as `0x10000 - twips`. `Exactly 12.5 pt` = 250
  twips → `dyaLine = 0xFF06` → low byte **6** → the paragraph becomes an **H6**.
- "At least" spacing is the raw twip count. `At least 13 pt` = 260 twips →
  `dyaLine = 0x0104` → low byte **4** → **H4**. 26 pt = 520 → **H8**.
- Low byte **0** (e.g. `dyaLine = 0x0200`, 25.6 pt) sets `outline_lvl_explicit`,
  which *suppresses* a genuine style-derived heading — the failure runs in both
  directions.

This is **inferred** arithmetic over a **specified** structure; not yet measured
against a real file.

## The corpus claim does not cover the SPRM path

The description says "Every file reports `fcStshf == 0`, so none reaches the new
STSH/outline code — the change is a no-op there." The `fcStshf` guard gates only
`parse_style_sheet`. `extract_pap_props` runs on every paragraph of every file,
so the `0x6412` arm **is** reached by the whole corpus; the corpus was
byte-identical because none of those 160 files happened to carry a `dyaLine`
whose low byte lands in 0..=9, not because the code was not executed.

## Smaller things

- `emit_prose` force-sets `bold = true` on every character of a styled heading.
  The IR then asserts bold runs the document may not contain, and the markdown
  renders `# **Section One**`. The DOCX path does not do this.
- `heading_level_from_name` requires the literal prefix `"heading "`. Misses
  `Heading1` (no space) and every non-English built-in name — though `sti`
  covers built-ins, so this only bites user-defined styles.
- Reading the style's own outline level is never attempted. A user style based
  on `Heading 3` carries `sti = 0x0FFE` and an arbitrary name; its level lives in
  its own `UpxPapx` grpprl and in the `istdBase` chain. Neither is read.

## What is right

- `cbSTDBaseInFile` validated against `0x000A` / `0x0012` matches Stshif
  verbatim ("MUST be 0x000A when the Stdf structure does not contain an
  StdfPost2000 structure and MUST be 0x0012 when it does").
- `LPStd` even-byte padding, and `cbStd == 0` meaning an empty style, match the
  LPStd definition verbatim.
- STSH layout (`LPStshi` = `cbStshi` + `STSHI`, then `rglpstd`) is correct, and
  the description says an earlier revision had it wrong and was fixed.
- `StdfBase.sti` as the low 12 bits is correct.

---

## CORRECTION (round C, after round A1) — error #3 above is overstated

> **Struck:** "3. **Precedence backwards.** The spec says the SPRM MUST be
> *ignored* when `istd` is 1–9; the change makes the SPRM override the style."

The spec clause is quoted correctly, but citing it as a defect is wrong: the
clause is **disputed in practice and deliberately not implemented** by the
reference readers.

Apache POI, `poi-scratchpad/.../sprm/ParagraphSprmUncompressor.java`, verbatim:

```java
case 0x40:
    // This condition commented out, as Word seems to set outline levels even for
    //  paragraph with other styles than Heading 1..9, even though specification
    //  does not say so. See bug 49820 for discussion.
    //if (newPAP.getIstd () < 1 && newPAP.getIstd () > 9)
{
    newPAP.setLvl((byte) sprm.getOperand());
}
```

Round A1 reports LibreOffice never had the guard either (I have not verified
the LibreOffice half at source; **inferred from A1**).

So "direct `sprmPOutLvl` wins over the style" — the precedence the change
implements — is what real readers do. **This item is withdrawn as a defect.**
The remedy for the other two errors is unchanged; only this justification moves.

The same POI file independently confirms the opcode finding from a second
source: `case 0x12` is `newPAP.setLspd(new LineSpacingDescriptor(...))`.

## ADDITION (round C, after rounds A1/A2) — where the wrong model came from

`src/ir.rs:743` and `src/docx/write.rs:264`, verbatim:
`/// Outline level (0 = body text, 1–9 = heading levels).`
`src/docx/formatting.rs:44`, verbatim:
`/// Outline level (0 = Heading 1, 1 = Heading 2, …).`

ECMA-376 §17.3.1.20: values 0–9, "9 specifically indicates that there is no
outline level specifically applied to this paragraph", omission defaults to 9.
The parser is right; the doc comment on the **public IR type** is wrong, and it
states exactly the model this change encodes into the `.doc` reader. This is a
shipped comment asserting a property the code does not have, with a now-measured
downstream cost.
