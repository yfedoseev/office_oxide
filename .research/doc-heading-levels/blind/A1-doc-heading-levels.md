# A1 — Legacy `.doc` heading levels: root-cause research

Repository baseline: `/home/yfedoseev/projects/office_oxide`, branch `main`, commit `5007d9c`
(clean; read-only investigation, nothing built, nothing modified).

Every correctness claim below is tagged **[specified]** (I read it in [MS-DOC] on
learn.microsoft.com, or in a real implementation's source) or **[inferred]** (I reasoned it out
from the code / the format and it is not stated anywhere I could cite).

---

## 1. Root cause

**There is no heading detection in the `.doc` path at all.** The `.doc` reader never reads the
paragraph style index or the outline level from the file. What the user sees is a text-shape
guess applied to *every* paragraph.

### 1.1 The guess

`/home/yfedoseev/projects/office_oxide/src/convert_doc.rs:447-483`

```rust
fn emit_prose(text: &str, tabs: &[TabStop], elements: &mut Vec<Element>) {
    let trimmed = text.trim();
    if trimmed.is_empty() { return; }

    let is_heading = trimmed.len() < 100                        // 453
        && !trimmed.ends_with('.')                              // 454
        && !trimmed.ends_with(',')                              // 455
        && (trimmed
            .chars()
            .filter(|c| c.is_alphabetic())
            .all(|c| c.is_uppercase())                          // 456-459  <- the ALL-CAPS rule
            || (elements.is_empty() && trimmed.len() < 60));    // 460      <- the "first line" rule

    if is_heading {
        ...
        elements.push(Element::Heading(Heading {
            level: if elements.is_empty() { 1 } else { 2 },     // 472      <- the level
            ...
```

This single function explains all three symptoms in the report, exactly:

| Symptom | Line | Mechanism |
| --- | --- | --- |
| "the level is never right" | `convert_doc.rs:472` | `level` is a hard-coded ternary: 1 if this is the first element in the document, otherwise 2. It is **not** derived from anything in the file. A "Heading 3" can only ever come out as `#` or `##`. |
| "only picks up headings typed in ALL CAPS" | `convert_doc.rs:456-459` | `.all(char::is_uppercase)` over the alphabetic characters is the only rule that can fire after the first element. A normally-cased "Background" styled Heading 2 fails it and falls to `Element::Paragraph` at line 477. |
| "most of my headings … end up as ordinary paragraphs" | same | Anything mixed-case, ≥100 chars, or ending in `.`/`,` is a paragraph regardless of its style. |

Note also the accidental false positives: any ALL-CAPS run of text — an acronym-only line
(`FAQ`, `NOTE:`), a name in caps, a caps table caption that escaped the table path, a caps
signature block — becomes an `Element::Heading` with level 2. **[inferred, from reading the
predicate]**

`emit_prose` is reached from two places:
* `walk_paragraphs` (`convert_doc.rs:373`) — the structured path used when the FIB advertises a
  `PlcfBtePapx`;
* `line_heuristic` (`convert_doc.rs:489-493`) — the fallback that splits the sanitised plain text
  on `'\n'` and calls `emit_prose` per line.

So both paths share the same guess. The doc comment at `convert_doc.rs:445-446` says this is
deliberate ("Mirrors the line-based heuristic so a PAPX-bearing document keeps the same
heading/title detection as the fallback path") — i.e. the structured path was made to *match* the
textual guess rather than to supersede it.

### 1.2 The information that is thrown away

**`istd` is parsed out of the PAPX and immediately discarded.**

`/home/yfedoseev/projects/office_oxide/src/doc/papx.rs:170`

```rust
let grpprl_start = p + 3; // skip cw (1) + istd (2)
```

The two `istd` bytes of the `GrpPrlAndIstd` ([MS-DOC] 2.9.114) are skipped over and never
returned. `FkpParagraph` (`papx.rs:24-32`) carries only `fc_start`, `fc_end`, `grpprl` — there is
no field for the style index.

**No outline-level SPRM is decoded.** `extract_pap_props` (`/home/yfedoseev/projects/office_oxide/src/doc/sprm.rs:378-434`)
dispatches on exactly six opcodes: `0x2416`, `0x6649`, `0xD608`, `0x460B`, `0x260A`,
`0xC615|0xC60D`. `grep` for `0x2640` (`sprmPOutLvl`), `0x4600` (`sprmPIstd`), `0x2602`
(`sprmPIncLvl`), `0x6646` (`sprmPHugePapx`), `0x646B` (`sprmPTableProps`) across `src/` returns
nothing. `PapProps` (`sprm.rs:190-218`) has no `istd` and no `outline_level` field.

**The stylesheet is never opened.** `grep -ri "stsh"` across `src/doc/` returns nothing. `Fib`
(`/home/yfedoseev/projects/office_oxide/src/doc/fib.rs:11-44`) exposes `clx_offset/clx_size`,
`fc_plcf_bte_papx/lcb_plcf_bte_papx`, `fc_plcf_lst/lcb_plcf_lst` and the text lengths — but not
`fcStshf`/`lcbStshf`. So no style can be resolved even in principle today.

**`Pcd.Prm` is read but never applied.** `piece_table.rs:105` documents the PCD layout as
`[u16 unused][u32 fc][u16 prm]` and the `prm` field is parsed past and dropped. Per [MS-DOC]
2.4.6.1 step 5 **[specified]**, `Pcd.Prm` can append further paragraph `Prl`s — including
`sprmPIstd`/`sprmPOutLvl` — to a paragraph's direct formatting. Complex (`fComplex`) documents
that carry a style change in `Prm0`/`Prm1` will therefore lose it.

### 1.3 Why the `.docx` of the same document is "much closer to right"

`/home/yfedoseev/projects/office_oxide/src/convert_docx.rs:434-447` does the real thing:

```rust
fn resolve_heading_level(p, doc) -> Option<u8> {
    let props = p.properties.as_ref()?;
    if let Some(lvl) = props.outline_level { return Some(lvl); }   // direct w:outlineLvl
    let style_id = props.style_id.as_ref()?;
    let styles = doc.styles.as_ref()?;
    styles.resolve_outline_level(style_id)                          // walk w:basedOn chain
}
```

with `StyleSheet::resolve_outline_level` (`src/docx/styles.rs:89-105`) walking the `basedOn`
chain with a depth cap of 20, and the caller mapping `level+1` clamped to 6
(`convert_docx.rs:200-202`). The `.doc` side has no equivalent because it has no stylesheet
reader. The asymmetry the reporter observed is exactly this.

---

## 2. Correct behaviour according to [MS-DOC]

All section numbers below were read on learn.microsoft.com during this investigation; URLs are
given so they can be re-checked.

### 2.1 The two carriers of outline level

**[specified]** [MS-DOC] **2.6.2 Paragraph Properties**
(<https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/484822ee-a9d9-4af4-8423-29fda67a6a58>)
defines two relevant SPRMs. Verbatim:

> **sprmPOutLvl (0x2640)** — "An unsigned 8-bit integer value that specifies the outline level of
> the paragraph. This value MUST be one of the following.
> **0x0 - 0x8** — The value is the zero-based outline level that this paragraph is in.
> **0x9** — The paragraph at any outline level; instead, the paragraph is body text.
> This MUST be ignored if the paragraph has an **istd** that is greater than or equal to 0x1 and
> less than or equal to 0x9. By default, paragraphs are body text, and are therefore not in any
> outline level."

(The "0x9" wording is garbled in the published text — the intent, confirmed by the parallel
OOXML rule, is *"the paragraph is not at any outline level; instead the paragraph is body text"*.)

> **sprmPIstd (0x4600)** — "An unsigned integer that specifies the **istd** of a paragraph style
> to apply. … An **istd** value in the range of 1 to 9, inclusive, also specifies the outline
> level of the paragraph (for example, by sprmPOutLvl), where the new outline level is equal to
> the value of the **istd** minus 1."

Also relevant, and not currently handled:

> **sprmPIncLvl (0x2602)** — "A signed 8-bit integer value. If the paragraph has an **istd** that
> is greater than or equal to 0x0001 and less than or equal to 0x0009, this value specifies an
> offset to the **istd** of the paragraph. … If the **istd** of the paragraph is not within the
> range that was specified earlier, this value specifies an offset to the outline level of the
> paragraph, unless the outline level of the paragraph is equal to 0x09, in which case this value
> MUST be ignored."

So the format gives you **two** sources, and they are not symmetric: the istd rule wins.

### 2.2 Where the paragraph's `istd` comes from

**[specified]** [MS-DOC] **2.4.6.1 Direct Paragraph Formatting**
(<https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/61b635c3-2c44-4155-bf17-fec281b30c71>),
steps 2-5: find the `BxPap` in the `PapxFkp`, find the `PapxInFkp` at `of + 2 × BxPap.bOffset`,
find the `GrpprlAndIstd` in it, take `grpprl`, and *then* append any paragraph `Prl`s from
`Pcd.Prm`.

**[specified]** [MS-DOC] **2.9.114 GrpPrlAndIstd** — the structure is `istd` (2 bytes) followed by
`grpprl`. That `istd` is the paragraph's style index; it is the field skipped at `papx.rs:170`.

**[specified]** [MS-DOC] **2.4.6.6 Determining Formatting Properties**
(<https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/d8b66123-1a3d-4e06-94c5-5ab16e9b6417>),
Part 2 steps 1-2: apply the property modifications implied by `GrpprlAndIstd.istd` **first**,
then apply `GrpprlAndIstd.grpprl`. A `sprmPIstd` inside the grpprl therefore overrides the
`GrpprlAndIstd.istd`, and a `sprmPOutLvl` in the grpprl overrides an outline level inherited from
the style — subject to the "ignore if istd ∈ 1..9" clause above.

### 2.3 Why `istd` 1..9 is meaningful without reading the stylesheet

**[specified]** [MS-DOC] **2.9.271 STSH**
(<https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/c8ee0f39-02c3-4caa-b27a-6a97600130fe>):

> "The beginning of the **rglpstd** array is reserved for specific 'fixed-index'
> application-defined styles. A particular fixed-index, application-defined style has the same
> istd value in every stylesheet."

with the table `istd 0 → sti 0`, `istd 1 → sti 1`, … `istd 9 → sti 9`, `istd 10 → sti 65`,
`istd 11 → sti 105`, `istd 12 → sti 107`. **[specified]** `Stshif.istdMaxFixedWhenSaved`
([MS-DOC] 2.9.274) "MUST be 0x000F".

**[specified]** [MS-DOC] **2.9.260 StdfBase**
(<https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/df0f4654-071d-442f-8563-752d7e0285ef>):
`sti` is "the invariant style identifier for application-defined styles, or 0x0FFE for
user-defined styles… The **sti** values correspond to the 'Index within Built-in Styles' table
column that is specified in [ECMA-376] part 1, section 17.7.4.9 (name)."

**[specified]** [MS-DOC] Glossary
(<https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/951dd5ff-6eb5-4265-b8c2-f4b7f3d745ca>):

> **heading style**: "A type of paragraph style that also specifies a heading level. There are as
> many as nine built-in heading styles, Heading 1 through Heading 9."
>
> **outline level**: "A type of paragraph formatting that can be used to assign a hierarchical
> level, Level 1 through Level 9, to paragraphs in a document."

Together: `istd` 1..9 *is* Heading 1..Heading 9 in every conforming file, regardless of the
stylesheet contents and regardless of the UI language the document was written in. This is why
the istd rule can be implemented **before** a stylesheet reader exists, and it is the highest
value-per-line change available. **[specified]**

### 2.4 When you do need the stylesheet

A paragraph styled with a **user-defined** style (istd ≥ 15, or a fixed-index style other than
1..9) that nevertheless has an outline level — "Chapter Title", "Appendix Heading", or a
renamed/derived heading style — carries its outline level in the style's own `grpprlPapx`, as a
`sprmPOutLvl`. **[specified]** [MS-DOC] 2.9.338 `UpxPapx`
(<https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/659f448a-473a-44c6-8b69-a25984af3645>)
lists the SPRMs a paragraph style MUST NOT contain — `sprmPIstd`, `sprmPIstdPermute`,
`sprmPIncLvl`, `sprmPChgTabs`, `sprmPHugePapx`, … — and `sprmPOutLvl` is **not** on that list, so
a style may and does carry it.

Resolution algorithm, **[specified]** [MS-DOC] 2.4.6.5 *Determining Properties of a Style*
(<https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/9258b41c-ff0a-4c96-a3a9-610664dabbeb>):

1. `FibRgFcLcb97.fcStshf` / `lcbStshf` → an `STSH` in the Table stream. These are pair index **1**
   of `FibRgFcLcb97` (`fcStshfOrig`/`lcbStshfOrig` are pair 0 and MUST be ignored), i.e. byte
   offset 8 into `FibRgFcLcb97`. **[specified]**, from the field order on
   <https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/0c9df81f-98d0-454e-ad84-b612cd05b1a4>.
   Since this repository already hard-codes `fcPlcfBtePapx` at absolute FIB offset `0x0102`
   (= `FibRgFcLcb97` base `0x9A` + 104, pair 13), the matching absolute offsets are
   **`fcStshf` = 0x00A2, `lcbStshf` = 0x00A6** for `nFib` 0x00C1. **[inferred]** — arithmetically
   derived from the spec's field order plus the offset this codebase already validates against
   real files; worth confirming against a real document before relying on it.
2. `STSH` = `LPStshi` then `rglpstd`. `LPStshi` is `cbStshi` (2 bytes) + `stshi`, so
   **`rglpstd` starts at `fcStshf + 2 + cbStshi`** — you never have to know the `STSHI` layout to
   find the style array. **[specified]** ([MS-DOC] 2.9.136, 2.9.272).
3. `rglpstd[istd]` is an `LPStd`: `cbStd` (2 bytes) + `std`. `cbStd == 0` means the style is
   *empty* (a legal state for the fixed-index styles 0..12). **[specified]** ([MS-DOC] 2.9.271).
4. `STD` = `stdf` + `xstzName` + `grLPUpxSw`. `stdf` is `cbSTDBaseInFile` bytes long — `Stshif.cbSTDBaseInFile`
   "MUST be 0x000A when the Stdf structure does not contain an StdfPost2000 … and MUST be 0x0012
   when [it] does". **[specified]** ([MS-DOC] 2.9.274, 2.9.258).
5. `StdfBase.istdBase` (12 bits) is the parent style, or `0x0FFF` for "no parent". Recurse.
   Prepend the parent's `Prl` array to the child's, so the child wins. **[specified]**
   ([MS-DOC] 2.9.260, 2.4.6.5 steps 5 and 8).

### 2.5 The correct derivation, stated once

**[specified]**, assembled from 2.6.2, 2.4.6.1, 2.4.6.5 and 2.4.6.6:

```
outline_level(paragraph):
    istd  := GrpprlAndIstd.istd                       # from the PAPX in the FKP
    lvl   := None                                     # "body text" is the default

    # 1. style contribution (recursive over istdBase, prepend-parent-first)
    for each sprm in resolved_style_grpprlPapx(istd):
        if sprm == sprmPOutLvl: lvl := operand

    # 2. direct contribution, applied left to right
    for each Prl in grpprl ++ Prm-derived Prls:
        if Prl == sprmPIstd:   istd := operand; lvl := (recompute from the new style)
        if Prl == sprmPOutLvl: lvl := operand
        if Prl == sprmPIncLvl: ...                    # istd offset if istd in 1..9, else lvl offset

    # 3. the override that must come LAST
    if 1 <= istd <= 9:  return istd - 1               # 2.6.2 sprmPIstd; sprmPOutLvl is ignored
    if lvl is None or lvl == 9: return None           # body text
    return lvl                                        # 0..8
```

> **Caveat — step 3 is disputed by every major implementation.** [MS-DOC] 2.6.2 says
> `sprmPOutLvl` "MUST be ignored if the paragraph has an istd … 0x1 … 0x9". Apache POI
> *deliberately deleted* that check with a source comment saying Word violates it, and
> LibreOffice never implemented it; only wvWare follows the letter. See §3.5 conclusion 3 — the
> model I actually recommend is LibreOffice's: `istd ∈ 1..9` **seeds** the level at `istd − 1`,
> and an explicit `sprmPOutLvl` (from the style or direct) **overrides** it, including
> `sprmPOutLvl == 9` meaning "Word demoted this heading to body text". **[specified]** that the
> spec says otherwise; **[inferred]** that following the implementations is the better choice.

and then, at the IR boundary:

```
heading_level = outline_level + 1      # 0-based file value -> 1-based markdown depth
markdown depth = clamp(heading_level, 1, 6)
```

**[specified]** that the file value is zero-based (2.6.2 sprmPOutLvl: "the zero-based outline
level"); **[specified]** that the user-facing range is 1..9 (glossary: "Level 1 through Level 9");
**[inferred]** that levels 7-9 should clamp to markdown `######` rather than being demoted to a
paragraph — the format has 9 levels and markdown has 6, and the spec says nothing about the
mapping. Clamping is what this repository already does on the DOCX side
(`convert_docx.rs:201`), so consistency argues for it.

### 2.6 Where the spec is silent or ambiguous

* **Whether a user style *based on* a built-in heading inherits that heading's outline level.**
  2.4.6.5 prepends the base style's `Prl` array — but Word does not necessarily write a
  `sprmPOutLvl` into Heading 1's own `grpprlPapx`, because for istd 1..9 the outline level is
  implied by the istd (2.6.2). So a user style with `istdBase == 3` may resolve to an empty
  `Prl` array and thus no outline level, even though a human would call it a Heading 3. The spec
  does not say whether `istdBase ∈ 1..9` should propagate the implied level. **[inferred]** —
  I could not find any statement either way. LibreOffice resolves it: `WW8RStyle::PrepareStyle`
  copies the parent's outline level down (`if (!rSI.IsWW8BuiltInHeadingStyle()) {
  rSI.mnWW8OutlineLevel = pj->mnWW8OutlineLevel; }`), so a derived style **does** inherit the
  level while a built-in heading keeps its own `sti − 1`. See §3.2. **[specified]** as
  LibreOffice's behaviour.
* **Word 6/95 (`wIdent == 0xA5DC`).** [MS-DOC] documents Word 97-2007. Word 6/95 files use 1-byte
  SPRM opcodes and a 10-byte `TC`, and this repository's `Fib::parse`
  (`src/doc/fib.rs:58-61`) *accepts* `0xA5DC` and then reads Word-97 absolute FIB offsets. Any
  outline-level work is undefined for those files and must not be assumed to work.
  **[inferred]**
* **No rule ties outline level to "is a heading" for rendering purposes.** The format has outline
  levels and heading styles; markdown has `#`. Choosing "outline level present and != 9 ⇒ emit
  `Element::Heading`" is a product decision, not a spec requirement. **[inferred]**

---

## 3. What production `.doc` readers actually do

Four independent implementations were read at source level. **None of them uses a text-shape
heuristic.** All four key off the paragraph's style index and/or the outline-level SPRM. They
disagree in interesting ways about one clause of the spec.

### 3.1 Apache POI HWPF (`poi-scratchpad`, trunk)

**[specified]** `model/types/PAPAbstractType.java`
(<https://raw.githubusercontent.com/apache/poi/trunk/poi-scratchpad/src/main/java/org/apache/poi/hwpf/model/types/PAPAbstractType.java>)
stores the outline level as `field_41_lvl` with **`field_41_lvl = 9;` as the constructed
default** — i.e. "body text" is the default state, exactly as [MS-DOC] 2.6.2 says. `ilvl` (list
level) is a *separate* field; POI does not conflate them.

`usermodel/Paragraph.java`
(<https://raw.githubusercontent.com/apache/poi/trunk/poi-scratchpad/src/main/java/org/apache/poi/hwpf/usermodel/Paragraph.java>):

```java
/** Returns the heading level (1-8), or 9 if the paragraph isn't in a heading style. */
public int getLvl() { return _props.getLvl(); }
public short getStyleIndex() { return _istd; }
```

Resolution order in `Paragraph.newParagraph` — **defaults → the style's PAPX (looked up by
`istd`) → the paragraph's own PAPX**:

```java
properties.setIstd( papx.getIstd() );
properties = newParagraph_applyStyleProperties( styleSheet, papx, properties );
properties = ParagraphSprmUncompressor.uncompressPAP( properties, papx.getGrpprl(), 2 );
```

**The notable divergence from the spec** — `sprm/ParagraphSprmUncompressor.java`
(<https://raw.githubusercontent.com/apache/poi/trunk/poi-scratchpad/src/main/java/org/apache/poi/hwpf/sprm/ParagraphSprmUncompressor.java>):

```java
case 0x40:                                // sprmPOutLvl (0x2640)
    // This condition commented out, as Word seems to set outline levels even for
    //  paragraph with other styles than Heading 1..9, even though specification
    //  does not say so. See bug 49820 for discussion.
    //if (newPAP.getIstd () < 1 && newPAP.getIstd () > 9)
    { newPAP.setLvl((byte) sprm.getOperand()); }
    break;
```

POI **deliberately removed** the "[MS-DOC] 2.6.2: ignore sprmPOutLvl if istd ∈ 1..9" guard,
because real Word output violates it. This is a first-hand report from a production reader that
the spec clause is not safe to implement literally.

`sti` is parsed (`StdfBaseAbstractType.getSti()`, `BitField(0x0FFF)`) but `StyleDescription`
keeps `_stdfBase` private with no getter, so **POI's public API cannot reach the built-in style
identifier** — only the istd index and the style name.

**POI's converters emit no headings at all.** `converter/WordToHtmlConverter.processParagraph`
(<https://raw.githubusercontent.com/apache/poi/trunk/poi-scratchpad/src/main/java/org/apache/poi/hwpf/converter/WordToHtmlConverter.java>)
creates a `<p>` unconditionally; grepping `AbstractWordConverter`, `WordToHtmlConverter`,
`WordToFoConverter`, `WordToTextConverter` and `WordExtractor` for `Heading` / `getLvl` /
`getStyleDescription` / `<h1>` yields zero hits. POI exposes the mechanism and leaves the
classification to the caller — so it is a good source for *how to read the field*, and no
guidance at all on *what to render*.

### 3.2 LibreOffice `sw/source/filter/ww8` — the most complete implementation

**[specified]** `sw/source/filter/ww8/ww8par.hxx`
(<https://raw.githubusercontent.com/LibreOffice/core/master/sw/source/filter/ww8/ww8par.hxx>)
states the convention outright:

```cpp
    // WW8 outline level is zero-based:
    // 0: outline level 1
    // ...
    // 8: outline level 9
    // 9: body text
    sal_uInt8 mnWW8OutlineLevel;
```

initialised to `MAXLEVEL` (= 10, `sw/inc/swtypes.hxx`) meaning **"unset"** — a third state,
distinct from both "level N" and "body text". And the conversion:

```cpp
    void SetOrgWWIdent( const OUString& rName, const sal_uInt16 nId ) {
        m_sWWStyleName = rName;  m_nWWStyleId = nId;
        // apply default WW8 outline level to WW8 Built-in Heading styles
        if (IsWW8BuiltInHeadingStyle()) { mnWW8OutlineLevel = m_nWWStyleId - 1; }
    }
    bool IsWW8BuiltInHeadingStyle() const { return GetWWStyleId() >= 1 && GetWWStyleId() <= 9; }

    static sal_uInt8 WW8OutlineLevelToOutlinelevel(const sal_uInt8 nWW8OutlineLevel) {
        if (nWW8OutlineLevel < MAXLEVEL) {
            if (nWW8OutlineLevel == 9) return 0;          // no outline level --> body text
            else return nWW8OutlineLevel + 1;             // outline level 1..9
        }
        return 0;
    }
```

Crucially, `m_nWWStyleId` is the **`sti`**, not the istd — set from the style definition in
`WW8RStyle::Import1Style` (`ww8par2.cxx`): `rSI.SetOrgWWIdent( sName, xStd->sti );`.

Style → Writer paragraph style mapping is by `sti`, `writerwordglue.cxx`
(<https://raw.githubusercontent.com/LibreOffice/core/master/sw/source/filter/ww8/writerwordglue.cxx>):

```cpp
static const SwPoolFormatId aArr[]= {
    SwPoolFormatId::COLL_STANDARD, SwPoolFormatId::COLL_HEADLINE1, ... COLL_HEADLINE9, ... };
//If this is a built-in word style that has a built-in writer
//equivalent, then map it to one of our built in styles regardless
//of its name
if (static_cast<size_t>(eSti) < std::size(aArr) && aArr[eSti] != RES_NONE)
    pRet = ...GetTextCollFromPool( aArr[eSti], false);
```

with `ww::sti` (`sw/source/filter/inc/wwstyles.hxx`) giving `stiNormal = 0`,
`stiLev1 = 1` … `stiLev9 = 9`, `stiUser = 0x0ffe`, `stiNil = 0x0fff`. The comment
**"regardless of its name"** is the design statement: LibreOffice never matches on the localized
style name for built-ins; name matching is only the fallback for styles with no built-in
equivalent.

`sprmPOutLvl` is handled for **both** styles and direct formatting, `ww8par6.cxx`
`SwWW8ImplReader::Read_POutLvl`
(<https://raw.githubusercontent.com/LibreOffice/core/master/sw/source/filter/ww8/ww8par6.cxx>) —
and, like POI, **without** the `istd ∈ 1..9` guard, even though the dispatch table still repeats
the spec text as a comment (`// pap.lvl;has no effect if pap.istd is < 1 or is > 9;byte;`).

Style inheritance, `ww8par2.cxx` `PrepareStyle` — this answers the ambiguity flagged in §2.6:

```cpp
if (!rSI.IsWW8BuiltInHeadingStyle()) { rSI.mnWW8OutlineLevel = pj->mnWW8OutlineLevel; }
```

i.e. **a non-heading style inherits its parent's outline level**, and a built-in heading style
keeps its own `sti − 1`. LibreOffice does **not** implement `sprmPIncLvl` (0x2602) or
`sprmPIstdPermute` (0xC601) — both dispatch to `nullptr`.

### 3.3 wvWare (`AbiWord/wv`, master)

**[specified]** `wv.h` declares `S8 lvl;` on the PAP and `sprmPOutLvl = 0x2640`; `pap.c`
`wvInitPAP` sets `item->lvl = 9;`. `wvInitPAPFromIstd` copies the style's PAP **and** its name
(`strncpy(apap->stylename, stsh->std[istdBase].xstzName, ...)`), and `wvAssembleSimplePAP` calls
`wvInitPAPFromIstd(apap, papx->istd, &ps->stsh)` before applying the direct grpprl.

wvWare **keeps** the spec guard that POI and LibreOffice dropped — `sprm.c`:

```c
case sprmPOutLvl:
    /*has no effect if pap.istd is < 1 or is > 9 */
    temp8 = bread_8ubit (pointer, pos);
    if ((apap->istd >= 1) && (apap->istd <= 9)) apap->lvl = temp8;
    break;
```

Its output side is name-driven in principle (`xml/wvHtml.xml` has `<style name="Heading 1">`
blocks emitting `<H1>`), but **every such block in `wvHtml.xml`, `wvDocbook.xml` and `wvAbw.xml`
is inside an XML comment**, and `wvConfig.c`'s `case TT_STYLE:` handler is a trace-only no-op.
So master wvWare parses the level correctly and then emits the same wrapper for every paragraph.
Not a model to copy.

### 3.4 Antiword

**[specified]** Antiword detects headings from **`istd` 1..9 alone** — no `sti`, no name, and it
never handles `sprmPOutLvl` (grep for `2640` over the tree returns nothing; its PAP sprm switch
in `prop8.c` reads `0x4600` and discards it).

`stylelist.c`:

```c
if (pStyle->usIstd >= 1 && pStyle->usIstd <= 9) {
    /* These are heading levels */
    return FALSE;   /* bStyleImpliesList */
}
```

`xml.c` `vSetHeadersXML` is the only place heading structure is emitted, and it is DocBook-only:

```c
if (usIstd == 0 || usIstd > 6) { ...; return; }
if (bTableOpen || uiListLevel != 0) {
    /* No headers when you're in a table or in a list */
    return;
}
```

Two behaviours worth stealing verbatim: **istd > 6 is dropped rather than clamped**, and
**headings are suppressed inside tables and inside lists** — exactly the two regression classes
predicted in §5.2 and §5.3, solved by the simplest possible rule. `stylesheet.c` treats
`ISTD_INVALID`/`STI_NIL`/`STI_USER` as "use the default style", and `usStc2istd` documents the
Word 1/2 mapping ("Heading 1 through 9 must become istd 1 through 9").

### 3.5 What the comparison actually tells us

| | Primary key | `sti` used? | Style name used? | `sprmPOutLvl` (0x2640)? | "Not a heading" |
| --- | --- | --- | --- | --- | --- |
| POI HWPF | exposes `istd` + `lvl`; classifies nothing | parsed, not publicly reachable | available, unused | yes — **istd guard removed** (bug 49820) | `lvl == 9` |
| LibreOffice ww8 | `istd` → STD → **`sti`** | **yes, primary**; `sti − 1` seeds the level | fallback only | yes, styles + direct, **no istd guard** | WW8 `9` → level 0 = body text; `MAXLEVEL` = unset |
| wvWare | `istd` → STSH → PAP + name | no | in config only (commented out) | yes, **with** the istd guard | `lvl == 9` |
| Antiword | **`istd` 1..9 directly** | no | no | **never handled** | `istd == 0`; `istd > 6` dropped |

Four conclusions relevant to this repository:

1. **Nobody guesses from text.** The heuristic at `convert_doc.rs:453-460` has no counterpart in
   any of them.
2. **`istd`/`sti` 1..9 is the load-bearing signal**, and Antiword shows it is sufficient on its
   own to get heading levels right on real documents — which is why Stage A of §4 is worth
   shipping before any stylesheet reader exists.
3. **The "[MS-DOC] 2.6.2: ignore sprmPOutLvl when istd ∈ 1..9" clause is disputed by
   implementations.** POI removed it with a written justification ("Word seems to set outline
   levels even for paragraph with other styles than Heading 1..9, even though specification does
   not say so"); LibreOffice never had it; only wvWare implements it literally. **[specified]**
   The safer model is LibreOffice's: `sti/istd ∈ 1..9` *seeds* the level at `id − 1`, and an
   explicit `sprmPOutLvl` — from the style or from direct formatting — **overrides** it. That
   also makes the "Word demoted this heading to body text" case (`sprmPOutLvl == 9` on a
   Heading-3-styled paragraph) come out right, which the literal spec reading gets wrong.
4. **Suppressing headings inside tables and lists is established practice**, not a hack —
   Antiword does it explicitly with a comment saying so.

---

## 4. Recommended fix — **generic**, in two stages

### 4.1 The recommendation

**Stage A (the correctness core, generic).** Read the paragraph's `istd` and its outline level,
and derive the heading level from the format instead of from the text.

1. `src/doc/papx.rs` — stop discarding the istd. Add `istd: u16` to `FkpParagraph`, read it from
   `GrpPrlAndIstd` at `grpprl_start - 2` with a bounds check, and carry it into `DocParagraph`
   via `build_paragraphs`.
2. `src/doc/sprm.rs` — add three opcodes to `extract_pap_props`: `0x2640` (`sprmPOutLvl`, 1-byte),
   `0x4600` (`sprmPIstd`, 2-byte), and — optionally — `0x2602` (`sprmPIncLvl`, 1-byte signed).
   Add `istd: Option<u16>` and `outline_level: Option<u8>` to `PapProps`. The existing
   `parse_grpprl` walker already sizes all three correctly from `spra` (`0x2640>>13 == 1` → 1
   byte; `0x4600>>13 == 2` → 2 bytes); no walker change is needed.
3. A single resolver producing `Option<u8>`, following **LibreOffice's model** rather than the
   literal spec (§2.5 caveat, §3.5 conclusion 3):
   `istd ∈ 1..9` seeds the level at `istd − 1`; an explicit `sprmPOutLvl` then overrides it;
   `sprmPOutLvl == 9` (and "no level at all") returns `None` = body text. Do **not** implement
   [MS-DOC] 2.6.2's "ignore sprmPOutLvl when istd ∈ 1..9" clause — POI deleted it with a written
   justification and LibreOffice never had it.
4. `src/convert_doc.rs` — `emit_prose` takes the resolved level and emits
   `Heading { level: (outlvl + 1).clamp(1, 6) }`, matching `convert_docx.rs:200-202` exactly.
5. Adopt Antiword's two suppression rules verbatim (`xml.c` `vSetHeadersXML`, §3.4): **no
   headings inside a table, no headings inside a list.** They are one `if` each, they are what a
   shipping reader does, and they pre-empt §5.2 and §5.3.

Stage A needs **no stylesheet reader** and no new FIB field. It is roughly 60 lines and it fixes
the reported defect for every document whose headings use the built-in Heading 1-9 styles, which
is the overwhelming majority.

**Stage B (the completeness half, also generic).** Add an `STSH` reader
(`fcStshf`/`lcbStshf` → `LPStshi` → `rglpstd` → `LPStd` → `STD` → `grLPUpxSw` → `UpxPapx`) with
`istdBase` recursion capped at a small depth, and consult it when the paragraph's istd is not in
1..9. This picks up user-defined and derived heading styles. It is strictly additive and can ship
separately; Stage A must not be blocked on it.

Stage B also lets the "is this a heading style" test be upgraded from `istd ∈ 1..9` to
`StdfBase.sti ∈ 1..9`, which is what LibreOffice keys on (§3.2). For conforming files the two are
identical — [MS-DOC] 2.9.271 fixes `istd n → sti n` for n ≤ 9 — so this is a robustness upgrade
for non-conforming producers, not a behaviour change. When resolving a user style, inherit the
outline level down the `istdBase` chain unless the style is itself a built-in heading, per
LibreOffice's `PrepareStyle`.

The two stages together mirror the DOCX side one-for-one — `props.outline_level` first, then
`styles.resolve_outline_level(style_id)` — which is the right architectural shape for this
repository, because it makes `.doc` and `.docx` of the same document produce the same IR.

### 4.2 Why generic, not a point patch

A point patch here would be something like "recognise ALL-CAPS *and* also bump the level by
counting how many previous headings there were", or "add a lookup for the specific style names
this reporter's documents use". Both would be wrong for the same reason: **the file already
carries the answer, unambiguously, in a field the parser walks past.** The current behaviour is
not an incomplete implementation of heading detection — it is a *substitute* for one. There is no
narrow version of "read the field".

Concretely, the narrow fix cannot be made correct because:

* The level is not recoverable from text shape at all. Nothing about "Background" tells you
  whether it is Heading 2 or Heading 4. Any level derived from ordering, indentation or font size
  is a second guess layered on the first.
* The set of things that are headings is not recoverable either. The reporter's own summary —
  "it only picks up headings that happen to be typed in ALL CAPS" — is the observation that the
  proxy signal and the real signal are uncorrelated.
* The repository's own standard says so: `CONTRIBUTING.md` lists **"outline levels"** as the first
  named example of "Structures the parsers don't read yet" for `src/doc/`, with the bar "Every
  claim traced to the MS-DOC … spec section it comes from". The project has already classified
  this as a missing-parser problem, not a tuning problem.
* The fix generalises beyond the reported symptom at no extra cost: `sprmPOutLvl` also gives
  correct levels to non-heading-styled paragraphs that carry an outline level (Word's
  "outline level" paragraph setting), which no text heuristic could ever find.

**Where a point patch *is* warranted**, and should be kept separate from the above: the
`level.min(6)` lower-clamp bug at `src/ir_render.rs:271` (§7.3) and the OOXML `outlineLvl == 9`
bug (§7.1). Those are genuine one-line defects with no general form.

### 4.3 What I am *not* recommending

* **Do not delete the text heuristic in the same change.** Doing so silently rewrites the output
  of every `.doc` that has no style information (§5.5), including `metadata.title`. Whether it
  survives as a fallback is a separate, arguable decision that deserves its own diff and its own
  corpus evidence.
* **Do not merge the resolved style's whole `grpprlPapx` into `PapProps`.** [MS-DOC] 2.6.2
  enumerates the properties that survive `sprmPIstd` (in-table, TTP, itap, inner-cell, table
  style, …) precisely because styles must not be allowed to disturb table structure (§5.6). Take
  the outline level from the style; take nothing else, for now.
* **Do not chase `sprmPHugePapx`/`sprmPTableProps` in this change.** It is a real gap (§5.7) but
  it is a different subsystem (the Data stream and `PrcData`) and mixing it in makes the diff
  unreviewable.

---

## 5. Predicted regressions

This is the section to read before writing any code. Assume "the obvious fix": plumb `istd`
through `FkpParagraph`/`PapProps`, decode `sprmPOutLvl`, compute `level = outlvl + 1`, and emit
`Element::Heading` from `emit_prose` when a level is present.

### 5.1 The naive fix turns every explicitly-body-text paragraph into an `######`

**Severity: highest. Affects: any document produced by Word that ever demoted a heading.**

`sprmPOutLvl` operand `0x9` means *body text*, not "level 9" ([MS-DOC] 2.6.2 **[specified]**).
Word writes it precisely when a paragraph must *cancel* an outline level it would otherwise
inherit — e.g. a paragraph whose style is based on a heading, or a "Body Text" style derived
from "Heading 4". A fix that computes `level = outlvl + 1` then clamps produces `9 + 1 = 10 → 6`,
i.e. `###### Some ordinary sentence.` on every such paragraph. Because the clamp hides the
out-of-range value, this will not look like a bug in code review; it will look like a document
full of H6s.

The same trap exists in reverse: a fix that treats `Option<u8>` presence as "is a heading" will
treat an explicit `sprmPOutLvl = 9` as a heading rather than as an explicit *negation*.

### 5.2 Numbered headings currently become bullet lists, and the fix will not reach them

**Severity: high. Affects: legal, standards, technical and government `.doc` — the exact corpus
that motivates this bug.**

`walk_paragraphs` (`src/convert_doc.rs:373-400`) dispatches in this order:

```
is_table_trailing_mark  ->  table.end_row(...)
f_in_table              ->  table.add_cell_paragraph(...)
is_doc_list_item(ilfo)  ->  list_items.push(...)          <- consumes the paragraph
else                    ->  emit_prose(...)               <- the only place headings are made
```

A Word "Heading 1"/"Heading 2" attached to a multilevel outline-numbered list (`1.`, `1.1`,
`1.1.1`) has a valid `ilfo` and is therefore consumed by the list branch at `convert_doc.rs:384`,
never reaching `emit_prose`. Today it renders as a bullet (`ordered = false`, see the comment at
`convert_doc.rs:384-389`). A fix confined to `emit_prose` leaves this untouched — the reporter will
still see "most of my headings aren't headings", now for a different reason.

The mirror-image regression is worse: **moving the heading check above the list check** turns
genuine list items into headings. Any list whose paragraphs are styled with a heading-ish style
(a numbered "Heading 3" body list, or a style that carries `outlineLvl`) loses its list structure
entirely — nested `Element::List` collapses into a run of `Element::Heading`, and
`flush_list`'s base-level logic (`convert_doc.rs:402-415`) stops seeing a contiguous run, so
even the surviving items are re-nested wrongly.

There is no ordering that is right for both. **[inferred]** Antiword settles it by fiat —
`xml.c` `vSetHeadersXML`: `if (bTableOpen || uiListLevel != 0) { /* No headers when you're in a
table or in a list */ return; }` (§3.4) **[specified]**. That loses numbered headings but never
destroys a list, and it is what a shipping reader does. Whatever is chosen, it must be an
explicit decision with a test, not a side effect of statement order.

### 5.3 Headings inside table cells corrupt the markdown table

**Severity: high. Affects: any `.doc` whose table cells use a heading style — extremely common in
forms and specification tables.**

Cell paragraphs go through `TableBuilder::add_cell_paragraph`
(`src/convert_doc.rs:105-120`), which unconditionally builds `Element::Paragraph`. If a fix
generalises heading emission to cells, `render_cell_markdown`
(`src/ir_render.rs:377-386`) special-cases only `Element::Paragraph`; everything else falls
through to `render_element_markdown`, which for a `Heading` returns `"### text"`. That string is
then interpolated between `|` pipes:

```
| ### Item | ...
```

which is not a heading and mangles the cell. Multi-paragraph cells are joined with `" "`, so the
`#` lands mid-cell. **[verified by reading `ir_render.rs:377-386` and `ir_render.rs:271-275`.]**

Keep the cell path emitting `Paragraph` (optionally bold), or teach `render_cell_markdown` to
flatten headings — either way it is a decision, not a default. Antiword's rule (§3.4) is simply
"no headers when you're in a table".

### 5.4 Empty heading-styled paragraph marks emit a bare `#`

**Severity: medium. Affects: nearly every real document.**

`emit_prose` returns early on empty text (`convert_doc.rs:448-451`), so today an empty paragraph
never becomes anything. Word documents are full of empty paragraph marks that inherit a heading
style — the paragraph after a heading before `istdNext` takes effect, blank separator lines a
user styled by accident, the final paragraph mark of the document. If the heading branch is added
before or beside the empty check, output gains lines that are literally `# ` with trailing
whitespace, which most markdown renderers show as an empty `<h1>`. Keep the empty-text guard
ahead of everything. **[inferred]**

### 5.5 Documents with no usable style information lose all headings — and their title

**Severity: high, and the most likely to be under-noticed.**

Three distinct populations get *nothing* from a style-based fix:

1. **No `PlcfBtePapx`.** `DocDocument::from_reader` (`src/doc/document.rs:100-108`) only builds
   structured paragraphs when `fib.fc_plcf_bte_papx != 0 && fib.lcb_plcf_bte_papx != 0`;
   otherwise `paragraphs` is empty and `doc_to_ir` falls to `line_heuristic`
   (`convert_doc.rs:14-19`), which has no paragraph properties at all, only text lines.
2. **`BxPap.bOffset == 0`.** [MS-DOC] 2.9.23 **[specified]**: a zero `bOffset` means there is no
   `PapxInFkp` and the paragraph takes default properties — istd is not even present. Same for
   `extract_grpprl`'s `cb < 3` early return (`papx.rs:167-169`).
3. **Direct-formatted headings.** A very large fraction of Word 97-era documents were written by
   people who never touched the style gallery: their "headings" are Normal paragraphs made bold,
   14pt and centred. `istd` is 0 for all of them. Nothing in the format marks them as headings.

Today all three populations still get *something*: the first line becomes an H1 and ALL-CAPS
lines become H2s. If the heuristic is deleted outright, these documents go from "some headings,
wrong levels" to "zero headings", which for population 3 is a real quality regression, and for
all three it silently drops `metadata.title` and `Section.title` (`convert_doc.rs:21-33` derive
both from the first `Element::Heading`; with no heading, both become `None`). Every consumer of
title metadata — the CLI `ir` command, the Python/Go/JS bindings, downstream `pdf_oxide` —
changes behaviour on those files.

The honest options are: (a) keep the heuristic strictly as a *fallback* used only when the
document yielded no style-derived heading at all; (b) keep it only for the `line_heuristic` path
where no properties exist; (c) delete it and accept the loss. Whichever is chosen must be stated,
because it changes output on a very large share of the corpus. **[inferred]**

### 5.6 Table-structure SPRMs must keep winning over the style

**Severity: medium-high. Affects: any table whose cells are heading-styled.**

**[specified]** [MS-DOC] 2.6.2, sprmPIstd: applying an istd MUST preserve, among others, "Whether
the paragraph is a Table Terminating Paragraph Mark (for example, by sprmPFTtp)", "Whether the
paragraph is in a table (for example, by sprmPFInTable)", and "The table depth of the paragraph
(for example, by sprmPItap)". If a future change resolves the style's `grpprlPapx` and merges it
into `PapProps` wholesale, a heading style that happens to contain a stale `sprmPFInTable` or
`sprmPItap` would corrupt table detection — turning prose into phantom table rows, or dropping
real rows. The style contribution must be restricted to the outline level, or explicitly filtered
per the preserved-properties list.

### 5.7 `sprmPHugePapx` paragraphs will look like `istd == 0`

**Severity: medium. Affects: heavily formatted paragraphs and wide tables.**

**[specified]** [MS-DOC] 2.6.2, sprmPHugePapx (0x6646): "If a Prl with a sprm of sprmPHugePapx is
contained in the grpprl array of a GrpPrlAndIstd structure, then it MUST be the only Prl in that
array and the **istd** member of that GrpPrlAndIstd structure MUST be zero." The real properties
live behind a `PrcData` in the Data stream. This repository does not follow that pointer, so such
a paragraph reads as `istd = 0` (Normal) with an empty grpprl. Today the ALL-CAPS heuristic can
still catch it; after the fix it becomes a plain paragraph. This is a *new* silent gap, not a
pre-existing one, from the reporter's point of view.

### 5.8 Malformed / hostile input — new panic and DoS surface

Everything below is new attack surface in a parser whose stated rule is "must return an error —
never panic, hang, or allocate unboundedly" (`CONTRIBUTING.md`, Robustness row).

| Structure | Hazard | Guard |
| --- | --- | --- |
| `GrpPrlAndIstd.istd` | reading `page[p+1..p+3]` when the FKP page is truncated | bounds-check; the existing `extract_grpprl` already returns `Vec::new()` for `cb < 3`, but the istd read happens *before* that check in the obvious implementation |
| `istd` value | spec says `0 <= istd < 0x0FFE` ([MS-DOC] 2.9.271) **[specified]**; hostile files exceed it | reject out-of-range instead of indexing `rglpstd` |
| `lcbStshf` / `cbStshi` | either can exceed the Table stream length, or be 0 | saturating slice, `checked_add`, and a "STSH unusable ⇒ no style info" path that still parses the document |
| `Stshif.cstd` | up to `0x0FFD` entries; a hostile file claims the max with a 20-byte stream | do not pre-allocate `cstd` entries; walk `LPStd`s bounded by the actual stream length |
| `LPStd.cbStd` | can be huge, or 0 (legally "empty style"), or overlap the next entry | clamp to remaining bytes; treat `cbStd == 0` as "no properties", never as an error |
| `StdfBase.istdBase` chain | spec says a loop MUST NOT occur; hostile files loop, or self-reference | depth cap — `src/docx/styles.rs:93` already uses `depth > 20` for exactly this on the DOCX side; mirror it |
| `Stshif.cbSTDBaseInFile` | MUST be `0x000A` or `0x0012`; hostile files say `0xFFFF` | validate against the two legal values, else give up on the STSH |
| `xstzName` | length-prefixed UTF-16, must be skipped to reach `grLPUpxSw` | length-checked skip; an over-long `cch` must not wrap |
| `LPUpxPapx.cbUpx` + `UPXPadding` | odd/even padding rules; a wrong `cbUpx` desynchronises the whole `grLPUpxSw` walk | bound every step by the enclosing `cbStd` |

`fuzz/fuzz_targets/fuzz_parse.rs` already feeds arbitrary bytes through
`DocumentFormat::Doc`, so a stylesheet reader will be fuzzed — but only if it is reached, which
requires the fuzzer to produce a plausible CFB + FIB. Corpus-seeding the fuzzer with the
synthetic builder's output would materially help here. **[inferred]**

### 5.9 The 0-based / 1-based boundary, and levels beyond 6

* File value is **0-based** (`sprmPOutLvl` 0..8) **[specified]**; the UI and the glossary speak of
  **Level 1..Level 9** **[specified]**; markdown wants **1..6**. Two conversions in a row is the
  classic place to be off by one. The `.docx` side of this repository already does
  `(level + 1).min(6)` (`src/convert_docx.rs:201`) — matching it is the only way the two formats
  agree on the same logical document, which is the reporter's actual complaint.
* **Do not** produce `Heading { level: 0 }`. `src/ir_render.rs:271` renders markdown as
  `"#".repeat(h.level.min(6))` with **no lower clamp**, so level 0 emits zero `#` and a leading
  space — the heading silently becomes body text with a stray space. (The HTML renderer at
  `ir_render.rs:440` does `clamp(1, 6)` and is safe; the two disagree.) **[verified by reading the
  code.]**
* Outline levels 6, 7, 8 (Heading 7/8/9) clamp to `######`. That is a lossy but defensible
  mapping. It means a document using Heading 6-9 loses hierarchy in markdown while keeping it in
  HTML — worth a note, not a blocker. **[inferred]** Antiword takes the other option and **drops**
  them (`if (usIstd == 0 || usIstd > 6) return;`, §3.4) **[specified]**; clamping is preferable
  here only because it matches what `convert_docx.rs:201` already does.

### 5.10 Existing tests that encode the current behaviour as intended

These will fail and must be *replaced*, not patched, because they assert the heuristic:

* `src/doc/document.rs:313-333` — `ir_allcaps_first_line_becomes_h1`,
  `ir_first_short_line_no_punct_becomes_h1`, `ir_allcaps_non_first_line_becomes_h2`.
* `src/doc/document.rs:335` — `ir_line_ending_with_period_becomes_paragraph`.
* `src/convert_doc.rs:555` — `soft_line_break_becomes_inline_break` deliberately ends its
  text with `'.'` "to keep this out of the heading heuristic". If the heuristic goes, the comment
  becomes a lie even though the test still passes.
* `go/office_oxide_test.go:58` asserts the rendered markdown `strings.Contains(md, "# ")`. If the
  Go fixture is a `.doc` without style information, this test starts failing on a change made in
  Rust. Check what document it uses before touching the fallback.

### 5.11 Corpus-diff blind spots

* Because *almost every* `.doc` currently gets a spurious H1 on its first line, this change alters
  output on essentially **every** `.doc` in a regression corpus. A whole-corpus diff will be
  enormous and therefore uninformative; it must be triaged into categories (gained a heading /
  lost a heading / changed level / changed title) rather than eyeballed.
* Conversely, the categories that matter most — heading styles nested in lists (§5.2), user styles
  with `sprmPOutLvl` (§2.4), `sprmPOutLvl == 9` (§5.1), `sprmPHugePapx` (§5.7), Word 6/95 — are
  each rare enough that a corpus of a few thousand mixed Office files may contain **zero**
  instances of some of them. A clean corpus diff is not evidence those paths work; it is evidence
  the corpus does not exercise them. Say which of those categories your corpus actually contains.

---

## 6. Testing strategy

The project rule (`CONTRIBUTING.md` "Unit Tests", `AGENTS.md` rule 4) is that in-tree reproducers
are **synthetic and built in code**, and the extra bar for `src/doc/` is *"A real file that
exercises the new path, not only a synthetic one"* (`CONTRIBUTING.md`, Where Help Is Wanted
table). Both apply here: synthetic tests go in the repo, the real-file evidence goes in the PR
description.

### 6.1 The builder needs two small extensions first

`tests/common/mod.rs` is the existing synthetic `.doc` writer. Two gaps block heading tests:

* `build_fkp_page` (`tests/common/mod.rs:192-215`) hard-codes the PAPX header as
  `vec![cw, 0, 0]` — **istd is always 0**. Add an `istd: u16` field to `Para` and write it into
  those two bytes. This alone unlocks the whole `istd ∈ 1..9` half of the feature.
* `write_fib` (`tests/common/mod.rs:151-161`) writes no `fcStshf`/`lcbStshf`. Add an optional
  stylesheet: a minimal `STSH` = `cbStshi` + a small `stshi` + `rglpstd` with a handful of
  `LPStd`s, one of which is a user style carrying `sprmPOutLvl` in its `UpxPapx`. Make the
  stylesheet *optional* so the "no style information" tests stay easy to write.

Both are additive and keep the existing table/list tests byte-identical.

### 6.2 Unit tests, named by defect class

In `src/doc/sprm.rs` (pure `grpprl` decode — no CFB needed):

| Test | Input | Expected |
| --- | --- | --- |
| `outline_level_sprm_is_decoded` | `[0x40, 0x26, 0x02]` | `outline_level == Some(2)` |
| `outline_level_nine_means_body_text` | `[0x40, 0x26, 0x09]` | resolves to *not a heading* |
| `direct_istd_sprm_overrides_fkp_istd` | `[0x00, 0x46, 0x03, 0x00]` | istd becomes 3 |
| `istd_one_to_nine_wins_over_outline_level_sprm` | istd = 3 **and** `sprmPOutLvl = 9` | heading level 3, per [MS-DOC] 2.6.2 |
| `truncated_outline_level_sprm_does_not_panic` | `[0x40, 0x26]` | empty result, no panic |

In `src/doc/papx.rs`:

| Test | Expected |
| --- | --- |
| `papx_istd_is_read_from_grpprl_and_istd` | both PAPX forms (`cw != 0` → `2·cw − 1`; `cw == 0` → `2·cb'`) yield the same istd |
| `papx_istd_read_is_bounds_checked` | a page truncated inside the istd bytes returns no properties, no panic |

In `src/convert_doc.rs` / `tests/`:

| Test | Expected |
| --- | --- |
| `heading_style_index_sets_heading_level` | istd 1/2/3 → `Heading{level:1/2/3}` |
| `heading_level_nine_clamps_to_markdown_six` | outline level 8 → markdown `######`, never 7+ hashes, never 0 |
| `body_text_outline_level_stays_a_paragraph` | `sprmPOutLvl = 9` → `Element::Paragraph` |
| `heading_styled_paragraph_without_text_emits_nothing` | empty text + istd 1 → no element (§5.4) |
| `heading_styled_cell_paragraph_stays_in_the_cell` | istd 2 + `fInTable` → cell content is not a `Heading`, and the rendered table row contains no `#` (§5.3) |
| `numbered_heading_is_not_swallowed_by_the_list_path` | istd 1 + valid `ilfo` → whichever behaviour you chose in §5.2, asserted explicitly |
| `document_without_style_information_still_yields_a_title` | no PAPX at all → asserts the fallback decision from §5.5 |
| `user_style_outline_level_resolves_through_the_stylesheet` | user istd whose `UpxPapx` carries `sprmPOutLvl = 1` → `Heading{level:2}` |
| `style_inheritance_loop_terminates` | `istdBase` pointing at itself / a 2-cycle → returns, no hang (§5.8) |
| `malformed_stylesheet_does_not_break_text_extraction` | `lcbStshf` past the end of the Table stream → text and tables still extract |

Every one of these is buildable from bytes in code; none needs a third-party document.

### 6.3 Renderer tests

Independent of the `.doc` change, because §5.9 is a latent IR bug:

* `markdown_heading_level_zero_is_not_emitted_as_body_text` — `Heading{level:0}` must not render
  as a bare space-prefixed line (`ir_render.rs:271`).
* Cross-check that `plain_text`, `markdown` and `html` all agree on the same `Heading`, per the
  "IR and renderers" extra bar in `CONTRIBUTING.md`.

### 6.4 Fuzzing

Add the synthetic builder's output (with and without a stylesheet) as seed inputs for
`fuzz/fuzz_targets/fuzz_parse.rs`. Without seeds, a byte-level fuzzer will essentially never
construct a valid enough CFB+FIB+Table stream to reach the STSH reader, so the new code would be
nominally fuzzed and actually untested. **[inferred]**

### 6.5 Real-file evidence (PR description, not the tree)

`CONTRIBUTING.md` demands it for `src/doc/`. Report, per category:

* count of `.doc` files in your corpus that contain at least one `istd ∈ 1..9` paragraph;
* count containing a `sprmPOutLvl` at all, and how many of those carry the value 9;
* count with a *user* style resolving to an outline level;
* count where a heading-styled paragraph also has a valid `ilfo` (§5.2) or `fInTable` (§5.3);
* before/after markdown for one document in each category.

If a category count is zero, say so — that is the honest statement that the branch is unproven,
which is exactly the bar the contributing guide sets.

---

## 7. Other things that look wrong (found while reading)

Each is separate from the reported defect. Not fixed, not filed — reported here only.

1. **`w:outlineLvl == 9` has the identical bug on the DOCX side.** **[verified]**
   `src/docx/formatting.rs:379-385` parses `w:outlineLvl` as a raw `u8`. OOXML defines 0-8 as
   levels 1-9 and **9 as "no outline level"** (ECMA-376 17.3.1.20; corroborated at
   <https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_outlineLvl_topic_ID0EKQCM.html> and
   <https://wordprocessingml.com/docs/headings/determining-if-a-paragraph-is-actually-a-heading/>).
   `src/convert_docx.rs:200-202` computes `(level + 1).min(6)` → **H6 for every explicitly
   body-text paragraph**, and `src/docx/text.rs:154-165` computes `(lvl as usize) + 1` with no
   guard. So the "much closer to right" comparison document has this defect too — fixing the
   `.doc` side without fixing this leaves the two formats disagreeing again.
2. **`src/docx/text.rs:189` emits up to nine `#`.** `"#".repeat(level.min(9))` — markdown has six
   levels; `####### x` renders as literal text in every CommonMark implementation. The IR path
   (`convert_docx.rs:201`) clamps to 6 but this direct text path does not, so `Document::markdown()`
   and the IR renderer can disagree on the same file. **[verified]**
3. **`ir_render.rs` markdown does not lower-clamp the heading level.** `"#".repeat(h.level.min(6))`
   at `src/ir_render.rs:271` vs `h.level.clamp(1, 6)` at `src/ir_render.rs:440` (HTML). A level-0
   heading silently becomes a body line in markdown and an `<h1>` in HTML. `Heading::level` is
   documented as "1–6" (`src/ir.rs:642`) but nothing enforces it and `serde` will happily
   deserialise 0. **[verified]**
4. **`Pcd.Prm` is parsed and dropped.** `src/doc/piece_table.rs:105` — [MS-DOC] 2.4.6.1 step 5
   **[specified]** says `Prm0`/`Prm1` append paragraph `Prl`s. Any property carried there (style
   changes in `fComplex` documents are the classic case) is lost. **[verified that the code drops
   it; inferred that real documents rely on it.]**
5. **`Fib::parse` accepts Word 6/95 and then reads Word 97 offsets.** `src/doc/fib.rs:58-61`
   allows `wIdent == 0xA5DC`, then `fib.rs:88-105` reads `fcClx` at `0x01A2` and `fcPlcfBtePapx`
   at `0x0102` — absolute offsets valid for `nFib == 0x00C1`. Word 6/95 also uses **1-byte** SPRM
   opcodes, which `parse_grpprl` (`src/doc/sprm.rs:117`) never handles (it always reads a 2-byte
   opcode). Such files will either bail out via the `Err` paths in `document.rs` or walk garbage
   SPRMs. `version` is stored on `Fib` but never branched on. **[verified by reading; the WW6
   SPRM difference is corroborated by the internal SPRM reference's `mnDelta` note but I did not
   confirm it against [MS-DOC], which does not cover Word 6/95.]**
6. **`tests/common/mod.rs:179-188` writes CPs where the `PlcfBtePapx` wants FCs.**
   `build_plcf_bte_papx` fills `rgfc` from `cp_starts`, but [MS-DOC] 2.8.6 **[specified]** says
   `PlcBtePapx` is indexed by **FC**. It happens not to matter because `parse_papx_paragraphs`
   uses the `rgfc` only to size the array and takes the real ranges from the FKP pages
   themselves — but the builder is therefore producing a document Word would reject, which
   weakens every test built on it. **[inferred]**
7. **Nested tables are flattened.** Already documented as a known limitation at
   `src/convert_doc.rs:46-51`; noted only because §5.6's "preserve table properties across
   sprmPIstd" interacts with it. **[verified — it is an acknowledged gap, not a new finding.]**
