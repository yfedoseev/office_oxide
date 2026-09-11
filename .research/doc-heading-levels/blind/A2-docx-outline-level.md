# A2 — DOCX outline level: root-cause research

Baseline: `/home/yfedoseev/projects/office_oxide` @ `main` (`5007d9c`), read-only. No build was run.

Every correctness claim below is tagged **[specified]** (read in the standard's text, in an
implementation note, or in a real implementation's source) or **[inferred]** (reasoned from code
I read, but not executed and not stated by a primary source).

**Disclosure:** while grepping the tree I hit `./.claude/worktrees/agent-a15412f1b9130f80a/`,
a checkout of this crate whose `src/doc/document.rs` contained a test name
(`ir_deep_outline_level_clamps_to_ir_max_depth`) that suggests somebody's in-flight fix for
this exact area. I stopped reading that directory immediately, used nothing from it, and
excluded it from all subsequent searches. Everything below is derived from `main` only.

---

## Summary of what the code actually does

One data path, two consumers:

```
w:pPr/w:outlineLvl ──> ParagraphProperties.outline_level : Option<u8>   (src/docx/formatting.rs)
w:styles.xml       ──> StyleSheet::resolve_outline_level  : Option<u8>   (src/docx/styles.rs:89)
                                   │
              ┌────────────────────┴─────────────────────┐
              │                                          │
   DocxDocument::to_markdown()                  docx_to_ir() ─> DocumentIR::to_markdown()
   src/docx/text.rs:155-194                     src/convert_docx.rs:434-447, :199-205
   heading iff Some(lvl); "#" x min(lvl+1, 9)   heading iff Some(lvl); level = min(lvl+1, 6)
                                                then src/ir_render.rs:272 "#" x min(level, 6)
```

Both consumers share the same two mistakes about the value space, and disagree about everything
else. That is the whole story of the three symptoms.

---

# Symptom 1 — body-text paragraphs come out as headings

> "In the source XML those paragraphs do have an explicit outline level set on them, but Word
> definitely doesn't show them in the navigation pane as headings."

## Root cause

The parser accepts *any* `u8` for `w:outlineLvl/@w:val` and both consumers treat "a value is
present" as "this paragraph is a heading". Nothing in the crate knows that `9` means *no outline
level*.

Parse sites — four copies, none validating the range (`src/docx/formatting.rs`):

* `:379-385` (namespace-resolving reader, `Event::Start`)
* `:417-422` (namespace-resolving reader, `Event::Empty`)
* `:604-610` (fast reader, `Event::Start`)
* `:685-691` (fast reader, `Event::Empty`)

each of which is literally:

```rust
b"outlineLvl" => {
    if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
        if let Ok(lvl) = val.parse::<u8>() {
            props.outline_level = Some(lvl);
        }
    }
}
```

`parse::<u8>()` happens to reject `-1` and `> 255`, but accepts `9` … `255`.

Consumers:

* `src/docx/text.rs:155-165` — `pp.outline_level.or_else(|| … resolve_outline_level(sid))`,
  then `.map(|lvl| (lvl as usize) + 1)`. Any `Some(_)` becomes a heading.
* `src/convert_docx.rs:434-447` (`resolve_heading_level`) + `:199-205` — same rule, feeding
  `Element::Heading`.
* `src/docx/styles.rs:89-105` (`resolve_outline_level`) walks the `w:basedOn` chain and returns
  the first level it finds — correct *shape*, but because `9` is not understood as "body", a
  style that deliberately sets `9` to *cancel* an inherited heading level (the standard idiom)
  is read as "level 10".

The two highest-yield real-world instances of this **[inferred, from the code + the style idiom]**:

1. **`TOCHeading`.** Word's TOC-heading style is `basedOn` `Heading1` and sets
   `<w:outlineLvl w:val="9"/>` precisely so the words "Table of Contents" do not appear inside
   the TOC it heads. Third-party evidence of that exact shape:
   [docx4java forum](https://www.docx4java.org/forums/docx-java-f6/genearting-toc-without-skippagenumbers-not-working-t3036.html)
   ("The user provided their custom TOCHeading style definition … sets `outlineLvl` to 9").
   Under `main`, such a paragraph becomes heading level 10.
2. **Paragraph-level "Body Text" outline level.** When a user sets *Outline level: Body Text* in
   Word's Paragraph dialog on a paragraph whose style carries a level, Word records
   `<w:outlineLvl w:val="9"/>` as direct formatting. Same result.

Both cases match the reporter's description exactly: an explicit outline level in the XML, and
Word not showing the paragraph in the navigation pane — because `9` *is* Word's encoding for
"not in the navigation pane".

There is a second, independent way this crate can *manufacture* such files — see
**Cross-cutting finding C3** (the writer's inverted `outline_level` contract).

## What the standard says

**[specified]** ECMA-376 Part 1 / ISO-IEC 29500-1, **§17.3.1.20 `outlineLvl` (Associated Outline
Level)**, quoted verbatim on Microsoft Learn (Open XML SDK reference reproduces the clause text
under `© ISO/IEC29500: 2008`):

> "This element specifies the *outline level* which shall be associated with the current
> paragraph in the document. … This level shall not affect the appearance of the text in the
> document, but shall be used to calculate the *TOC* field (§17.16.5.68) if the appropriate field
> switches have been set, and can be used by consumers to provide additional application
> behavior.
>
> The outline level of text in the document (specified using the val attribute) can be from *0*
> to *9*, where *9* specifically indicates that there is no outline level specifically applied to
> this paragraph. If this element is omitted, then the outline level of the content is assumed to
> be *9* (no level)."

— https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc882417(v=office.14)
(also at https://learn.microsoft.com/dotnet/api/documentformat.openxml.wordprocessing.outlinelevel).

Answers to the questions posed, from that text and its neighbours:

* **Permitted value range:** prose says 0–9 **[specified]**. The *schema* does not enforce it:
  the content model is `CT_DecimalNumber` and `@w:val` is `ST_DecimalNumber` (§17.18.10), a
  restriction of `xsd:integer` **[specified]**; MS-OE376 Part 4 §2.18.16 notes Word narrows that
  to `xsd:int`
  (https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/e8c9b787-495c-4f04-9862-523e8db56a04)
  **[specified]**. So `10`, `-1`, `99999` are schema-valid and prose-invalid; the standard does
  **not** say what a consumer must do with them (silent).
* **Is every value a heading?** No. `9` means *no outline level*, i.e. body text, and `9` is also
  the default when the element is absent **[specified]**. `0`–`8` are outline levels 1–9.
* **Relationship to the built-in `Heading N` styles:** the standard establishes no normative
  identity. The convention — `Heading N` style's `w:pPr` carries `w:outlineLvl = N-1` — is visible
  in Microsoft's own sample `styles.xml` (`Heading1` → `<w:outlineLvl w:val="0"/>`,
  https://learn.microsoft.com/en-us/dotnet/standard/linq/style-part-wordprocessingml-document)
  **[specified]**. Microsoft's implementation notes make the direction of dependence explicit:
  MS-OI29500, *Part 1 Section 17.16.5.68, TOC*, note (c): "The standard states that the \o switch
  will direct the TOC field to include paragraphs formatted with the specified built-in heading
  style(s). — **Word will include any style whose definition includes an `outlineLvl` ([ISO/IEC-29500-1]
  §17.3.1.20; outlineLvl) paragraph property corresponding to the heading level(s) specified by
  the \o field argument.**"
  (https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/68f1577f-0efe-453b-bad9-cfd69f740b30)
  **[specified]** — i.e. Word keys off the outline level, not the style *name*.
* **Direct `outlineLvl` vs. a style that sets one — which wins?** Direct formatting.
  §17.3.1.27 `pStyle` states the style hierarchy explicitly: "Document defaults / Table styles /
  Numbering styles / Paragraph styles (this element) / Character styles / **Direct Formatting**",
  with the worked example showing direct `w:ind` overriding the style's
  (https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc866133(v=office.14))
  **[specified]**. The crate already gets this right (`text.rs:159`, `convert_docx.rs:440`
  prefer the paragraph's own value) — precedence is *not* the bug.
* **What should happen to out-of-range values?** The standard is **silent** **[specified — by
  absence]**. LibreOffice's importer documents Word's actual behaviour in a code comment:
  "invalid value is ignored by MS Word" (see below).
* **Navigation pane:** not a format question at all; the standard says the level "shall not affect
  the appearance of the text" and is for the TOC field plus "additional application behavior"
  **[specified]**. Secondary sources on Word's UI conflict: Office Watch says "A Table of Contents
  can be built from Outline Levels while the Navigation Pane shows Headings"
  (https://office-watch.com/2025/word-headings-vs-outline-levels/), whereas Charles Kenyon's
  reference states for Word 2010+ "Word displays text in the Navigation Pane based entirely on the
  Outline Level of the paragraph. It does not guess."
  (https://www.addbalance.com/usersguide/navigationPane.htm). I could not resolve this from a
  primary source and did not run Word. Note that under *either* reading, a paragraph whose
  effective level is `9` is absent from the navigation pane — which is what the reporter observed.

## What other tools do

* **LibreOffice Writer (DOCX import)** — `sw/source/writerfilter/dmapper/DomainMapper.cxx:2140-2157`
  **[specified — source read]**:

  ```cpp
  case NS_ooxml::LN_CT_PPrBase_outlineLvl:
      if (nIntValue < WW_OUTLINE_MIN || nIntValue > WW_OUTLINE_MAX)
          break; // invalid value is ignored by MS Word
      …
      // convert MS body level (9) to LO body level (0) and equivalent outline levels
      sal_Int16 nLvl = nIntValue == WW_OUTLINE_MAX ? 0 : nIntValue + 1;
      rContext->Insert(PROP_OUTLINE_LEVEL, uno::Any(nLvl));
  ```

  with `WW_OUTLINE_MIN = 0`, `WW_OUTLINE_MAX = 9` (`dmapper/PropertyMap.hxx:583-584`). The same
  9→body mapping is applied to *style* definitions in `StyleSheetTable.cxx:1247-1258`, and
  `StyleSheetTable.cxx:946-956` inherits a missing level from the parent style — i.e. LibreOffice
  does exactly what `styles.rs:89` does, *plus* the 9→body rule and the out-of-range rule.
  (Source: https://raw.githubusercontent.com/LibreOffice/core/master/sw/source/writerfilter/dmapper/DomainMapper.cxx)
* **Pandoc (docx reader)** — never reads `w:outlineLvl` at all. Heading detection is purely by
  style *name*: `getHeaderLevel` in `src/Text/Pandoc/Readers/Docx/Parse/Styles.hs:297-304`
  strips a case-insensitive `"heading "` prefix off `w:name`/`w:styleId` and requires `n > 0`;
  `pHeading` (`…/Docx/Parse.hs:837-838`) resolves it through the `basedOn` chain via
  `getParStyleField`. **[specified — source read]** Consequence: pandoc never mis-promotes
  `outlineLvl 9`, but it also misses genuinely-structural custom styles.
* **mammoth.js** (the engine behind Microsoft's `markitdown`) — style-map only, levels 1–6:
  `"p.Heading1 => h1:fresh"`, `"p[style-name='Heading 1'] => h1:fresh"`, … `Heading6`
  (`lib/options-reader.js:7-24`). No `outlineLvl`. **[specified — source read]**
* **Apache POI XWPF** — exposes no outline level at all; `XWPFParagraph.getStyle()`
  (`XWPFParagraph.java:1452-1459`) returns the `w:pStyle` id and that is the only hook POI-based
  converters have. **[specified — source read]**
* **Word itself** — includes in a `\o` TOC any style whose definition carries a matching
  `outlineLvl` (MS-OI29500 §17.16.5.68 note c, quoted above) **[specified]**.

So: of the tools examined, the only two that read `outlineLvl` are Word and LibreOffice, and both
treat `9` as "not a heading"; nobody treats `9` as a heading.

## Recommended fix — **generic**

Make the value space explicit at the parse boundary, and give the two consumers one shared
decision function.

1. In `src/docx/formatting.rs`, parse into a validated representation instead of a raw `u8`.
   Concretely: keep the field, but normalise on the way in — reject anything outside `0..=9`
   (Word ignores it; LibreOffice ignores it; the standard is silent so following Word is the
   defensible choice), and keep `9` as a *stored, meaningful* value rather than dropping it, so
   the writer can round-trip it. A newtype (`OutlineLevel(u8)` with a private constructor and a
   `heading_level(self) -> Option<NonZeroU8>` accessor returning `None` for `9`) is the shape
   that makes the class of bug unrepresentable. The four parse sites should call one helper —
   they are already four verbatim copies, which is how a range check gets added in three places
   and forgotten in the fourth.
2. `StyleSheet::resolve_outline_level` keeps returning the *first* value found up the `basedOn`
   chain — that is correct and must stay, because an explicit `9` in a derived style is how a
   style cancels an inherited heading level. It becomes correct automatically once `9` is
   understood downstream.
3. Exactly one function decides "is this paragraph a heading, and at what level" — used by both
   `text.rs` and `convert_docx.rs` (and ideally only by `convert_docx.rs`; see Symptom 3).

**Why not the point patch** (`if lvl < 9` at `text.rs:159` and `convert_docx.rs:440`): it leaves
four unvalidated parse sites that will keep admitting `10..=255`; it leaves the writer able to emit
values the reader would then mis-read (C3); and it adds a *second* pair of duplicated rules to a
codebase whose third reported symptom is precisely that this pair has already drifted apart.
The defect class here is "a domain value with a sentinel is carried as a bare integer", and the
generic fix is to stop doing that.

---

# Symptom 2 — more than six `#`

## Root cause

`src/docx/text.rs:189`:

```rust
let hashes = "#".repeat(level.min(9));
```

`level` is `outline_level + 1` (`text.rs:165`), so the direct path emits **up to nine** `#`.
Reachable two ways:

* legitimately, from `Heading 7`/`8`/`9` (outline levels 6/7/8) → `#######`…`#########`;
* spuriously, from the Symptom-1 bug: `outlineLvl 9` → level 10 → clamped to 9 → `#########`.

The IR path does not have this bug (`convert_docx.rs:201` `(level + 1).min(6)`;
`ir_render.rs:272` `"#".repeat(h.level.min(6))`), which is why the symptom appears only on one of
the two outputs — and is why Symptom 2 and Symptom 3 are partly the same defect.

Two adjacent weaknesses in the same area **[inferred, code-read]**:

* `ir::Heading.level` is documented "Heading level 1–6" (`src/ir.rs:641-644`) but nothing
  enforces it. It is `#[serde(default = "default_heading_level")]` (`ir.rs:659-661`) and public,
  so IR arriving from JSON (MCP `ir` output round-trip, Python/WASM bindings, pdf_oxide) can carry
  `0`, `7`, `99`. `ir_render.rs:272` uses `min(6)` with no lower bound → `level: 0` renders a line
  that begins with a space and no `#` at all. The HTML renderer instead uses
  `clamp(1, 6)` (`ir_render.rs:441`) and the DOCX writer `clamp(1, 6)` (`write.rs:1253`,
  `write.rs:525`): three renderers, three different normalisation rules, none of them the type.
* `(level + 1)` at `convert_docx.rs:201` is `u8` arithmetic on a value the parser lets reach
  `255`: `<w:outlineLvl w:val="255"/>` overflows — panic in a debug build, wrap-to-`0` in release
  (`Cargo.toml:155-159` sets no `overflow-checks`). See Other findings O1.

## What the standard says

**[specified]** ECMA-376 says nothing about markdown; it caps the value space at 9 levels
(§17.3.1.20) and no more. The markdown side is governed by CommonMark: an ATX heading is 1–6
`#` characters; seven or more is not a heading. GitHub therefore renders `#######` literally,
as the reporter observed. So a level-7..9 outline level has **no** faithful ATX representation
and the choice of degradation is a product decision the format does not make for us
**[specified — by absence]**.

## What other tools do

* **Pandoc**: its Markdown writer does *not* clamp — `literal (T.replicate level "#")`
  (`src/Text/Pandoc/Writers/Markdown.hs:573`) **[specified — source read]** — but its docx reader
  can only produce levels from `heading N` style names, so >6 is rare rather than impossible.
  Its **HTML** writer does degrade: `let classes' = if level > 6 then "heading":classes else classes`
  and `case level of 1 -> h1 … 6 -> h6; _ -> H.p` (`src/Text/Pandoc/Writers/HTML.hs:1025-1048`)
  **[specified — source read]** — i.e. pandoc's considered answer for "deeper than the target
  format supports" is *demote to a paragraph carrying a marker class*, not "emit invalid markup".
* **mammoth.js**: only maps Heading 1–6; a `Heading 7` paragraph falls through to `<p>`
  **[specified — source read]**.
* **LibreOffice**: has ten levels internally (body + 9), so the question doesn't arise on import.

## Recommended fix — **generic**

Enforce the `Heading.level` invariant once, in the type, rather than clamping at each renderer:

* give `ir::Heading` a constructor that clamps into `1..=6` (and a `serde` `deserialize_with`
  that does the same), so `convert_docx.rs:201`, `convert_doc`, `convert_pptx`, `convert_xlsx`,
  the JSON deserialiser and any binding all land in range by construction;
* then delete the ad-hoc `min(6)` / `clamp(1,6)` at `ir_render.rs:272`, `ir_render.rs:441`,
  `write.rs:1253`, `write.rs:525` — or keep them as cheap belt-and-braces, but they must all agree.

`min(9)` at `text.rs:189` should not be "fixed to `min(6)`" as a point patch — the right move is
that the direct path stops existing (Symptom 3). If it must survive an interim release, it must
call the same normalisation as the IR path, not a hand-written `min`.

Policy recommendation for levels 7–9: clamp to 6. It matches what the crate's own DOCX writer
can express (only `Heading1`…`Heading6` styles are generated, `write.rs:3157-3162`) and what
mammoth does. Pandoc's "demote to a paragraph" is the more information-preserving option but
changes far more output and loses the heading nature entirely; clamping keeps the paragraph
navigable. Either way, decide once and encode it in the type, not per renderer.

---

# Symptom 3 — `to_markdown()` ≠ `to_ir().to_markdown()`

## Root cause

Not a drifted pair of constants: **two independent markdown implementations**.

* `DocxDocument::to_markdown()` → `src/docx/text.rs:32-58` → `markdown_blocks` (`text.rs:150-233`).
* `Document::to_markdown()` dispatches to that same DOCX-native path (`src/lib.rs:337-339`,
  `src/docx/mod.rs:1465-1467`), while `Document::to_ir()` (`lib.rs:347`) goes through
  `convert_docx::docx_to_ir` and then `DocumentIR::to_markdown` (`ir_render.rs:151-159`).

Every difference I found by reading both, on heading-bearing input:

| # | Input shape | `to_markdown()` (text.rs) | `to_ir().to_markdown()` | Where |
|---|---|---|---|---|
| 1 | `outlineLvl` 6–9 | 7–9 `#` | 6 `#` | `text.rs:189` vs `convert_docx.rs:201` |
| 2 | first heading of a section | rendered once | rendered **twice**, once as `## <text>` | `ir_render.rs:253-258` + `convert_docx.rs:70-84,111-113` |
| 3 | heading paragraph that also has `w:numPr` | heading wins → `# Foo` | list wins → `- Foo` | `text.rs:188-194` vs `convert_docx.rs:167-176` (list branch runs first, `continue`s) |
| 4 | headers/footers | included, deduped | **dropped** — `Section.header`/`footer` are populated (`convert_docx.rs:25-59`) but `render_section_markdown` never reads them | `text.rs:38-52` vs `ir_render.rs:253-267` |
| 5 | heading inside a table cell | flattened to plain text | rendered as `## Foo` **inside the pipe cell** | `text.rs:308-327` vs `ir_render.rs:377-386` |
| 6 | empty paragraph with bottom border | plain empty line | `ThematicBreak` → `---` | `convert_docx.rs:181-196` only |
| 7 | page break inside a paragraph | `\n\n---\n\n` inline in the run text | paragraph split + `ThematicBreak` element | `text.rs:257-260` vs `convert_docx.rs:207-223` |

**[all verified by code reading; none executed]**

Item 2 deserves emphasis: `Section.title` is set to the text of the section's *first* `Heading`
(`convert_docx.rs:70-84`) and `render_section_markdown` unconditionally prints it as a level-2
heading before the elements. So for a document whose first paragraph is `outlineLvl 0`, the IR
markdown is `## Report Title\n\n# Report Title\n\n…`. The same double-emission exists for PPTX
(`convert_pptx.rs:32-38` pushes a `level: 2` heading *and* `:76` sets `title`) — so this is a
cross-format defect in the IR renderer, not a DOCX one.

The in-tree tests cannot see any of this: `tests/office_integration.rs:812-838`
(`ir_markdown_with_formatting`) asserts `ir_md.contains("## Section Header")` — which is satisfied
by the duplicate *and* by the real heading, and `tests/office_integration.rs:787-806` asserts
`contains` on plain text likewise. Substring assertions are structurally blind to duplication and
to divergence between the two paths.

## What the standard says

Nothing — this is an internal API-consistency defect, not a format question **[specified — by
absence]**. The only standard-adjacent observation: since ECMA-376 defines exactly one meaning for
a document's outline structure (§17.3.1.20 + the §17.3.1.27 hierarchy), a library that produces two
different structures for one file is wrong at least once, no matter which output the user prefers.

## Recommended fix — **generic / structural**

Delete the second implementation. `DocxDocument::to_markdown()` should become
`convert_docx::docx_to_ir(self).to_markdown()`, and `markdown_blocks`/`markdown_run`/
`markdown_table`/`split_headers_footers` in `src/docx/text.rs` should go with it. `plain_text`
stays (it is a genuinely different, cheap product and has no heading logic).

Prerequisites, all of which are bugs in their own right and are worth fixing regardless:

* `ir_render::render_section_markdown` must render `Section.header`/`Section.footer` (today's
  regression #4 above), and must **stop** emitting `Section.title` as a synthetic `##` — the title
  is metadata, already present as an element (#2).
* `convert_docx` must decide heading-vs-list deliberately rather than by branch order (#3).
  Word's own model allows a numbered heading; the IR has no representation for one, which is the
  real gap. Minimum viable: check the heading test before the list test so a numbered heading
  stays a heading, matching the direct path and Word's navigation pane.
* Table cells must not emit `#` inside a pipe row (#5): `render_cell_markdown` should render a
  `Heading` as inline text (or as `**bold**`), the way it already special-cases `Paragraph`.

Argument against the matched-pair point patch (fix `min(9)`→`min(6)` in `text.rs` and call it
done): it addresses one of seven observed divergences, and the mechanism that produced all seven —
two hand-maintained renderers over the same data — is untouched, so the next feature added to one
renderer re-opens the class. The project's stated standard ("fixes classes of problems, not
instances", CONTRIBUTING §"The rules that matter most") points the same way. The cost is honest
and should be stated in the PR: **the DOCX markdown output changes for essentially every
document** (see regressions), so it is a minor-version, CHANGELOG-worthy behaviour change, and it
is the only fix that makes symptom 3 impossible rather than currently-absent.

If the maintainer judges that too large for one change, the defensible intermediate is: extract a
single `docx::heading_level(props, styles) -> Option<u8>` used by both paths *now* (kills symptoms
1 and 2 and divergence #1 in one edit), and land the path unification as a tracked follow-up.
That is still generic with respect to the outline-level defect class; it is a point patch with
respect to symptom 3, and should be labelled as such.

---

# Cross-cutting findings

**C1 — the three symptoms are two defects, not three.** Symptoms 1 and 2 are both "the value
space of `w:outlineLvl` is not modelled": `9` is a sentinel that isn't handled, and `0..=8`+1
exceeds the target format's range. Symptom 3 is a separate, structural defect (duplicated
renderer) that *amplifies* both — the same file yields `#########`, `######` or nothing depending
on which entry point you call. A `TOCHeading` paragraph triggers all three at once.

**C2 — four copies of the `outlineLvl` parse.** `formatting.rs:379/417/604/685`. Any per-site fix
is a 1-in-4 chance of being incomplete. The Start/Empty × resolving/fast matrix is a general
hazard in this file (the same duplication exists for `pStyle`, `jc`, `ind`, `spacing`), and note
that the *fast* variants additionally handle `framePr`, `sectPr` and `pBdr` that the resolving
variants do not (`formatting.rs:620-660` vs `:350-395`) — the two readers already disagree about
what a `w:pPr` contains **[verified]**.

**C3 — the writer's `outline_level` contract is inverted relative to the standard, and it is a
plausible source of the very files in symptom 1.** `src/docx/write.rs:264-265` and
`src/ir.rs:743-744` both document `outline_level` as "**0 = body text, 1–9 = heading levels**".
Per §17.3.1.20 that is wrong twice: `0` is Heading 1 and `9` is body text **[specified]**. The
value is written to the file verbatim (`write.rs:1604-1608`) and passed through from
`ir::Paragraph.outline_level` (`create.rs:218`, `write.rs:1378`). `convert_docx` never populates
`ir::Paragraph.outline_level`, so the only producers are external — e.g. pdf_oxide, which this
crate's own memory notes as a downstream consumer. A producer that follows the doc comment and
writes `Some(0)` for body text emits `<w:outlineLvl w:val="0"/>` on every body paragraph; Word
then lists them all as Heading 1 in the navigation pane and TOC, and *this crate reads them back
as `# headings`* **[inferred — I did not read pdf_oxide; the contradiction between
`formatting.rs:44` ("0 = Heading 1") and `ir.rs:743`/`write.rs:264` ("0 = body text") inside one
crate is verified]**. Whatever is done for symptoms 1–3, these two doc comments must be
reconciled to the standard, and the writer should validate `0..=9` on the way out.

**C4 — the project's own spec note is incomplete.** `docs/specs/docx_spec.md:605`:
`| w:outlineLvl | w:val (0-9) | Outline level (0 = Heading 1). |` — correct as far as it goes,
but it omits the one fact that causes symptom 1 (`9` = no level, and `9` is the default when
absent). Whoever wrote the consumer code was reading a spec note that didn't warn them
**[verified]**.

---

# Predicted regressions

## Fix 1 (treat 9 as body; ignore out-of-range)

* **Which documents change:** any with `<w:outlineLvl w:val="9"/>` on a paragraph or in a style —
  in practice, documents containing a Word-generated table of contents (`TOCHeading`), documents
  where an author explicitly set *Outline level: Body Text*, and documents from generators that
  emit `9` as "none". Also anything with a value ≥10 (rare; malformed or generator bugs).
* **Direction:** strictly *fewer* headings. `######### Table of Contents` (direct path) and
  `###### Table of Contents` (IR path) both become body text.
* **Silent ripple to watch:** `Section.title` and `DocumentIR.metadata.title` are derived from the
  *first* `Heading` in the section/document (`convert_docx.rs:70-88`). If a document's first
  false heading disappears, the document title changes — and that title flows into the MCP `ir`
  tool output, the JSON bindings, and PDF/HTML rendering. A consumer keying on `metadata.title`
  sees a different value for the same file. **[inferred]**
* **Would a golden-output suite catch it?** Only if the corpus contains a file with an
  `outlineLvl 9` anywhere. A corpus of hand-made or generator-made DOCX without a Word TOC will
  show **zero diff** and give false confidence. This is the failure mode to call out in the PR.

## Fix 2 (clamp emitted heading depth to 6)

* **Which documents change:** those using `Heading 7/8/9` or outline levels 6–8 — deep legal,
  standards and government documents; also every document already hit by Fix 1 if Fix 1 is not
  applied first (`9`→10→9 hashes).
* **Direction:** `#######`…`#########` → `######`. Levels 6–9 become indistinguishable in the
  output; any downstream consumer building a tree from hash counts sees a flattened tail.
  If instead the "demote to paragraph" policy were chosen (pandoc's HTML behaviour), those lines
  would lose heading status entirely — a larger change, and one that breaks TOC extraction.
* **Would a golden-output suite catch it?** Only with a Heading-7+ document in the corpus. Rare.
  A cheap alternative check that *is* corpus-independent: assert that no produced line matches
  `^#{7,}` — that catches the class on whatever files the corpus happens to hold.

## Fix 3 (unify the two markdown paths)

* **Which documents change:** effectively **all** DOCX. Beyond headings: header/footer text
  appears/disappears depending on direction, list markers change (the direct path renders every
  ordered item as `level.start`, i.e. always the same number — `text.rs:176-180`), blank-line
  placement changes (the direct path emits an extra newline after every heading, `text.rs:224-226`;
  `ir_render` joins blocks with `\n\n`), section separators change (`ir_render.rs:159` joins
  sections with `\n\n---\n\n`; the direct path has no section concept), and page breaks change
  shape.
* **Direction:** toward the richer IR rendering; mostly additive (headers/footers, thematic
  breaks, real list numbering), but whitespace-different everywhere.
* **Would a golden-output suite catch it?** It would flag nearly every file, which makes it
  useless as a *safety* check for this change — the diff is the point. The honest test is not
  "output unchanged" but "the two paths now agree", plus human review of a handful of goldens.
* **Interaction:** Fix 3 must land with the `Section.title` double-emission fix, or unification
  makes output *worse* (every DOCX gains a duplicated `##` first heading).

---

# Testing strategy

The project's constraints (CONTRIBUTING §"The rules that matter most": synthetic in-code
reproducers; no third-party fixtures; the real corpus is private and not distributed) mean an
in-tree diff suite can only catch what its own synthetic files exercise. So the tests have to be
written as *properties over a value space*, not as goldens.

1. **Value-space table test** (unit, `src/docx/formatting.rs` or `tests/docx_integration.rs`,
   using the existing `make_minimal_docx` helper in `tests/office_integration.rs`). One synthetic
   `document.xml` per case, asserting the resulting heading level (or absence):
   `w:val` = `0`, `1`, `5`, `6`, `8` → headings 1,2,6,7→6,9→6; `9` → **not** a heading;
   `10`, `255`, `-1`, `abc`, missing `@w:val`, empty string → not a heading, no panic.
   Name it for the class, e.g. `outline_level_value_space_is_honoured`, per the CONTRIBUTING rule
   about naming by defect class.
2. **Style-resolution table test** (`tests/docx_integration.rs`, next to the existing
   `resolve_outline_level` test at `:158`): `Heading1` (level 0) → H1; `TOCHeading`
   `basedOn=Heading1` + `outlineLvl 9` → body; three-deep `basedOn` chain where only the root sets
   a level → inherited; a `basedOn` cycle → terminates, no heading; direct `outlineLvl 9` on a
   paragraph styled `Heading1` → body (this is the §17.3.1.27 precedence case); direct
   `outlineLvl 0` on a `Normal` paragraph → H1.
3. **Cross-path equivalence property** — the test that makes symptom 3 non-recurring. Over the
   whole battery of synthetic bodies from (1) and (2):
   `assert_eq!(doc.to_markdown(), doc.to_ir().to_markdown())`. If full string equality is too
   strong before unification lands, assert equality of the extracted heading sequence
   `Vec<(usize /*hashes*/, String /*text*/)>` parsed from both outputs. Once the paths are
   unified this test costs nothing and pins the property forever.
4. **Markdown-validity property**, applied to every markdown-producing test in the suite (a small
   helper in `tests/common`): no output line matches `^#{7,}`; no output line is a bare `#` or
   `# ` with empty content. Corpus-independent, catches the class rather than the instance.
5. **Round-trip property**: build a DOCX with `add_heading(level)` for 1..=6, re-read, assert the
   level survives; and (after C3) assert `ir::Paragraph.outline_level` semantics by writing a
   paragraph with a known value and re-parsing it — this pins the writer's contract to the
   standard rather than to a doc comment.
6. **Fuzz coverage gap:** `fuzz/fuzz_targets/fuzz_parse.rs:12-22` only calls
   `Document::from_reader`. The `(level + 1)` overflow at `convert_docx.rs:201` lives past that
   boundary. Add `to_ir()` / `to_markdown()` / `to_html()` calls to the target so the conversion
   and rendering layers get the same "never panic" guarantee the parsers have.
7. **Real-corpus reporting** (required by CONTRIBUTING rule 2, not automatable here): run the
   private corpus before/after and report specifically (a) how many documents lost headings,
   (b) how many changed `metadata.title`, (c) how many produced `#{7,}` before the change and
   zero after. Those three counters are the evidence the change is the intended one and not a
   wider blast radius.

---

# Anything else I found while reading

**O1 — arithmetic overflow on hostile input. [verified by reading; not executed]**
`src/convert_docx.rs:201` computes `(level + 1).min(6)` on a `u8` that the parser will happily set
to `255` (`formatting.rs:379-385` and its three siblings). `<w:outlineLvl w:val="255"/>` therefore
panics in a debug build and wraps to `Heading { level: 0 }` in release (`Cargo.toml:155-159`
declares no `overflow-checks`). A `level: 0` heading renders as `"".repeat` + `" "` + text
(`ir_render.rs:272`), i.e. a markdown line starting with a stray space. This violates the
project's own robustness rule ("Malformed, truncated or hostile files must return an error — never
panic", CONTRIBUTING:159) and is not covered by the fuzz target (O6). Range-validating at the
parse boundary fixes it as a side effect.

**O2 — `Section.title` is emitted as a synthetic `## ` heading. [verified]**
`ir_render.rs:253-258` (markdown) and `:166-172` (plain text) prepend `Section.title`, which
`convert_docx.rs:70-84` sets to the text of the section's first `Heading` element — so that
heading is rendered twice, the second time at a level unrelated to its real one. Same for PPTX
(`convert_pptx.rs:32-38` + `:76`). This is probably the single largest contributor to symptom 3's
"different heading output" and is trivially demonstrable with a two-paragraph synthetic file.

**O3 — IR markdown drops headers and footers. [verified]**
`convert_docx.rs:25-59` builds `Section.header`/`Section.footer`; `render_section_markdown`
(`ir_render.rs:253-267`) never looks at them. So the IR path loses content the direct path
deliberately preserves (`text.rs:26-31` documents *why* it preserves it).

**O4 — header/footer classification by index. [inferred]**
`convert_docx.rs:30-46` decides header-vs-footer with `if idx < n_header_refs`, i.e. it assumes
`doc.headers_footers` is ordered all-headers-then-all-footers. `docx/text.rs:80` uses the
per-part `hf.is_header` flag instead. If the parts are interleaved (a document with
first/even/default headers and footers per section), the IR path will label footers as headers.
The correct field is right there and unused.

**O5 — heading rendered inside a table cell. [verified]**
`ir_render.rs:377-386` renders any non-`Paragraph` cell element through
`render_element_markdown`, so a `Heading` in a cell emits `## Foo` *inside* a `| … |` row,
producing broken markdown. The direct path flattens cells to plain text (`text.rs:320-322`).

**O6 — fuzz target stops at parse. [verified]** See testing item 6.

**O7 — ordered-list numbering in the direct path is constant. [verified]**
`src/docx/text.rs:176-180` formats every numbered item as `format!("{}. ", level.start)` — the
level's *start* value, not a running counter — so a 1,2,3 list renders as `1. 1. 1.`. Unrelated
to outline levels but in the same function and would disappear with path unification.

**O8 — empty heading paragraphs emit a bare marker. [inferred]**
`text.rs:188-191` pushes the hashes before rendering content, so a paragraph carrying an outline
level and no runs yields a line that is just `#`. `ir_render.rs:271-275` yields `"# "` (non-empty,
so it survives the `filter(|s| !s.is_empty())` at `:262`). Both are markdown noise; worth a guard
in whichever renderer survives.

**O9 — `docs/specs/docx_spec.md:605` under-documents the value space.** See C4.

**O10 — legacy `.doc` headings are a text heuristic, not outline levels. [verified]**
`src/convert_doc.rs:443-472` classifies a paragraph as a heading by `trimmed.len() < 100` and
similar. CONTRIBUTING:156 already lists "outline levels" as an unimplemented legacy-format
structure, so this is known; noting it because whatever `heading_level` abstraction comes out of
the fix above is the natural place for the `.doc` `sprmPOutLvl` path to plug in later.

---

## Sources

- [ECMA-376/ISO-IEC 29500-1 §17.3.1.20 `outlineLvl` (clause text reproduced by Microsoft)](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc882417(v=office.14))
- [Same clause on current Learn (OutlineLevel class)](https://learn.microsoft.com/dotnet/api/documentformat.openxml.wordprocessing.outlinelevel?view=openxml-3.0.1)
- [ECMA-376/ISO-IEC 29500-1 §17.3.1.27 `pStyle` — style hierarchy / direct formatting precedence](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc866133(v=office.14))
- [MS-OI29500, Part 1 Section 17.16.5.68 TOC — Word includes any style whose definition has a matching `outlineLvl`](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/68f1577f-0efe-453b-bad9-cfd69f740b30)
- [MS-OE376, Part 4 Section 2.18.16 ST_DecimalNumber — restriction on XSD integer; Word uses int](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/e8c9b787-495c-4f04-9862-523e8db56a04)
- [Microsoft sample `styles.xml` — `Heading1` carries `<w:outlineLvl w:val="0"/>`](https://learn.microsoft.com/en-us/dotnet/standard/linq/style-part-wordprocessingml-document)
- [WdOutlineLevel enumeration — `wdOutlineLevelBodyText` = 10 (i.e. the 10th level, XML value 9)](https://learn.microsoft.com/office/vba/api/word.wdoutlinelevel)
- [Word VBA `Paragraphs.OutlineLevel` — heading-styled paragraphs' outline level is fixed by the style](https://learn.microsoft.com/office/vba/api/word.paragraphs.outlinelevel)
- [LibreOffice `DomainMapper.cxx` — out-of-range ignored "by MS Word"; 9 → LO body level 0](https://raw.githubusercontent.com/LibreOffice/core/master/sw/source/writerfilter/dmapper/DomainMapper.cxx)
- [LibreOffice `PropertyMap.hxx` — `WW_OUTLINE_MIN`/`WW_OUTLINE_MAX` = 0/9](https://raw.githubusercontent.com/LibreOffice/core/master/sw/source/writerfilter/dmapper/PropertyMap.hxx)
- [LibreOffice `StyleSheetTable.cxx` — style-level 9 → body, parent-style inheritance](https://raw.githubusercontent.com/LibreOffice/core/master/sw/source/writerfilter/dmapper/StyleSheetTable.cxx)
- [Pandoc `Docx/Parse/Styles.hs` — `getHeaderLevel` uses the `heading N` style name only](https://raw.githubusercontent.com/jgm/pandoc/main/src/Text/Pandoc/Readers/Docx/Parse/Styles.hs)
- [Pandoc `Docx/Parse.hs` — `pHeading` resolves through the style chain](https://raw.githubusercontent.com/jgm/pandoc/main/src/Text/Pandoc/Readers/Docx/Parse.hs)
- [Pandoc `Writers/HTML.hs` — level > 6 degrades to `<p class="heading">`](https://raw.githubusercontent.com/jgm/pandoc/main/src/Text/Pandoc/Writers/HTML.hs)
- [Pandoc `Writers/Markdown.hs` — ATX hashes are not clamped](https://raw.githubusercontent.com/jgm/pandoc/main/src/Text/Pandoc/Writers/Markdown.hs)
- [mammoth.js `lib/options-reader.js` — default style map, Heading 1–6 only](https://raw.githubusercontent.com/mwilliamson/mammoth.js/master/lib/options-reader.js)
- [Apache POI `XWPFParagraph.java` — `getStyle()` returns the pStyle id; no outline-level API](https://raw.githubusercontent.com/apache/poi/trunk/poi-ooxml/src/main/java/org/apache/poi/xwpf/usermodel/XWPFParagraph.java)
- [wordprocessingml.com — "Determining if a paragraph is actually a heading" (outlineLvl vs pandoc's style-name approach)](https://wordprocessingml.com/docs/headings/determining-if-a-paragraph-is-actually-a-heading/)
- [Charles Kenyon — Navigation Pane is driven by outline level ("It does not guess")](https://www.addbalance.com/usersguide/navigationPane.htm)
- [Office Watch — headings vs outline levels (conflicting secondary account)](https://office-watch.com/2025/word-headings-vs-outline-levels/)
- [docx4java forum — a real `TOCHeading` definition with `<w:outlineLvl w:val="9"/>`](https://www.docx4java.org/forums/docx-java-f6/genearting-toc-without-skippagenumbers-not-working-t3036.html)
