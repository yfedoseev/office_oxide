# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.2] - 2026-05-14

> Round-trip fidelity, IR layout features, embedded fonts, XLSX number formatting, and an O(1) style-lookup perf win.

### Performance

- **XLSX styles**: cell-format lookups now use a `HashMap`, replacing
  the linear `Vec` scan in `format_cell_value` / `is_date_cell`.
  Per-cell formatting becomes O(1); large styled workbooks parse
  noticeably faster with no API change.

### Round-trip fidelity (PDF → office → PDF)

- **Alignment, spacing, footers, and horizontal rules** preserved end-to-end
  through both `to_docx` and `to_pptx` writers.
- **Images, fonts, and column layouts** preserved across DOCX, PPTX, and
  XLSX. Source-PDF font programs that previously registered as empty
  subsets now embed correctly.
- **`Element::ThematicBreak`** encoded in PPTX as a centered 30-char run
  of `U+2500 BOX DRAWINGS LIGHT HORIZONTAL`. Downstream PDF renderers
  detect the all-U+2500 content and re-emit a real horizontal rule.
- **DOCX horizontal rules** recovered from the conventional encoding
  (empty paragraph + `<w:pBdr><w:bottom/>`) back into `Element::ThematicBreak`.

### DOCX

- **`<w:framePr>` parsed into IR** as `FramePosition` (twips, page-anchored)
  on both `Paragraph` and `Heading`. Used by layout-preserving paths
  (e.g. pdf_oxide's `to_docx_bytes_layout`).
- **Floating drawings and vector shapes**: `<wp:anchor>` images plus
  `<wps:wsp>` preset shapes (line, rect) with stroke/fill RGB and
  stroke width round-trip through `DrawingInfo`.
- **Per-section page sizes** preserved through `to_ir`; multi-section IR
  emits per-section `<w:sectPr>`.
- **`<w:sz>` preserved** through to IR's `font_size_half_pt`.
- **Run colour** from `<w:rPr><w:color w:val="RRGGBB"/>` propagated into
  `TextSpan.color` during `to_ir`, so PDF→DOCX→PDF round-trips keep
  coloured text. Only the `ColorRef::Rgb` variant is plumbed today;
  theme / system / `auto` colours still fall through to the renderer
  default (proper resolution needs `theme.xml` threaded into the
  convert path).
- **Headers and footers** now included in `to_markdown` and `to_ir`
  (previously silently dropped).
- **Embedded fonts** under `/word/fonts/` exposed on
  `DocxDocument.embedded_fonts`. `strip_embedded_font_filename` recovers
  the original face name from `font_<n>_<face>.<ext>` (fixes greedy
  alphabetic-trim regression where `TeXGyreTermesX-` was returned
  instead of `TeXGyreTermesX-Regular`).
- **`parse_drawing` decomposed** into focused recursive helpers
  (`parse_inline_or_anchor_body`, `parse_anchor_position`,
  `parse_shape_properties`, etc.) for readability.
- **Run-level `<w:rFonts w:ascii>` plumbed into `TextSpan.font_name`**;
  `<w:cols>` propagated to `Section.columns`.

### PPTX

- **Pagination**: each slide forces a `SectionBreakType::NextPage` so two
  slides never share a rendered page.
- **Real Title+Body slide layout** emitted by the writer instead of a blank
  layout, so PowerPoint shows placeholder hints in edit mode.
- **Slide background**: `<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr>`
  parsed into `Slide.background_rgb` and propagated to `Section.background_rgb`.
- **Positioned text boxes**: shapes with explicit `<a:xfrm>` coordinates
  wrap their content in `Element::TextBox` so downstream renderers can
  place them at absolute EMU coordinates. Zero-size shapes skip the wrapper.
- **Slide size → page setup**: `<p:sldSz cx=… cy=…>` propagated to each
  section's `PageSetup`.
- **Run font sizes preserved** via new `TextRun.font_size_hundredths_pt`
  (parsed from `<a:rPr sz="…"/>`).
- **Run colour preserved** via new `TextRun.color_rgb: Option<[u8; 3]>`
  parsed from `<a:rPr><a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>`
  and propagated to `TextSpan.color` in IR. The parser tracks an
  `in_solid_fill` flag so sibling effects (e.g. `<a:hl><a:srgbClr/>`
  for hyperlink colour) don't leak into the run's own fill; non-sRGB
  fills (gradient, scheme colour) fall back to `None`.
- **Paragraph alignment** parsed from `<a:pPr algn="…"/>` (all five
  variants: `l` / `ctr` / `r` / `just` / `dist`) into
  `TextParagraph.alignment`. **Space-before** parsed from
  `<a:spcBef><a:spcPts val=…/>`.
- **Title alignment propagation**: `find_title` returns text + first
  paragraph's alignment, seeding both `Section.title` and the synthesised
  level-2 Heading's alignment.
- **Picture shapes** now carry `embed_rid`, `data`, and `format`
  (resolved via a pre-built media map at parse time, so the parallel
  slide parser doesn't need the OPC reader).
- **Font embedding** under `/ppt/fonts/`.
- **Structured chart text extraction**: `<c:chart>` parts parsed into
  per-chart text blocks rendered as `## Chart N` in markdown / search /
  PDF without needing a graphical chart renderer.
- **Compaction**: consecutive H1/H2 cover-page headings fold into one
  slide instead of fragmenting; long XLSX paragraphs split across cells
  to respect ~32k char-per-cell limits.
- **Slide cap**: writer caps at ~250 slides (PowerPoint's hard limit).

### XLSX

- **Per-worksheet `page_setup`** round-trips via `<pageMargins>` (inches)
  and `<pageSetup>` (paperWidth/paperHeight with mm/cm/in suffix or
  `paperSize` enum 1–13). New `Worksheet.page_setup`.
- **`numfmt` module** (`crate::xlsx::numfmt`): built-in IDs 0–44 (general,
  fixed, commas, percent, currency, scientific, accounting) and a
  simplified custom format-string parser (multi-section, `[Red]` color
  directives stripped, currency prefix from `[$€-407]`, quoted literal
  suffix, percent and thousands separators). Applied to numeric cells
  during `format_cell_value` and `write_cell_value_fast`.
- **Font sizes** preserved through IR; long-text single-column sheets
  emit as paragraphs instead of a tall 1-column GFM table.
- **Unique worksheet names** in `ir_to_xlsx` (duplicates suffixed with
  `_2`, `_3`, …).
- **Drawings**: `xl/drawings/drawingN.xml` parsed into
  `Worksheet.images` (`WorksheetPicture` with EMU coords + bytes) and
  `Worksheet.text_shapes` (`WorksheetTextShape` for layout-mode text
  boxes from `to_xlsx_bytes_layout`).
- **Embedded fonts** under `/xl/fonts/`.

### IR enrichment

- **New types**: `Shape` (vector shape anchored at absolute EMU coords),
  `ShapeGeom` (`Line`, `Rect`), `FramePosition` (twip-anchored frame).
- **`Heading`** gains `frame_position` + `alignment`.
- **`Section`** gains `background_rgb`.
- **`ParagraphAlignment`** gains the `Distribute` variant.
- **`Element::Shape(Shape)`** variant for vector shapes.
- **New helpers**: `first_inline_font_size_pt`, `inline_to_element_block`,
  `build_nested_list` (flat / 2-level / 3-level recursion).
- **Centralized defaults** in `ir_render::block_default`: ThematicBreak
  renders as `"---"` / `<hr />`; PageBreak / ColumnBreak / Shape are
  invisible in flow; TextBox / Footnote / Endnote recursively render
  children. Adding a new `Element` variant forces a compile error
  in `block_default::default_plain` instead of silent fallthrough.

### Core

- **`crate::core::core_properties`**: shared `docProps/core.xml` generator
  used by all three writers. Emits `dc:title`, `dc:creator`, `dc:subject`,
  `dc:description`, `cp:keywords`, `dcterms:created`, `dcterms:modified`
  from the IR's `Metadata`. Empty fields are omitted entirely.
- **`crate::core::embedded_fonts`**: unified font-embedding helper
  (`write_embedded_fonts`, `sanitize_font_filename`). All three formats
  share the layout `<prefix>font_<n>_<safe_name>.ttf`.
- **`HalfPoint::from_word_sz` / `from_drawingml_sz` / `to_drawingml_sz` /
  `from_points_rounded`**: cross-format font-size invariants
  (DrawingML hundredths-of-a-point vs WML half-points).

### Dependencies

- **`quick-xml` 0.37 → 0.40**: upstream removed `BytesText::unescape()`
  and deprecated `Attribute::unescape_value()` (its replacement
  `normalized_value()` has different semantics — no entity
  unescaping). Migration added two helpers in `core::xml`:
  `unescape_text(BytesText) -> Result<String>` (used by 6 call sites)
  and `unescape_attr_value` (used by 6 call sites, with
  `#[allow(deprecated)]` localised to the helper so call sites stay
  deprecation-free). 535 / 535 tests still pass; clippy clean.
- **`koffi` 2.16.1 → 2.16.2** in `js/` (patch bump).

### Documentation

- **CLI / MCP crate-level docs**: `office_oxide_cli` and
  `office_oxide_mcp` previously opened with `mod commands;` /
  `mod protocol;` and had no crate-level rustdoc. Added a short
  `//!` block plus `#![warn(missing_docs)]` so future items in
  either binary stay documented.
  `RUSTDOCFLAGS="-D missing_docs" cargo doc --workspace --no-deps
  --features parallel,mmap` now passes with zero errors.

### Tests

- **+98 unit tests** across the modules touched in this release:
  `core::embedded_fonts`, `core::core_properties`, `core::units`,
  `xlsx::numfmt`, `xlsx::worksheet`, `docx::formatting`, `docx::mod`,
  `pptx::slide`, `ir`, `ir_render`.
- **535 / 535 tests pass** across default, `--features parallel`,
  `--features mmap`, and `--features parallel,mmap` builds.
- `cargo fmt` clean. `cargo clippy --workspace --all-targets -- -D warnings`
  clean.

### Bindings

- **Python wheel** (maturin, PyO3 0.28) builds cleanly and exposes
  `Document`, `EditableDocument`, `XlsxWriter`, `PptxWriter`,
  `OfficeOxideError`, `create_from_markdown`, `extract_text`,
  `to_markdown`, `to_html`, `version`.
- **WASM** package (`wasm-pack build --target web/node/bundler`) builds
  cleanly with `--features wasm`.
- **C#** package bumped to 0.1.2 (csproj only — no API changes).

[0.1.2]: https://github.com/yfedoseev/office_oxide/compare/v0.1.1...v0.1.2

## [0.1.1] - 2026-04-30

> Richer IR type system, DOCX writer output, improved PPTX/XLSX IR renderers, and writer APIs in all language bindings

### IR — extended type system

- **`TextSpan`** gains nine typography fields: `font_name`, `font_size_half_pt`,
  `color`, `highlight`, `underline` (`UnderlineStyle` enum), `vertical_align`
  (`VerticalAlign`: Superscript / Subscript / Baseline), `all_caps`,
  `small_caps`, `char_spacing_half_pt`.
- **`Paragraph`** gains twelve layout fields: `alignment` (`ParagraphAlignment`:
  Left / Center / Right / Justify / Distribute), `indent_left_twips`,
  `indent_right_twips`, `first_line_indent_twips`, `space_before_twips`,
  `space_after_twips`, `line_spacing` (`LineSpacing`: Auto / Multiple / Exact /
  AtLeast), `background_color`, `border` (`ParagraphBorder`), `keep_with_next`,
  `keep_together`, `page_break_before`, `outline_level`.
- **`Table`** gains `width_twips`, `column_widths_twips`, `border`
  (`TableBorder`), `alignment` (`TableAlignment`), `indent_left_twips`,
  `cell_padding_twips`, `caption`.
- **`TableRow`** gains `height_twips`, `allow_break`, `repeat_as_header`.
- **`TableCell`** gains `background_color`, `border`, `vertical_align`
  (`CellVerticalAlign`), `text_align`, `width_twips`, `padding` (`CellPadding`),
  `text_direction` (`TextDirection`).
- **`Image`** gains `data`, `format` (`ImageFormat`), `pixel_width`,
  `pixel_height`, `display_width_emu`, `display_height_emu`, `decorative`,
  `positioning` (`ImagePositioning`: Inline or Floating with `FloatingImage`).
- **`Section`** gains `page_setup` (`PageSetup`), `columns` (`ColumnLayout`),
  `break_type` (`SectionBreakType`), and six header/footer slots
  (`header`, `footer`, `first_page_header`, `first_page_footer`,
  `even_page_header`, `even_page_footer`).
- **`Metadata`** gains `author`, `subject`, `keywords`, `created`, `modified`,
  `description` (written to `docProps/core.xml`).
- **New `Element` variants**: `TextBox`, `PageBreak`, `ColumnBreak`,
  `Footnote(Note)`, `Endnote(Note)`, `CodeBlock`.
- **`List`** gains `start_number`, `style` (`ListStyle`), `level`.
  `ListItem.content` promoted from `Vec<InlineContent>` to `Vec<Element>`
  to allow block-level content (tables, images) inside list items.
- **`InlineContent`** gains `FootnoteRef` and `EndnoteRef` variants.
- **New supporting types**: `BorderLine`, `TableBorder`, `ParagraphBorder`,
  `CellPadding`, `FloatingImage`, `HeaderFooter`, `TextBox`, `Note`,
  `FootnoteRef`, `CodeBlock`, `PageSetup`, `ColumnLayout`.
- **New enums**: `UnderlineStyle`, `ParagraphAlignment`, `LineSpacing`,
  `BorderStyle`, `CellVerticalAlign`, `TableAlignment`, `TextDirection`,
  `ImageFormat`, `ImagePositioning`, `SectionBreakType`, `VerticalAlign`,
  `FloatAnchor`, `TextWrap`, `ListStyle`.
- All new fields are `Option<_>` or default to `false`/`None` — fully
  backwards compatible; existing callers require only `..Default::default()`
  on struct literals.

### DocxWriter — OOXML emission for all new fields

- **Run properties**: `<w:rFonts>`, `<w:sz>`/`<w:szCs>`, `<w:color>`,
  `<w:shd>` (highlight), `<w:u>`, `<w:vertAlign>`, `<w:caps>`,
  `<w:smallCaps>`, `<w:spacing>` (character spacing).
- **Paragraph properties**: `<w:jc>`, `<w:ind>`, `<w:spacing>` (before/after/
  line), `<w:pBdr>`, `<w:shd>`, `<w:keepNext>`, `<w:keepLines>`,
  `<w:pageBreakBefore>`, `<w:outlineLvl>`.
- **Table**: `<w:tblW>`, `<w:tblInd>`, `<w:tblBorders>`, `<w:jc>` (table),
  `<w:tblCellMar>`, `<w:tblGrid>`/`<w:gridCol>`.
- **Table row**: `<w:trHeight>`, `<w:tblHeader>`, `<w:cantSplit>`.
- **Table cell**: `<w:tcW>`, `<w:gridSpan>`, `<w:vMerge>`, `<w:shd>` (cell),
  `<w:tcBorders>`, `<w:vAlign>`, `<w:textDirection>`, `<w:tcMar>` (per-edge
  padding), cell-level `text_align` propagated to contained paragraphs.
- **Table caption**: emitted as a `Caption`-styled paragraph before `<w:tbl>`.
- **Images**: inline `<wp:inline>` and floating `<wp:anchor>` with
  `<wp:wrapSquare>` / `<wp:wrapTight>` / `<wp:wrapThrough>` /
  `<wp:wrapTopAndBottom>` / `<wp:wrapNone>`.
- **Text boxes**: `<wp:anchor>` + `<wps:txbx>` + `<w:txbxContent>`.
- **Sections**: `<w:sectPr>` with `<w:pgSz>`, `<w:pgMar>`, `<w:cols>` (uniform
  and per-column widths with separator rule), `<w:type>` (Continuous / NextPage
  / EvenPage / OddPage). Header/footer parts written to `/word/header*.xml` and
  `/word/footer*.xml` with correct relationship entries.
- **Footnotes / endnotes**: `/word/footnotes.xml` and `/word/endnotes.xml`
  parts; inline `<w:footnoteReference>` / `<w:endnoteReference>` runs with
  `FootnoteReference` / `EndnoteReference` character styles.
- **`{PAGE}` / `{NUMPAGES}` sentinels** in header/footer text spans emitted as
  `<w:fldChar>` / `<w:instrText>` field runs.
- **Page break**: `<w:br w:type="page">`. **Column break**: `<w:br
  w:type="column">`.
- **Code blocks**: `Code`-styled paragraph with preserved whitespace and
  line breaks.
- **Lists**: `<w:numFmt>` driven by `ListStyle`; `<w:startOverride>` for
  non-1 `start_number`; block-level item content (tables, images) written
  alongside the numbered paragraph.
- **Metadata**: `author`, `subject`, `keywords`, `description`, `created`,
  `modified` written to `docProps/core.xml` as Dublin Core properties.

### PPTX IR renderer — `ir_to_pptx`

- **Rich text runs**: paragraphs and headings now emit styled `<a:r>` runs via
  `add_rich_text()` — bold, italic, font size, color, and font name from
  `TextSpan` are all preserved. Previously all formatting was stripped.
- **Table elements**: rendered as tab-separated cell text (rows joined with `\n`)
  instead of being silently dropped.
- **Image elements**: embedded as native PPTX media via the new
  `SlideData::add_image()` API — writes a `<p:pic>` shape with `<p:blipFill>`
  and an OPC media part. PNG, JPEG, and GIF are supported.
- **CodeBlock elements**: rendered with Courier New font run instead of being
  silently dropped.
- **Slide dimensions**: first `Section.page_setup` is now forwarded to
  `PptxWriter::set_presentation_size()` (1 twip = 914 400/1 440 EMU), fixing
  clipped output for landscape A4 and other non-16:9 documents.
- **New `PptxWriter::set_presentation_size(cx, cy)`** method; emits `<p:sldSz>`
  with the correct EMU values instead of always writing the 16:9 default.

### XLSX IR renderer — `ir_to_xlsx`

- **Header row styling**: rows with `TableRow.is_header = true` are now written
  with bold weight and a grey (`D3D3D3`) background via `set_cell_styled()`.
  `TableCell.background_color` overrides the default grey when set.
- **Cell background color**: non-header cells with `TableCell.background_color`
  set now get a solid fill style applied.
- **Column widths**: `Table.column_widths_twips` is now converted to Excel
  character-width units (`twips × 96 / (1440 × 7)`, clamped 3–80) and written
  via `set_column_width()`.
- **Merged cells**: `TableCell.col_span` and `row_span` > 1 now emit a
  `<mergeCells>/<mergeCell ref="…"/>` block in the worksheet XML instead of
  being ignored.
- **New `SheetData::merge_cells(row, col, row_span, col_span)`** method; inserts
  `<mergeCells>` between `</sheetData>` and `</worksheet>`.
- Row cursor tracks absolute position across all elements in a section, so
  paragraphs and headings interleaved with tables land in the correct rows.

### Writer APIs — all language bindings

`XlsxWriter` and `PptxWriter` (previously Rust-only) are now callable from
every binding layer via a new index-based C FFI surface:

**New C FFI symbols** (`include/office_oxide_c/office_oxide.h`):

- `office_xlsx_writer_new/free`, `office_xlsx_writer_add_sheet` (returns sheet
  index), `office_xlsx_sheet_set_cell`, `office_xlsx_sheet_set_cell_styled`
  (bold + hex background), `office_xlsx_sheet_merge_cells`,
  `office_xlsx_sheet_set_column_width`, `office_xlsx_writer_save`,
  `office_xlsx_writer_to_bytes`
- `office_pptx_writer_new/free`, `office_pptx_writer_set_presentation_size`,
  `office_pptx_writer_add_slide` (returns slide index),
  `office_pptx_slide_set_title`, `office_pptx_slide_add_text`,
  `office_pptx_slide_add_image` (PNG/JPEG/GIF bytes + EMU geometry),
  `office_pptx_writer_save`, `office_pptx_writer_to_bytes`

**Go** — `XlsxWriter` and `PptxWriter` structs with CGo wrappers and
`runtime.SetFinalizer` for safe GC.

**C# / .NET** — new `OfficeOxide.XlsxWriter` and `OfficeOxide.PptxWriter`
classes (`IDisposable`); P/Invoke declarations added to `NativeMethods`.

**Node.js** — `XlsxWriter` and `PptxWriter` ESM + CJS classes; koffi
function prototypes; TypeScript `ImageFormat` type and class declarations.

**Python** — `XlsxWriter` and `PyO3PptxWriter` PyO3 classes calling Rust
directly (no C FFI round-trip); exported from `office_oxide` with full type
stubs in `_native.pyi`.

### Bug fixes

- `TableCell.padding` (`CellPadding`) was defined in the IR but silently
  dropped by the writer; now emits `<w:tcMar>` with per-edge twip values.
- `TableCell.text_align` was defined in the IR but silently dropped; now
  propagated to contained paragraphs (respects pre-existing paragraph
  alignment, so explicit paragraph alignment is never overwritten).
- `Table.caption` was defined in the IR but silently dropped; now emitted
  as a `Caption`-styled paragraph immediately before the table.

## [0.1.0] - 2026-04-27

> Initial public release

### Cross-language bindings

- **Rust core** (`office_oxide` on crates.io): unified `Document` handle for
  all six formats, `EditableDocument` for DOCX/XLSX/PPTX editing, format-
  agnostic `DocumentIR`.
- **Python** (`office-oxide` on PyPI): context-manager `Document` /
  `EditableDocument`, `os.PathLike` support, complete type stubs in
  `_native.pyi` (`Literal` format names, `_Path` alias).
- **Go** (`github.com/yfedoseev/office_oxide/go`): CGo wrapper over the C FFI
  with idiomatic `Open` / `Close` / error-return API, `go/cmd/install` helper
  that fetches the matching native archive and prints the
  `CGO_CFLAGS` / `CGO_LDFLAGS` to export.
- **C# / .NET** (`OfficeOxide` on NuGet): `LibraryImport` P/Invoke,
  `IDisposable`, `async/await`, `IsAotCompatible=true`, `IsTrimmable=true`.
  Four `SetCell` overloads + `SetCellEmpty`. Net 8 and net 10 target
  frameworks.
- **Node.js native** (`office-oxide` on npm): [koffi](https://koffi.dev)-based,
  no node-gyp, ESM + CJS entry points with an `exports` map, TypeScript
  definitions, `Symbol.dispose` support, platform prebuilds staged into
  `prebuilds/<platform>-<arch>/`.
- **WASM** (`office-oxide-wasm` on npm): three sub-path exports — default
  ESM for bundlers, `office-oxide-wasm/node` for CJS, `office-oxide-wasm/web`
  for native-ESM browser imports. TypeScript definitions shipped.
- **C FFI** (`include/office_oxide_c/office_oxide.h`): stable
  `office_document_*` / `office_editable_*` surface with out-param error
  codes and explicit memory ownership. Exported from the cdylib + staticlib;
  the substrate that Go, C#, and Node-native link against.

### Tooling

- **CLI** (`office-oxide` binary): `text`, `markdown`, `html`, `info`, `ir`
  subcommands.
- **MCP server** (`office-oxide-mcp` binary): `extract` and `info` tools
  over JSON-RPC 2.0 / stdio.

### Performance

- Up to 100× faster than `python-docx`, `openpyxl`, `python-pptx`, `xlrd`.
- Beats `calamine` on XLSX and all Rust / Python alternatives on .xls.
- **100% pass rate on valid Office files** (6,062-file corpus: LibreOffice,
  Apache POI, python-pptx, python-docx, Pandoc, etc.). All 97 non-passing
  files are invalid inputs — corrupted ZIPs, missing required parts, malformed
  XML, or non-Office files with Office extensions.

### Documentation & examples

- Per-language getting-started guides in [`docs/`](docs/): Rust, Python,
  Go, C#, JavaScript (native), WASM, and C / raw FFI.
- Per-binding READMEs in [`python/`](python/README.md), [`go/`](go/README.md),
  [`csharp/OfficeOxide/`](csharp/OfficeOxide/README.md),
  [`js/`](js/README.md), [`wasm-pkg/`](wasm-pkg/README.md).
- Identical `extract` / `replace` / `read_xlsx` demos per language under
  [`examples/`](examples/).

### Release CI

- Version parity across `Cargo.toml`, `pyproject.toml`,
  `wasm-pkg/package.json`, `js/package.json`, and
  `csharp/OfficeOxide/OfficeOxide.csproj`.
- 6-target native-lib build matrix producing `.tar.gz` / `.zip` archives
  with the shared library, static archive, and public header.
- 3-target WASM build (bundler / nodejs / web) with per-target module-type
  hints so Node + bundlers + browsers each load the right code.
- NuGet packaging with `runtimes/<rid>/native/` prebuilts, Node
  `prebuilds/<platform>-<arch>/` staging, and a `go/v*` module tag for the
  Go module proxy.

## [0.1.0-draft] - 2026-04-07 (not released)

> Internal milestone before the public release above.

Single consolidated `office_oxide` crate with modules per format, plus
companion workspace crates `office_oxide_cli` and `office_oxide_mcp`.

### Added

- **Unified API** (`crate::Document`)
  - `Document::open()`, `Document::from_reader()`, `plain_text()`,
    `to_markdown()`, `to_html()`, `to_ir()`, `format_name()`
  - Format auto-detection from file extension for all 6 formats
  - `as_docx()` / `as_xlsx()` / `as_pptx()` / `as_doc()` / `as_xls()` /
    `as_ppt()` escape hatches for format-specific types
  - Convenience functions: `extract_text()`, `to_markdown()`, `to_html()`
    at crate root

- **Format-agnostic IR** (`crate::ir::DocumentIR`)
  - Sections, Elements (Heading, Paragraph, Table, List, Image,
    ThematicBreak), serializable to/from JSON
  - DOCX→IR: heading detection via outline level + style resolution,
    list grouping, table vMerge → row_span
  - XLSX→IR: worksheets → sections, cell grids → tables, first row as
    header
  - PPTX→IR: slides → sections, spatial sort, title placeholder → section
    title
  - IR renderers: `plain_text()` and `to_markdown()`

- **OOXML module — `crate::docx`** (36 tests)
  - SAX-style parsing of `document.xml`, `styles.xml`, `numbering.xml`
  - Hyperlink resolution, heading detection, list grouping
  - Plain text, Markdown, HTML extraction

- **OOXML module — `crate::xlsx`** (57 tests)
  - Shared string table, cell parsing, 1900 date system with Lotus bug
  - Built-in + custom number format detection
  - CSV (RFC 4180), Markdown (pipe tables), HTML output
  - Rich-text and style support

- **OOXML module — `crate::pptx`** (40 tests)
  - Slide parsing with spatial sort, shape types (AutoShape, Picture,
    Group, GraphicFrame, Connector), text body extraction
  - Notes slide support, inline hyperlink resolution
  - Plain text, Markdown, HTML extraction

- **Legacy module — `crate::cfb`** (18 tests)
  - Pure Rust CFBF / OLE2 container reader
  - Supports v3 (512-byte) and v4 (4096-byte) sectors, mini-streams,
    case-insensitive stream access, path-based lookup

- **Legacy module — `crate::doc`** (15 tests)
  - Word Binary (.doc) parser built on `crate::cfb`
  - FIB parsing, piece-table extraction, dual encoding
    (compressed CP1252 + Unicode UTF-16LE)
  - Field-code stripping and special-char sanitization

- **Legacy module — `crate::xls`** (24 tests)
  - Excel Binary (.xls) BIFF8 parser built on `crate::cfb`
  - CONTINUE record merging, SST (compressed + wide + rich text),
    RK decode
  - Cell records: LABELSST, NUMBER, RK, MULRK, FORMULA, BOOLERR,
    LABEL, BLANK

- **Legacy module — `crate::ppt`** (15 tests)
  - PowerPoint Binary (.ppt) parser built on `crate::cfb`
  - 8-byte record headers, container vs atom records
  - Text extraction from TextCharsAtom (UTF-16LE) and TextBytesAtom
    (Latin-1), SlideListWithText grouping

- **Shared core — `crate::core`** (55 tests)
  - `OpcReader` / `OpcWriter` for ZIP-based OPC packages
  - Theme parsing, color resolution, unit types (`Twip`, `HalfPoint`,
    `Emu`)
  - Namespace-aware XML utilities with OOXML Strict support

- **Creation API** — write OOXML documents from scratch
  - `crate::docx::write::DocxWriter`, `crate::xlsx::write::XlsxWriter`,
    `crate::pptx::write::PptxWriter`
  - `crate::create::create_from_ir()` and
    `create_from_ir_to_writer()` for IR-to-format conversion

- **Editing API** — modify existing documents while preserving unmodified
  parts
  - `crate::docx::edit::EditableDocx`, `crate::xlsx::edit::EditableXlsx`,
    `crate::pptx::edit::EditablePptx`
  - Unified `crate::edit::EditableDocument` with `replace_text()`
    and `set_cell()`
  - `crate::core::editable::EditablePackage` round-trips rels via
    `RelationshipsBuilder::add_with_id()`

- **Python bindings** — PyO3 0.28, extension module `office_oxide._native`
  - Type stubs (`_native.pyi`) and PEP 561 `py.typed` marker
  - Wheels for Linux, macOS, Windows across Python 3.8–3.14

- **WASM bindings** — `wasm-bindgen`, `WasmDocument` class
  - `new(data, format)`, `plainText()`, `toMarkdown()`, `toHtml()`,
    `toIr()`
  - npm package `office-oxide-wasm`

- **CLI** — `office_oxide_cli` workspace crate, binary `office-oxide`
  - Subcommands: `text`, `markdown`, `html`, `info`, `ir`

- **MCP server** — `office_oxide_mcp` workspace crate, binary
  `office-oxide-mcp`
  - JSON-RPC 2.0 over stdin/stdout
  - Tools: `extract` (text / markdown / html / ir), `info`

- **Feature flags**
  - `python` — PyO3 extension module
  - `wasm` — wasm-bindgen bindings
  - `mmap` — `memmap2`-backed file reading
  - `parallel` — `rayon`-based parallel parsing

### Robustness

- OOXML Strict namespace support via dual-namespace matching
  (`strict_alternate()`) and relationship-type normalization
- Case-insensitive ZIP entry lookup with backslash path normalization
- Namespace-agnostic attribute lookup (`optional_prefixed_attr_str()`)
  for `d3p1:id` and similar
- Percent-encoding decoding in `PartName::new()`
- CRC checksum tolerance in `read_zip_entry()` (accepts data on mismatch)
- Tolerant numeric parsing: `parse_numeric()` strips unit suffixes,
  handles decimals
- Shared-string DoS cap: `MAX_CELL_STRING_LEN = 32_768` in `xlsx`
- Optional parts (`numbering.xml`, theme) degrade gracefully on read
  errors
- XLSX border aliases: `start`/`end` mapped to `left`/`right` for Strict
  OOXML

### Validation

- **98.4% pass rate on a 6,062-file corpus** (5,965 / 6,062) across 11
  open-source test suites (LibreOffice Core, Apache POI, Open XML SDK,
  ClosedXML, Pandoc, python-docx, python-pptx, Apache Tika, calamine,
  openpreserve, oletools)
- Zero failures on legitimate Word 97+ / Excel 97+ / PowerPoint 97+
  files — all 97 non-passing files are invalid inputs: 43 invalid
  ZIP/CFB archives, 21 missing required parts, 18 malformed XML, and
  15 non-Office files (WordPerfect, pre-OLE2 Excel 3/4) misnamed with
  Office extensions
- See [BENCHMARKS.md](BENCHMARKS.md) for per-format timings and the
  full failure breakdown

[0.1.1]: https://github.com/yfedoseev/office_oxide/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/yfedoseev/office_oxide/releases/tag/v0.1.0
