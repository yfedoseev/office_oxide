# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.16] - 2026-10-29

### Added

- **Hidden text and tracked deletions can be kept.** New process-wide settings `read_options::set_include_hidden_text` and `set_include_deleted_text` (and `office-oxide --include-hidden` / `--include-deleted` on every command) bring back text a run hides (`w:vanish`, `sprmCFVanish`) and text a tracked change deleted (`w:del` runs and `w:cellDel` cells, `sprmCFRMarkDel`), in `.docx` and `.doc`, for every output. Both are off by default, so nothing read before changes; text under `w:moveFrom` stays out because the move's destination carries it.

### Fixed

- **A `.docx` body or a `.pptx` slide whose XML is malformed or cut off part-way keeps what came before the damage.** The whole document (or slide) was an error or came back as if complete. The body is now read up to the last complete top-level block (the shapes of a slide, up to the last complete shape), closed with the end tags still open there, and the loss is reported in the warnings and in the text as an unreadable part. A slide cut off mid-element used to parse to what was read and report success; its root must now have closed. Damage before the first complete block is still an error. Headers, footers, notes and other parts are not salvaged.

- **A `.doc` run whose symbol font comes from a style is read by the font's table too.** The reader looked only at the run's own `sprmCRgFtc2`. The style sheet's character properties (`UpxChpx`, [MS-DOC] §2.9.258) are now read: a run takes its own font, else the font of the character style it applies (`sprmCIstd`), else its paragraph style's, following `istdBase` (bounded, cycles end the walk). The paragraph styles are read only when a style names a symbol font.

- **A worksheet whose XML is malformed part-way through keeps the rows before the damage.** The whole sheet was reported unreadable although everything before the damage parsed. The sheet is now read up to the end of the last complete row (a row the damage reaches is dropped whole, so no row is cut in two), a notice in every rendering and a warning say the sheet was cut short, and `text_truncated` is set. A sheet with no complete row stays unreadable.

- **A `.docx` run whose symbol font comes from a style is read by the font's table too.** The mapping looked only at the run's own `w:rFonts`; Word writes a font on a run only when it differs from the style's, so a Symbol or Wingdings paragraph or character style, a style based on one, or a document default left the private-use characters as they were. The effective font (document defaults, paragraph style chain, character style chain, then the run) is now resolved once over the parsed document (body, tables, hyperlinks, headers, footers, notes and comments), without copying a face name per run.

- **Text typed in a symbol font is read by the font's table in a `.doc` too.** The reader did not read the font table or a run's font at all, so `μM` stayed `U+F06D` followed by `M` (86 corpus files hold symbol runs). A character outside U+0020..U+007F takes its font from `sprmCRgFtc2` ([MS-DOC] §2.6.1), an index into the font table (`SttbfFfn`, §2.9.286); in a run whose font is Symbol, Wingdings, Wingdings 2 or 3 or Webdings the private-use characters are now the font's, in the flat text and the structured paragraphs alike. Documents whose font table names no such font are not touched.

- **Text typed in a symbol font is read by the font's table in a `.docx` run.** Symbol, Wingdings, Wingdings 2 and 3 and Webdings store their text as private-use characters (`U+F000` plus the byte), so `μM` was written as `U+F06D` and `M`, which render as nothing. In a run whose own font is one of those, the characters are now the font's; a character without a mapping, and every run in another font, is untouched.

- **A `<w:sym>` character in Wingdings, Wingdings 2, Wingdings 3 or Webdings is read by that font's table.** Only Symbol had one, so a check mark (Wingdings `F0FC`) printed as `ü`, a checked box as `R`, and two hard-coded codes gave a bullet and an arrowhead that Wingdings does not draw there. The tables are the Unicode Consortium's normative source list for these four fonts (the `UCS;ID` list that closes "Status of encoding of Wingdings and Webdings Symbols", UTC L2/12-368 / WG2 N4384), which gives the character of each of their 866 encoded glyphs; a code the list does not name is the replacement character.

- **A `.doc` whose header version is not a Word version says that the language of its 8-bit text is unknown.** The FIB's language ID decides how 8-bit text is read; a fuzzed header can have its version, language, flags and character set overwritten, and the text (Hebrew in one file) was read in Windows-1252 with nothing said. The reader does not guess a script from byte statistics, which would be a plausible-but-wrong result; the warnings now state that the header is damaged, name the version found, and say the text was read as Windows-1252. Every genuine Word version (0x65 to 0x112) is silent.

- **A `.doc` whose piece table has a character position out of order is read, not refused.** [MS-DOC] 2.8.35 requires the positions to ascend, and 2.4.1 finds the text at a position as "the largest i such that aCp[i] <= cp", a rule defined for any array. The pieces are now resolved by that rule (piece `i` covers `[aCp[i], min(aCp[i+1..]))`, which is the old range for a sorted table), a piece left with no range covers nothing, and the warnings say the table was out of order. A fast-saved file with one wrong position was refused whole although its text was recoverable.

- **A compound file whose byte-order field is wrong but whose other header fields are intact is read.** `CfbHeader::parse` refused any value but `0xFFFE`, although [MS-CFB] 2.2 states that as a requirement on writers and gives readers no rule; a workbook with the marker swapped and a fuzzed `.doc` header were refused although only that one field was damaged. The file is read as little-endian and the container warnings name the field. A header that is really big-endian still fails, because its other fields do not validate.

- **A `.doc` reads the text of the Word documents embedded in it.** A document made of embedded Word objects came out as `[Embedded Microsoft Word Document]` placeholders and nothing else, although the objects hold the article text. The text of each object follows its placeholder in every rendering. It is read through the document's embedded-package account: the objects are counted, their streams share the byte allowance, the text is charged to the text budget (a cut is reported), and an object that cannot be read keeps its placeholder and is named in the warnings. Word 6/95 documents are read from their own stream, and the objects inside an embedded document follow it, to the account's depth and under the same limits.

- **A `<w:sym>` character in the Symbol font is read by the Symbol encoding.** The code (`F0xx`) was read as Latin-1, so the Symbol font's nabla printed as `Ñ` and its club as `§`. The table is generated from Adobe's `Symbol.afm` and the Adobe Glyph List; a code the font has no glyph for is the replacement character. Other fonts keep their reading.

- **A Word 6/95 `.doc` saved with a Macintosh character set is decoded as Mac Roman.** The reader chose the codepage from the language ID alone, so a curly quote (0xD5) or an accented letter came out as `Õ` or `Ž`. The marker is `FibBase.chs` = 0x0100. Word 97 and later files keep Windows-1252: they store compressed text that way even when they carry the marker.

- **A `.xls` formula cell whose text result follows its array, shared-formula or data-table record keeps that text.** The result was forgotten on any record other than `STRING`, so such a cell came out empty although the file holds the text.

- **A deflated package part whose length differs from the size its zip header declares is read and reported, instead of being dropped as unreadable.** A wrong CRC-32 was already tolerated; a complete deflate stream whose header length is wrong is the same kind of damage to the header, not to the content. The bytes are kept, `Metadata::warnings` names the part, and an entry that inflates more than a quarter past its declared size is still refused as a decompression bomb. A stream that ends short of the declared size is reported too.

- **A Markdown table no longer carries grid columns that only continue column spans.** A table whose rows each have their own cell boundaries (typical of Word 6/95 tables) was written as one column per boundary, mostly empty; a column in which no row starts a cell holds no text and is now left out.

- **A `.docx` or `.pptx` whose package-level or main-part relationships part the zip cannot decode opens again, with the loss reported.** A damaged `_rels/.rels`, `word/_rels/document.xml.rels` or `ppt/_rels/presentation.xml.rels` (an unsupported compression method, an encrypted entry) was read as an empty set of relationships and the document opened; an earlier change in this release made such a document refuse to open. It opens as it always did (the main part is found through `[Content_Types].xml`), and the part is listed in `unreadable_parts`, in the metadata warnings and in the text, with `text_truncated` set. Editing such a document is still refused, because saving it would write out a package without its relationships.

- **A damaged theme no longer adds a notice to a document's text or sets `text_truncated`.** A theme holds colours and fonts, no text, so the document yields every word it did before; the part is still listed in the metadata warnings (`.xlsx`, `.docx`, `.pptx`). Only a part the reader reached through a theme relationship is taken for one without text; a part's name never decides it, so a header stored under a name that looks like a theme is still reported as lost text.

- **The notices for unreadable parts in a spreadsheet keep bidirectional-override characters out of the text, as the other formats do, and a copy of an `.xlsx` document derives its date styles and number formats again rather than sharing the original's.**

- **A `.docx` or `.pptx` part that cannot be read is reported in the text, not only in the metadata.** A malformed header, footer, note or layout part left `plain_text`, Markdown, HTML and the IR byte-identical to a good file's, with the loss only in `Metadata::warnings`, while an `.xlsx` wrote a notice. The same bounded list (at most 100 parts, each cut at 256 characters, and the count of the rest) now ends the text of every rendering and the last IR section, charged to the text budget before it is written. Part names and reasons come from the file, so a line is one line with its control characters replaced, and Markdown escapes it (the IR's renderers escape it for HTML as for any text). A file that lost nothing is unchanged.

- **A `.docx` header, footer or picture part that relationships spell in different cases is read, parsed and reported once.** The per-part caches of headers and pictures compared names exactly while a zip lookup and a part name compare without regard to case, so `header1.xml` and `Header1.xml` were inflated twice and a readable header was kept as two entries. They now compare as a lookup does.

- **A relationship to something that is not there no longer reports a lost part.** A `.docx` numbering relationship to a part the package does not hold recorded a missing required part and marked the text truncated, where the theme, charts and pictures were silently absent. In a `.pptx`, a relationship marked external (a theme, layout, notes, comments, chart or SmartArt target) was resolved as a part name and could read a part of the package that its target happens to spell; it is now not a part.

- **A `.pptx` with no readable slide names a slide in its error.** The error named the first part recorded as unreadable, which is a theme or another part met before the slides when one failed, so a deck of broken slides blamed its theme.

- **Packages embedded in a `.docx` are bounded in depth, number and size, and no longer hang the parse.** An embedded OOXML package (`<o:OLEObject Type="Embed">`) was opened as a full document with fresh limits, and it could embed packages of its own: entries that name one blob made the work ten to the power of the depth from a few kilobytes (a 25 KB file embedding ten copies, five levels deep, took 64 s; ten copies of ten, four levels, over two minutes). Each nested open also waited for a parse slot, so a chain of embedded packages longer than the pool of slots, or as many documents embedding a package as the pool has slots, never finished. All of an embedded package's cost is now charged to the one document that holds it: it is opened on its own thread that never waits for a slot; packages nested deeper than `limits::max_embedded_depth` (default 3) or beyond `max_embedded_packages` (256) or `max_embedded_bytes` (512 MiB) are not opened, each is listed in `unreadable_parts` and `Metadata::warnings` with the reason, and the document is flagged `text_truncated`; a package reached through several entry names is opened once; the text of the embedded packages counts against the text budget and the bytes they inflate against their allowance. New limits: `limits::{max_embedded_depth, max_embedded_packages, max_embedded_bytes}` with setters and `DEFAULT_*` constants.

- **The `.docx` and `.pptx` report of unreadable parts tells parts apart by their names, not by a 64-bit hash.** A part is recorded once by its normalised name (case and `\` against `/` ignored), kept exactly, so a name crafted to collide with another part's can no longer hide that part's failure. The names remembered are bounded to 1 MiB; past it a part can be counted more than once among the rest, but is never taken for one already recorded.

- **A package part the zip cannot decode is reported as damaged, not taken for absent.** An entry with a compression method the reader does not implement, or an encrypted entry, was reported by the zip layer as a missing part, and the readers treat a missing optional part as a feature the file lacks: an `.xlsx` whose shared strings could not be decoded read as a workbook of empty cells with no warning and `text_truncated` unset. The package layer now returns a zip error for an entry that is in the package and cannot be decoded (`Error::MissingPart` is kept for a name the package does not have), so the part, with its reason, is in `unreadable_parts`, `Metadata::warnings` and `text_truncated` for `.xlsx` (shared strings, styles, theme, workbook relationships), `.xlsb`, `.docx` and `.pptx`; an unreadable `_rels/.rels` now fails the open instead of being read as empty.

- **The `.xlsx` notices for sheets and parts that could not be read echo names as inert text, and a streamed sheet keeps the warning of a notice the budget could not pay.** A sheet or part name from the file (which can hold newlines, `](...)`, `<img>` or a `# heading`) was written as it was in the notices of `plain_text` and `to_markdown`, and a sheet's name was the notice's section heading. Control characters become spaces, the Markdown renderers escape the lines, and the heading is the fixed `Unreadable sheet`; `Metadata::warnings` keep the bounded text unchanged. When the text budget cannot pay for a notice, `sheets_ir()` still hands the warning to the caller in `SheetIr::warnings` (the sheet's section may then be empty), and a notice is charged in characters.

- **An `.xlsx` relationship to a drawing, table, comments, pivot or chart part that is not in the package is silent again, and an external target is never looked up.** The unreadable-part report had begun to list a dangling optional reference as a damaged part (with `text_truncated` set). Only a part that is present and cannot be read or parsed is reported; a `TargetMode="External"` target is not read from the package even if a part of that name exists.

- **An `.xlsx` workbook of many sheets over many cell formats no longer costs their product.** The date-style set and the compiled custom number formats were rebuilt for every sheet by `plain_text`, `to_markdown`, `to_csv`, `sheet_ir` and the per-sheet renderers (3,000 sheets over 50,000 custom formats: `to_markdown` took 76 s), and the theme was re-parsed for every `sheet_ir` call. They are built once per document, on first use, and shared; output is unchanged.

- **The MCP server lowers the library's text and image budgets only while `extract` and `info` run, and says what its 32 MiB limit bounds.** The server used to lower them once at startup, for the whole process (the text budget to 32 Mi characters, the image budget to 24 MiB). They are now lowered for the length of an `extract` or `info` call (the two that convert a document) and restored afterwards, under a lock; `replace_text`, which reads neither, no longer runs under them. The 32 MiB response limit counts UTF-8 bytes (the text budget counts characters, so CJK text is refused at fewer characters than ASCII, with a message saying bytes). Only a spreadsheet is written as it is produced; every other format, and the IR, is still built whole before the response limit can refuse it, since the library has no streaming writer for them, so for those the memory bound is the two budgets, not the limit. A request line over 16 MiB is rejected with `-32600`. The README says all of this.

- **`.docx` `plain_text` now reports hyperlink targets cut while reading, as `to_markdown` and the IR do.** When composing link targets outgrew the text budget, `plain_text` ended without the notice the other renderings carry. (`.pptx` has no read-time link budget; its `plain_text` already ends with the copied-text notice.)

- **A `.docx` or `.pptx` part that cannot be read is tried once and reported once, and the report is bounded.** A header, footer or note that failed to read or parse was tried again by every section that named it and added another identical line to `unreadable_parts`: 20,000 sections naming one malformed 9.5 KB header took 9 s and listed the part 20,000 times. A chart, SmartArt, embedded-package or picture part, and a `.pptx` comments part, had the same shape. Each failure is now remembered as a success is, a part is listed once whatever its spelling, at most 100 are kept (the new `DocxDocument::unreadable_parts_omitted` and `PptxDocument::unreadable_parts_omitted` count the rest; `Metadata::warnings` ends with "and N more unreadable parts not listed" and `text_truncated` is set), and a name or reason is cut at 256 characters. A theme, numbering, alternate-content, embedded-font or picture part that cannot be read, a relationship target that cannot be resolved, a layout or master whose shapes do not parse, and a malformed `.pptx` comments or comment-authors part are now reported too, where they were dropped silently or only logged. Pictures named by a header, footer or note share one copy with the main document's.

- **An `.xlsx` hyperlink target is no longer dropped at 2,083 characters.** Links to longer targets (signed storage and sharing URLs) were dropped with a warning only, while `.docx` and `.pptx` accept a relationship target up to 64 KiB. The per-link limit and `Worksheet::hyperlinks_dropped` are gone; one threshold (64 KiB, in the package layer) applies to every format, and the IR's copies of a target are charged to the text budget.

- **An `.xlsx` relationships part that cannot be read is reported, not emptied.** A sheet, drawing, pivot table, chartsheet or rich-value relationships part that failed to parse (a target past 64 KiB, say) was replaced by an empty set, so the sheet silently lost its hyperlinks, pictures, comments, tables and pivots. The part and the reason are now in the new `XlsxDocument::unreadable_parts`, `Metadata::warnings`, and `Metadata::text_truncated`. Each part is reported once, at most 100 are kept (`unreadable_parts_omitted` counts the rest), and a part name or reason is cut at 256 characters.

- **`Document::to_ir` of an `.xlsx` no longer drops every sheet after the one that spent the padding allowance.** The conversion stopped at the first sheet to spend the workbook-wide allowance for empty padding cells, and the sheets after it were absent with only "later rows are missing" to say so. The sheet that spent it is cut with its notice, a later sheet that needs no padding is converted whole, and one that does is cut at the rows it cannot pay for, with a notice and a warning naming it. When the run's character budget ends the conversion, a section and a warning list the sheets left out by name.

- **A drawing, comments part, table, pivot table or chart that many `.xlsx` sheets name is parsed once, and each sheet's use of it is charged.** Every sheet that named the part read, parsed and copied it again: 300 sheets over one drawing of 20,000 anchors (250 KB) took 72 s and 1.9 GB to open, 300 over one comments part 25 s and 1.8 GB, 300 over one table definition 13 s and 1.5 GB. The parts are read once by the zip entry they resolve to and shared (`Worksheet::images`, `text_shapes` and `comments` are `Arc<[...]>`, `tables` and `pivot_tables` hold `Arc`s, `chart_text` holds `Arc<str>`); a sheet's picture alt text, text shapes, chart text and comment notes are charged to the text budget before the IR and the plain-text and Markdown renderers copy them, and a comment of one character costs the note that holds it. Pictures placed in cells are in the new `Worksheet::cell_images`, no longer mixed into `images`.

- **An `.xlsx` pivot cache definition that many pivot tables name is read and scanned once, and no package part can be inflated without bound.** Each pivot table part re-read and re-scanned the one definition it named (the package byte budget charged an entry only the first time), so 500 pivot tables over a 20 MB definition took about 30 s to open from a 314 KB file. The parsed definition (or the fact that it could not be read) is kept for every table that names it, as is the list of chart parts a drawing names for the chartsheets that display it, and the package's byte budget now also totals the bytes inflated by reading an entry again, or by a read that failed after inflating, against the same limit (`limits::set_max_package_bytes`).

- **An `.xlsx` threaded-comments part paired with many legacy comments parts is grouped once, and a merge made after the comment allowance is spent does no work.** Every sheet naming a new pair of parts grouped all of the threaded part's messages by cell before the first charge, so a few large threaded parts and a few hundred small legacy ones took minutes to open from a few megabytes. The messages are grouped once per threaded part, each thread's note, and each message as the string that holds it, is now charged, and a merge starts only while the allowance (`limits::set_max_text_chars`) lasts; a sheet merged after it is spent keeps no comments and is flagged `text_truncated`, as before.

- **An `.xlsx` report of what could not be read is bounded, charged to the text budget, and in every rendering.** The list of unreadable sheets had no cap, no name truncation and no charge: 6,000 `<sheet>` entries with 60 KB names naming one part were echoed in full by Markdown, HTML and the IR (about 1 GB from a 525 KB file). Unreadable sheets and unreadable parts now share one bounded list (the first 100 kept, each name and reason cut to 256 characters, the rest counted in the new `XlsxDocument::unreadable_sheets_omitted` and `unreadable_parts_omitted`), and every renderer writes it after charging it to the text budget. `plain_text`, `to_markdown`, `write_*`, the CLI and the streaming HTML/Markdown paths now also carry an in-band `[skipped unreadable part ...]` notice (previously only the IR metadata did), and the section for unreadable parts is a `SheetIr` whose `warnings` and `text_truncated` name them, so a caller that drops each `SheetIr` still learns of them. `Metadata::warnings` now also lists unreadable sheets.

- **An `.xlsx` part that cannot be read or parsed is reported, and the failure is remembered.** A comments, threaded-comments, table, pivot (table, cache definition or records), chart, drawing, picture or font part, and the styles, theme, shared-strings, persons, rich-value and workbook-relationship parts, that failed to read or parse was skipped with a debug log or nothing at all. Each is now recorded once (part name and reason, in `XlsxDocument::unreadable_parts`, `Metadata::warnings`, the text and `text_truncated`), and a picture that cannot be read is remembered, so a damaged image named by hundreds of anchors is no longer inflated again for each. A linked (external) picture target is no longer looked up in the package.

- **An `.xlsx` chartsheet's chart is no longer repeated among the embedded charts when its path is spelled differently.** The check compared the drawing's raw target with the package's entry names, so `..\Charts\CHART1.XML` and `xl/charts/chart1.xml` were two charts and the chart's text appeared twice. Parts are now compared by the entry they resolve to, as everywhere else in the reader.

- **An `.xlsx` picture that one budget refuses no longer spends the others.** The text, per-sheet picture and total picture budgets were charged one after another, so a picture refused by a later one left the earlier ones spent for a picture that was not produced, and later content was cut for it. A picture is now paid for from every budget or from none.

- **The allowances an `.xlsx` is converted under are now stated for each entry point.** `Document::to_markdown_with`, `to_html_with` and the `write_*_with` methods convert a sheet at a time with a sheet's own padding (`limits::max_sheet_ir_padding_bytes`) and picture (`limits::max_image_bytes`) allowances and document-wide text and total-picture budgets; `Document::to_ir` (and `to_ir().to_markdown_with`) holds every sheet and uses document-wide ones, so a workbook that fits sheet by sheet but not as a whole is rendered whole by the first and cut, with a notice, by the second. The streaming methods return no metadata; the cut sheets' `SheetIr::warnings` and `text_truncated` come from `XlsxDocument::sheets_ir`. Documentation only; behaviour is unchanged.

- **A `.pptx` style level's bullet or underline kind is no longer copied into every paragraph that inherits it.** A master or layout `buChar`, `buAutoNum` type or `defRPr u` of a megabyte was cloned into each inheriting paragraph or run (300 paragraphs inheriting 50 KB took 61 MB to open). A value longer than 64 characters, which no presentation application writes, is not inherited.

- **A `.pptx` part whose relationships cannot be read is now reported, and the IR metadata lists every skipped part.** A notes, layout or master part with an over-long relationship `Target` lost its links and pictures with no record, and a layout, master, picture or chart part that failed to decompress vanished silently. Each is now recorded in `PptxDocument::unreadable_parts`, `text_truncated` is set, and `Metadata::warnings` names the part and the reason.

- **A `.pptx` converted past the picture budget now says so in its content and flags the cut.** Pictures left without their bytes (`limits::set_max_image_bytes`) were reported in a metadata warning only; the IR now also sets `text_truncated` and closes the last section with a "[some pictures not shown ...]" paragraph, as the spreadsheet conversion does, so Markdown and HTML renderings show it.

- **Chart and SmartArt text in a `.pptx` is no longer copied for every frame or slide that names the part.** The lines of a chart or SmartArt data part were cloned into each slide's reference map, each graphic frame and the IR: a 13 KB deck of 3,000 frames naming one chart of 1 MB of text took 2.9 GB to open. The reader keeps one shared copy (`GraphicContent::Text` now holds an `Arc<[String]>`), and the IR, `plain_text` and Markdown charge each frame's copy to the text budget before making it, leaving the text out, flagged and noticed, once it is spent.

- **A `.docx` table style's font name is no longer kept for every table that uses it.** The conversion memoises the folded formatting per table-style layer for its whole run, and every table has a layer of its own, so a 1 MB face name in a table style and 300 tables held 640 MB at once. The kept folds are charged to the copied-text budget; past it they are kept without face names, tab stops and colour names, and the conversion reports it (`text_truncated`, a warning naming `limits::set_max_text_chars`).

- **A `.docx` chart or SmartArt part named by many drawings is no longer copied into each.** The part's text was cloned into every drawing at open, again into the IR for each, and written again by `plain_text`/`to_markdown`: a 1 MB chart named by 5,000 drawings was 5 GB at open. The reader now keeps one shared copy (`DrawingInfo::chart_text` and `dgm_text` are `Arc<[String]>`, a breaking change for code that builds or mutates them), and every copy into an IR or a rendering is charged to the copied-text budget first; past it the drawings are written without the text and the output reports it (`text_truncated`, a warning naming `limits::set_max_text_chars`; `plain_text` and `to_markdown` end with the same notice).

- **A `.docx` header, footer or note whose relationships part cannot be read is now reported.** A relationship target over 64 KiB (or any other error reading `word/_rels/headerN.xml.rels`, `footnotes.xml.rels`, ...) silently left the part without its images and links. The part is kept, and the relationships part is named in `unreadable_parts`, a `skipped unreadable part` warning and `text_truncated`, like every other unreadable part. The main document's own relationships still refuse the package, as a presentation's do.

- **A `.docx` whose pictures exceed the picture budget is now flagged and says so in the document.** Once `limits::set_max_image_bytes` was spent, later pictures lost their bytes with only a metadata warning: `text_truncated` stayed false and nothing in the content said so, so a Markdown or HTML rendering with embedded images dropped them silently. The IR now sets `text_truncated` and ends its last section with a "some pictures not shown" notice, as an `.xlsx` does.

- **A `.docx` style's tab stops are no longer copied, uncharged, into every paragraph that uses it.** The folded paragraph style (tab stops, border and shading names, the paragraph mark's run properties) was cloned for every paragraph, and the tab stops copied again into the IR: about 5,000 tab stops in one style and 5,000 styled paragraphs (a 16 KB file) took 1.4 GB of IR JSON. Each copy is charged to the copied-text budget first (a tab stop at its serialised size); once it is spent paragraphs keep their indents, spacing and numbering without the tab stops and the conversion reports it (`text_truncated`, a warning naming `limits::set_max_text_chars`). `to_markdown` folds a style chain without tab stops, which it does not read.

- **A `.docx` style's font name is no longer copied, uncharged, into every run that uses it, and Markdown no longer refolds it for every run.** A 200 KB `w:rFonts` name in a paragraph style and 5,000 paragraphs each with one bold run (a 5.2 KB file) took 22.8 s and 3 GB in `to_html`, and the IR JSON was 954 MB. The folded style was cloned for every run with direct formatting and the name copied into every span; the copy is now charged to the copied-text budget first, and once it is spent runs keep their formatting without the name and the conversion reports it (`text_truncated`, a warning naming `limits::set_max_text_chars`). `to_markdown` folds a style chain without face names, which it does not read.

- **A `.docx` numbering level's `w:lvlText` is no longer copied into every list that uses it uncharged.** The text was cloned into each list's marker pattern: a 200 KB `lvlText` and 5,000 one-item lists (a 6.5 KB file) took 1.9 GB. Each copy is charged to the copied-text budget first; once it is spent the lists keep their items and are numbered `%N.`, and the conversion reports it (`text_truncated`, a warning naming `limits::set_max_text_chars`).

- **A Markdown table of a realistic size no longer drops the rest of the document.** `DocumentIR::from_markdown` charged each table cell as four blocks of the 100,000-block limit, so about 25,000 cells (1,000 rows of 25 columns, or 4,200 of 6) used it up and everything after the table was lost. Table cells now have their own limit, `limits::set_max_markdown_table_cells` (default 150,000, about 180 MB of IR); past it only the table rows that do not fit are dropped, with a notice, and the rest of the document is read.

- **A conversion's image and copied-text budgets end with the conversion.** They were set at the start of a conversion and never cleared, and an image charge made outside any conversion was unlimited. Each is now held by a guard that restores the previous state when the conversion ends (including by an early return or a panic), and a charge made outside a conversion is held to the configured limit. Tests that change a process-wide limit now take a shared lock and restore the limit.

- **A `.ppt` slide whose container overlaps another is left out like an aliased one.** Slide-list entries naming a container that overlaps an earlier one stayed in the deck as empty slides while the notice said they were left out, and entries naming an already-listed slide were dropped, so later slides were numbered differently depending on why an entry was skipped. Both are now left out, and the notice counts them.

- **A multi-sheet `.xls` of wide, partly filled sheets is no longer cut by the document-wide padding allowance.** The 8 Mi allowance of empty slots that pad rows was shared by every sheet, so about four sheets of 65,000 rows by 60 columns at 60% fill exhausted it and later rows were dropped. Each sheet now has its own allowance (8 Mi, `limits::set_max_sheet_grid_padding_cells`) within a document total of 16 Mi (`limits::set_max_grid_padding_cells`, default raised from 8 Mi), so one sparse sheet cannot starve the rest and the worst case stays bounded. A sheet that is cut is named in the notice.

- **A `.doc` whose comment boundaries run backwards no longer copies the comment text once per range.** The `PlcfandTxt` boundaries were not checked to be non-decreasing, and each range was cut with a character scan from the start of the story: 1,000 ranges over a 100,000-character story (a 243 KB file) took 14 s and 50 MB. A `PlcfandTxt` whose boundaries decrease is now refused, with a warning, and the comments are shown as one block; the ranges of a valid one are cut in a single pass.

- **Compressed metafile pictures in a `.ppt` or `.doc` can no longer inflate to gigabytes.** A zlib-compressed EMF/WMF/PICT declares its uncompressed size, and each was inflated to it (up to 256 MiB apiece): four pictures in a 1 MB `.ppt` produced 3.3 GB from `to_ir`. A document's compressed metafiles now share an allowance of 256 MiB of inflated bytes (`limits::set_max_inflated_image_bytes`), the declared size is reserved before anything is inflated and the unused part is handed back, and a picture whose declared size exceeds 1,032 times its stored length (plus 1 MiB), the most a deflate stream can hold (RFC 1951), is refused outright. A picture that is refused is left out of the image list; genuine metafiles, flat-colour bitmaps included, are below both bounds.

- **A shared-formula budget failure no longer makes a whole `.xlsx` sheet unreadable, and no longer depends on thread scheduling.** A sheet whose shared formulas passed the document's text budget, or whose master was longer than 8,192 characters, was reported unreadable and lost every cell; with the `parallel` feature, which sheet ran out first depended on which thread charged first. Followers are now written after all sheets are parsed, one sheet at a time in workbook order, against one document-wide budget: the followers past it, and a refused master with its group, keep their cell and value and have no formula, counted in `xlsx::Worksheet::formulas_dropped`, flagged in `Metadata::text_truncated` and named in `Metadata::warnings`. Public types: `Worksheet` gains the public field `formulas_dropped`, and `xlsx::shared_formula::SharedFormulas` records followers (`follower` no longer returns the formula; `add_master` takes `&str` and returns whether it was accepted).

- **`to_ir`/`to_html` of a wide `.xlsx` are no longer cut where `to_markdown` is not, and the empty cells that pad a row are paid for at their size.** The IR charged each empty table cell that pads a row to the sheet's width as 64 characters of the text budget, Markdown charged 3 per absent column, so a wide multi-sheet export was complete in Markdown and truncated in the IR, with a notice blaming the character budget and no warning. The padding cells made (the grid's columns that no cell of the row occupies) are now charged 3 characters each, as Markdown charges them, from a character budget of their own at the configured text limit, and their memory (a `TableCell` and its element, about 600 bytes each) against a dedicated allowance, 512 MiB by default: `limits::set_max_ir_padding_bytes`, about 900,000 cells. When either is spent the IR stops at that row, with a notice and a `Metadata::warnings` entry that name the cause, and `Metadata::text_truncated` is set.

- **The text of a rich-text cell's runs is charged in full.** The charge for a rich-text cell's spans left out the runs' own text, and the displayed cell text it was counted against stops at 32 KB, so a shared string of runs holding far more could put more text in the IR than the budget allowed. Each run's text is now charged with its span.

- **The "pictures not shown" notice of an `.xlsx` is given once.** Once the picture budget (`limits::set_max_image_bytes`) was spent, the notice was appended to every later sheet, including sheets without pictures. It now appears on the sheet whose pictures spent the budget, and a `Metadata::warnings` entry names the limit.

- **A custom number-format code is no longer copied into every cell that uses it.** The code was cloned onto each numeric cell's IR entry: a 1 MB `formatCode` used by 300 cells in a 3 KB `.xlsx` took 297 MB in `to_ir`. A code longer than 255 characters, Excel's own limit, is now left off the IR cells (they keep `number_format_id` and their values) and a `Metadata::warnings` entry says so; ordinary codes are listed exactly as before. The same file also made the format parser overflow its counter of decimal placeholders (a panic in debug builds, a wrong count in release); it now saturates.

- **A shared string with many rich-text runs is no longer rebuilt, free of charge, in every cell that names it.** The runs were cloned into each referencing cell, and each became a span with its own font name: 20,000 runs named by 300 cells in a 12 KB `.xlsx` took 2.1 GB in `to_ir` against a budget that counted one character per run. The runs are now shared (`xlsx::shared_strings::SharedString::rich_text` and `xlsx::Cell::rich_runs` hold an `Arc<[RichTextRun]>`, were `Vec`), and the size of the spans a cell will get (the size of a `TextSpan` plus the font name, per run) is charged to the document's text budget (`limits::set_max_text_chars`) before they are made; the cell that does not fit ends the sheet, with the usual notice, warning and `Metadata::text_truncated`.

- **A spreadsheet hyperlink target is no longer copied for every link.** A relationship's target (a megabyte, in a hostile file) was cloned into each `<hyperlink>`, each cell and each IR span: a 5 KB `.xlsx` with 400 links to one 1 MB target took 1.5 GB in `to_ir`. A target longer than 2,083 characters, Excel's own hyperlink limit, is now refused: the links naming it are left out, counted in `Worksheet::hyperlinks_dropped` and named in `Metadata::warnings`, and the cells keep their values. The others are read once per relationship and shared; the copy each IR span owns (the target and hover text, once per rich-text run) is charged to the document's text budget (`limits::set_max_text_chars`) before it is made, and a budget that runs out now also adds a `Metadata::warnings` entry. Public types: `xlsx::HyperlinkTarget::External` holds an `Arc<str>` (was `String`) and `xlsx::Worksheet` gains the public field `hyperlinks_dropped`.

- **A `.docx` style sheet of tens of thousands of chained styles is cut at 20,000 styles.** Each style costs about a kilobyte for roughly a hundred bytes of XML: 60,000 chained `w:basedOn` styles (a 422 KB file) took 1.1 s and 210 MB to open. The chain itself was already followed only 20 links deep, so the cost was the count; the sheet now keeps the first 20,000 styles (about five times Word's own maximum), and `Metadata::warnings` says how many were left out. `StyleSheet` gains a `pub(crate)` field.

- **The default budget for picture bytes in an IR is 128 MiB, down from 1 GiB.** The IR gives each picture its own copy of the bytes, so a 5 KB `.docx` drawing one 1 MB image 1,100 times converted to 1 GB of IR and 1.4 GB of JSON. The budget (`limits::set_max_image_bytes`, shared by the `.docx`, `.pptx` and `.xlsx` converters, each resetting it per conversion) now stops the copies at 128 MiB; later pictures keep their place without bytes and the conversion adds the image warning. A document whose distinct pictures, counted once per use, exceed 128 MiB loses the bytes of the later ones: raise the budget for such documents. The IR JSON schema is unchanged, so identical image bytes are still written once per picture within the budget.

- **Pictures in a `.docx` header or footer shared by many sections are charged to the image budget, not the text budget.** A shared header's per-section IR copies were charged at their base64 JSON size to the text budget, so a 1 MB logo in a mail-merge document with about 190 sections exhausted it and later sections lost their headers, while body pictures used the image budget. The picture bytes of each copy now come off the image budget (`limits::set_max_image_bytes`) before the copy is made, and the rest of the copy (text, structure, link targets) off the text budget. When the image budget cannot take another copy of the pictures, the copy keeps its text and its pictures keep their place without bytes, and the image warning says so.

- **A `.docx` hyperlink's hover text is charged with its target.** Each IR span inside a hyperlink received its own copy of the link's `w:tooltip` without counting it: one 20 KB tooltip over forty runs made 800 KB of IR. The tooltip is now charged to the text budget together with the URL before either is copied; past the budget the text stays unlinked and the usual truncation flag and warning are set.

- **A picture's link target is charged where `.pptx` Markdown copies it.** `PptxDocument::to_markdown` wrote a linked picture's target once per picture without counting it, so the only bound was the slide's whole-output budget, which dropped the entire slide: forty pictures linked to one long target lost the slide's text too. The target is now charged to the text budget before it is copied; past it the pictures are named without their target, the rest of the slide stays, and the output ends with the truncation notice.

- **A `.pptx` OLE object's preview picture is shared and charged to the image budget.** Each `p:oleObj` frame copied its fallback picture's bytes at open and again into the IR, outside the image budget: 50 frames over one image made 50 copies and no warning. `OleObject::preview_data` is now an `Arc<[u8]>` shared with the picture part, and the IR image copy is charged like any picture's (past the budget the image keeps its place without bytes, and the usual warning is added).

- **A hyperlink target named by many `.pptx` runs is stored once and its copies are charged.** Every `a:hlinkClick` on a text run (and every linked picture or media reference) copied its relationship's target: a 2 KB deck with 300 runs linked to a 1 MB target took 296 MB to open and 586 MB to render. `pptx::HyperlinkTarget` (both variants), `PictureShape::link_target` and `MediaReference::target` are now `Arc<str>` shared by everything that names the relationship, and the copy each IR span or picture receives is charged to the text budget before it is made (past it the run keeps its text and loses the link, with `Metadata::text_truncated` and a warning). A slide whose relationships part cannot be read (for instance one with a `Target` over the 64 KiB cap) is now reported in `PptxDocument::unreadable_parts` instead of silently opening without its relationships.

- **A hyperlink target named by many `.docx` links is stored once, its copies are charged, and a relationship `Target` over 64 KiB is refused.** A relationship target is stored once, but every `w:hyperlink` (and every linked picture) that names it copied the whole text at open, again into the IR and again into Markdown: a 2 KB document with 400 links to a 1 MB target took 393 MB to open and 776 MB to render. `Relationship::target` is now an `Arc<str>`, `HyperlinkTarget::External` (docx) and `DrawingInfo::linked_image` hold the same `Arc<str>`, and links that name one relationship share one allocation. A copy that is still needed (a link with a `w:anchor` fragment joined onto its target, an internal target, the URL in each IR span, the URL in Markdown) is charged to the text budget (`limits::set_max_text_chars`) before it is made; past the budget the link keeps its text and loses its target (or its fragment), and `Metadata::text_truncated`, a warning and a Markdown notice say so. One budget now covers every copy an IR conversion makes out of shared document text (link targets, shared headers and footers, shared notes) instead of one per kind. A relationship `Target` longer than `limits`' 64 KiB cap (`core::relationships::MAX_RELATIONSHIP_TARGET_BYTES`; Word stops a hyperlink address near 2,000 characters) makes the relationships part, and so the package, fail with `Error::PackageLimit`. `DocxDocument` gains a `pub(crate)` field, so external struct literals of it no longer compile.

- **The MCP server no longer builds a result far larger than its 32 MiB limit before checking it, and no longer reads a request line without bound.** `extract` rendered the whole output (up to the library's 256 Mi-character text budget for a spreadsheet) before comparing it with the limit, and `info` converted the whole document to read its metadata. At startup the server now lowers the library's text budget to 32 Mi characters and its image budget to 24 MiB, so a larger spreadsheet is cut with the library's truncation notice (reported by `info` as a warning) instead of being built whole and then refused, and the IR `info` reads is bounded the same way. A request line longer than 16 MiB is rejected with a `-32600` error and skipped as it is read, instead of being buffered; a request line that is not valid UTF-8 now gets a `-32700` error instead of ending the server.

- **`DocumentIR::from_markdown` is bounded in the blocks it builds.** A few bytes of Markdown become 0.5-1.4 KB of IR per section, paragraph or table cell: 100,000 sections were 214 MB, 200,000 paragraphs 285 MB and one table row of 100,000 columns 490 MB. The parser now charges every block (a section, paragraph, heading, code block, break, list or table at 1, a table cell at 4) against `limits::set_max_markdown_elements` (default 100,000) before it is built; the block that does not fit and the rest of the document are not read, a notice paragraph ends the last section, and `Metadata::text_truncated` and a warning say so. Documents within the limit parse exactly as before; list items keep their own limit.

- **`.docx` `replace_text` is linear in a paragraph's runs and matches.** For every match the replacement code walked all of the paragraph's text nodes to copy the unchanged text and searched them again for the node holding the match start: one paragraph of 20,000 one-character runs with a match in each took 2.7 s, and 80,000 took 62.6 s. Both walks now move forward over the sorted matches and node spans only. The edited XML is unchanged, checked against the previous code on varied run splits. (The `.pptx` editor already works one text element at a time and was not affected.)

- **Markdown inline parsing is linear in the line length.** Every `[`, `**`, `__`, `*`, `_`, `~~` or backtick searched the rest of the line for its closer, so a line of unclosed openers took quadratic time: `[` repeated 195,000 times took 6.4 s and `[](` repeated 195,000 times took 44 s in `DocumentIR::from_markdown`. Each marker now remembers its last search (a closer found at some position answers every later search that starts before it, and "none left" answers all later ones), so each stretch of the line is scanned once per marker. The parsed output is unchanged, checked against the previous parser on every short string of the characters it reacts to.

- **Writing a workbook with thousands of sections no longer takes super-quadratic time.** Each new worksheet copied every existing sheet name and probed every numeric suffix against all of them: 48 KB of Markdown `---` sections took 33.6 s to write as `.xlsx`. The workbook now keeps its names in a set and remembers the next free suffix per base name, so adding a sheet costs time in the length of its name; the names chosen are the same. Past 998 sheets with one name the old code repeated the name (a workbook Excel refuses); the suffixes now keep counting. The number of sections written to `.xlsx` is bounded by `limits::set_max_write_sheets` (default 32,767, so the package stays under the 65,535 entries of a non-ZIP64 zip); an IR with more is refused with `PackageLimit` before any sheet is built, and `XlsxWriter` itself stops adding sheets at the limit (the sheet that does not fit is not kept, `add_sheet_get_index` returns an out-of-range index) and `save`/`write_to` return the same `PackageLimit` error. Sheet names that differ only after the 27th character, and so share their suffixed candidates, share their suffix progress too.

- **Compressed metafile pictures in a `.ppt` or `.doc` can no longer inflate to gigabytes.** A zlib-compressed EMF/WMF/PICT declares its uncompressed size, and each was inflated to it (up to 256 MiB apiece): four pictures in a 1 MB `.ppt` produced 3.3 GB from `to_ir`. A document's compressed metafiles now share an allowance of 256 MiB of inflated bytes (`limits::set_max_inflated_image_bytes`), the declared size is reserved before anything is inflated and the unused part is handed back, and a picture whose declared size exceeds 100 times its stored length (plus 1 MiB) is refused outright. A picture that is refused is left out of the image list and logged; genuine metafiles are far below both bounds.

- **A `.ppt` text run with a non-BMP character and many style spans or hyperlink ranges no longer takes quadratic time.** Style-run counts and hyperlink offsets are UTF-16 code units, and each one was converted to a character index by scanning the run's text from the start: a 1.2 MB deck took 17 s. The conversion now builds one map of the run's non-BMP characters and answers each offset by binary search.

- **A `.doc` with thousands of sections no longer multiplies its header and footer text.** Each section received its own merged copy of the header and footer stories, and the IR copied them again: a 60 KB file produced 3.7 MB of Markdown. Inherited stories are now shared rather than cloned, a story is sliced out of the header document through a character-offset table built once instead of by scanning from the start, and the text of each story read and of each per-section copy in the IR is charged to the document's text budget (`limits::set_max_text_chars`) before it is made. The section that does not fit, and those after it, are left without headers or footers, the file is reported truncated (`Metadata::text_truncated`) and a warning names the limit.

- **A `.ppt` hyperlink target and placeholder role are no longer copied onto every text run for free.** A shape's hyperlink target (kilobytes, in a hostile file) was cloned onto each text run in the shape, and onto each hyperlink range, without being counted; a cloned outline run also copied its formatting spans and link ranges uncounted. The copies are now charged to the document's text budget (`limits::set_max_text_chars`) at their real size before they are made, and the shape walk stops with the budget as it does for text.

- **A `.ppt` whose persist directory points slide entries into one another no longer re-reads the same bytes once per entry.** Distinct slide or notes offsets can nest inside one container that covers most of the stream, and each was parsed and copied in full whatever the text budget. A `Slide` or `Notes` container that overlaps one already read is now left out and counted in `Metadata::warnings` (the file is reported truncated), so the work of reading all of them is bounded by the stream's length; a genuine deck's containers are disjoint.

- **A `.doc` whose embedded-object directory loops no longer takes minutes to open.** The `ObjectPool` walk had no visited set, only a step cap, so storages whose sibling pointers named themselves were listed 50,000 times and each listing walked its own children 50,000 times: a 5 KB file took 49.8 s. The walk now visits each directory entry once, as the stream lookup already did, and classifies a child tree shared by several objects once.

- **A `.ppt` slide list whose entries all name one slide no longer copies that slide once per entry.** Repeating one persist id was already collapsed, but entries with different persist ids that the persist directory maps to the same `Slide` container each re-read it: 10,000 entries over a 100,000-character slide were 1 GB of text from a 395 KB file. A slide (or notes page) is now read once however many entries name it; the later entries are left out and counted in `Metadata::warnings`, and the file is reported truncated (`Metadata::text_truncated`). The per-document text budget (`limits::set_max_text_chars`) now also covers slide text, including a notes page copied onto each slide that names it: the slide that does not fit, and those after it, are left out with a warning naming the limit. Within a slide, an `OutlineTextRefAtom` copies the run it names, so thousands of references to one stored run multiplied it; each copy is now charged before it is made and the walk stops when the budget is spent. Ordinary text atoms and the deck's header/footer text copied onto each slide are charged the same way, and list entries whose persist id resolves to no slide are collapsed to one placeholder instead of one empty slide each.

- **A `.doc` piece table can no longer multiply a small file into hundreds of megabytes of text.** Pieces select ranges of the `WordDocument` stream, and nothing checked that they select different ones: a 157 KB file with a few thousand pieces over the same paragraph, and `ccpText` raised to cover them, decoded to 180 MB. The reader now sums what the pieces read and refuses the file (`Err`) when that exceeds 16 times the stream's size plus 1 MiB; real fast-saved files share only a few bytes between pieces. Pieces whose bytes lie outside the stream read nothing and cost nothing, as before. The per-document text budget (`limits::set_max_text_chars`), until now spent only by the spreadsheet readers, also covers the decoded text of a `.doc`, Word 6/95 files included: past it the piece table is cut where the budget runs out, the file is reported truncated (`Metadata::text_truncated`) and a warning names the limit.

- **A sparse `.xls` sheet no longer multiplies its file size into gigabytes of padding.** Each sheet row is stored up to its last used column, so one number in column 255 of each of 65,535 rows is 255 empty slots per record: a 7 MB workbook of six such sheets allocated 2.6 GB and rendered 101 MB of separators. The empty slots that pad rows out to their last cell are now charged against a per-document allowance, 8 Mi slots by default (`limits::set_max_grid_padding_cells`; it was 32 Mi, which still let the 7 MB reproducer reach 0.9 GB on open and 2.1 GB in `to_ir`); cells themselves cost nothing, so a dense 65,536 x 256 sheet still reads in full, as does a sheet with a cell in the last column of up to about 32,000 rows. The row that does not fit, and every row after it in the document, is left out and the workbook is flagged truncated (`Metadata::text_truncated`, with a notice in the text output), as when the text budget runs out. The BIFF2-4 reader shares the bound.

- **Many `<sheet>` entries naming one part no longer read and parse it once each, in `.xlsx` and `.xlsb`.** Each entry re-read the worksheet part, copied it and parsed it again: 20 entries over an 849 KB part took 10.4 s and 1.8 GB, and 200 were killed for memory. A sheet is a part of its own ([ECMA-376] §18.2.19), so the first entry owns the part and every other entry that resolves to it is reported as an unreadable sheet (in every rendering and in `Metadata::text_truncated`), not parsed. A part is identified by the zip entry it resolves to, so differently spelled paths (case, `\`, `..`, a missing part that falls back to a positional name) are one part. Workbooks whose sheets have parts of their own are unaffected.

- **One pivot cache named by many pivot tables in an `.xlsx` is shared and its output is charged.** Every `pivotTable` relationship to the same pivot part copied the whole cached source data, and `plain_text()`, `to_markdown()` and `to_ir()` wrote the table out once per pivot: 100 relationships over a 50,000-record cache rendered 513 MB of Markdown from an 8 KB file. `SheetPivotTable::cache_data` is now an `Arc<PivotCacheData>` shared by the pivots that name the cache part, and the renderers and the IR conversion charge each pivot's table to the text budget, ending with the usual truncation notice. Workbooks with several pivots, or several pivots over one cache, within the budget are unaffected.

- **A comment author named by many `.xlsx` comments is stored once and charged where it is rendered.** Every comment held its own copy of its author's name, and `plain_text()`, `to_markdown()` and `to_ir()` wrote it again for each: a 4 KB workbook with a 1 MB author and 5,000 comments took 7.3 s and 5 GB to open. `SheetComment::author` is now an `Arc<str>` shared by the comments of one author (and one threaded-comment person), and the renderers and the IR conversion charge each comment (cell reference, author, text) to the text budget, ending with the usual truncation notice. A threaded-comment thread repeats the author on every message, so its text is charged to a workbook-wide budget; threads past it are left out (their cell keeps its legacy note), the sheet keeps its cells, and the cut is stated in every rendering and in `Metadata::text_truncated`. A few comments are unaffected.

- **One picture drawn from many anchors in an `.xlsx` is no longer copied once per anchor.** Every `<xdr:pic>` (and every in-cell rich-value image) got its own copy of the media part's bytes, read and decompressed again each time: 100 anchors of a 20 MB picture took 4.7 s and 2.8 GB to open and 4.8 GB to convert to the IR. `WorksheetPicture::data` is now an `Arc<[u8]>` shared by every anchor naming the part, and each media part is read once. The IR still gives each image its own bytes, so `to_ir()` counts every picture against a new budget, `limits::set_max_image_bytes` (default 1 GiB, a picture counted once per anchor); pictures past it are left out, a paragraph says so, and `Metadata::text_truncated` is set. Workbooks with many different pictures, or a repeated one within the budget, are unaffected.

- **Padding a sparse `.xlsx` sheet out to its full width is charged to the text budget.** Every row is laid out as wide as the sheet's widest cell, so a sheet with cells at `A1` and `XFD1` and thousands of rows between became 16,384 columns per row: 937 MB of Markdown from an 86 KB file, and 5 GB of empty table cells in `to_ir()` and `to_html()` from a 3 KB file. The Markdown renderer now charges the unused columns of each row (as `to_csv` already did), and the IR conversion charges each empty padding cell, so the table ends with the usual truncation notice and `Metadata::text_truncated` is set. A sheet whose rows are filled is not charged for anything.

- **A shared formula copied into many `.xlsx` cells is bounded.** Every follower of a shared-formula group (`<f t="shared" si="…"/>`) held its own translated copy of the master's text, and the translation ran at a few MB/s: a 15 KB workbook whose 900 KB master was shared by 2,000 cells took 234 s and 3.1 GB to open. A master longer than 8,192 characters, Excel's own formula limit (ECMA-376 Part 1 §18.17), is now refused, and the translated followers of the whole workbook are charged to one document-wide text budget (`limits::set_max_text_chars`). A sheet that goes over is unreadable, reported like any other unreadable sheet, and the workbook is refused when no sheet is readable. Ordinary shared formulas are unaffected.

- **A `.pptx` slide list that names one slide part many times keeps one slide.** `p:sldIdLst` is an ordered list of the presentation's slides and a slide is one part, but each entry re-read, re-parsed and stored the part, and every renderer emitted it again: a 11 KB deck listing a 1 MB slide 2,000 times opened as 1.9 GB and rendered 1.9 GB per output. A part is now kept at its first entry; each repeat is reported in `PptxDocument::unreadable_parts`. A slide list longer than the package entry limit (`limits::max_package_entries`, 100,000 by default) is refused with `Error::PackageLimit`, since it can only repeat itself.

- **Image bytes are shared by the pictures that use one image part, and the copies in an IR are budgeted.** One image relationship used by thousands of pictures cloned the whole image into each (an 8 KB `.docx` with 2,000 references to a 1 MB image converted to a 2.5 GB IR). The readers now hold each image part once: `DocxDocument::images` maps to `Arc<[u8]>` and `PictureShape::data` is `Option<Arc<[u8]>>` (public type changes; `as_deref()` and slicing work as before). The IR keeps owning its bytes, so its serde form, the Python/WASM/C surfaces and the CLI are unchanged for any document within budget; the bytes copied into one IR are charged to a new per-document budget, `limits::set_max_image_bytes` (1 GiB by default). When it is spent, further pictures keep their place, size and alt text but carry no bytes, and `Metadata::warnings` says so.

- **Markdown with a very large number of list items no longer multiplies into gigabytes of IR.** Each `- x` line (four bytes) became about 250 bytes of IR, and a 3.8 MB file of them took 3.1 GB to parse and 7.5 s to write. `DocumentIR::from_markdown` now keeps at most `limits::max_markdown_list_items` list items (100,000 by default, nested ones included; `limits::set_max_markdown_list_items` changes it); the rest are dropped, a notice paragraph ends the last section, and `Metadata::text_truncated` and a warning say so.

- **Writing an IR with a very large table to `.docx` or `.pptx` is refused instead of exhausting memory.** The writers build every cell in memory (a one-million-cell table took 2.2 GB and 4 s to write). `create_from_ir` and `create_from_ir_to_writer` now count the cells the writer would build for the tables of the IR (rows times the padded width, as the `.pptx` writer pads every row to the widest and the `.docx` writer lays out the grid its column spans widen, clamped to its 4,096-column limit; nested tables, list items, text boxes, notes, headers and footers included) and return `Error::PackageLimit` past `limits::max_write_table_cells` (250,000 by default; `limits::set_max_write_table_cells` changes it). Spreadsheets are not affected.

- **The per-section header/footer and per-slide notes copies in an IR are charged by their real size.** The budget for these copies counted only visible text, so a header or notes part made of link targets, tooltips, font names, empty elements or images cost almost nothing while every section or slide still received a full copy. A copy is now charged at its serialised size, which holds every string, the structure and any image bytes.

- **A `.pptx` notes slide related to by many slides is read, parsed and held once.** Each slide re-read and re-parsed the shared part and kept its own copy (1,000 slides over a 1 MB notes part opened as 975 MB). The part is now cached by name and `Slide::notes` is an `Arc<TextBody>` shared by those slides (public type change: `Option<TextBody>` became `Option<Arc<TextBody>>`; reads through `Deref` are unchanged). The IR converts the shared body once; the per-slide copies in `Section::speaker_notes`, and the text of the direct `plain_text()`/`to_markdown()` renderers, are charged to the text budget, and when it is spent output stops with the budget notice (the IR sets `Metadata::text_truncated` and a warning).

- **A `.docx` header or footer shared by many sections is parsed and stored once.** Every `w:sectPr` reference re-read, re-parsed and pushed a new copy of the part: an 8 KB document with 2,000 sections over one 1 MB header opened as 3.8 GB. `DocxDocument::headers_footers` now holds one entry per distinct part, type and role (so `plain_text()` of a document whose sections share a header prints it once, as `to_markdown()` already did), and each section's references are resolved through an index, which also keeps the IR's per-section slots right when a header part is unreadable. The IR still gives every section its header (its type has no sharing), but the copies are charged to the per-document text budget (`limits::set_max_text_chars`); when it is spent, further sections get no header, and `Metadata::text_truncated` and a warning say so. `DocxDocument` gains a `pub(crate)` field, `hf_by_ref`, so external struct literals of it no longer compile.

- **A continuation of a vertical merge in a `.doc` table must cover the same columns as the cell it continues, and the text it holds is kept ([#474](https://github.com/yfedoseev/office_oxide/issues/474)).** A `fvmMerge` cell was paired with the merge above it by its left boundary alone, so a cell narrower than the merged one was absorbed into it and the row came out wider than the grid. It now continues a merge only when its left and right boundaries both match, each within the 4-twip tolerance the column grid uses, measured from the cell that started the merge; otherwise it stays an ordinary cell, which keeps the row on the grid (LibreOffice's importer instead accepts any continuation that lies inside the merged cell). A continuation that holds text (Word writes them empty, but a damaged file can not) had that text dropped silently; it is now added to the merged cell, in row order. Empty continuations are still absorbed without a trace.

- **A vertically merged cell in a `.doc` table is matched with the cell above it by column, not by cell index ([#472](https://github.com/yfedoseev/office_oxide/issues/472)).** When the rows above and below had different numbers of cells (typically horizontally merged cells above), the continuation was absorbed but credited to no cell: the merged cell kept a `row_span` of 1 and the row below came out one cell short. A continuation now joins the merge above it by column position (the exact rule is in the next entry), and one with nothing above it stays an ordinary cell instead of being dropped. A continuation is paired with its merge by position, as LibreOffice's importer, the `.docx` conversion and python-docx do.

- **Mapping a `.doc` property run to the pieces it overlaps no longer does runs × pieces work.** A fast-saved file can point any number of pieces at the same bytes, and the search for a run's pieces revisited every small piece inside a long one: a crafted 254 KB Word 6/95 file took 6 s and 700 MB, and Word 97 files shared the code. The pieces a run overlaps are now found in time proportional to the overlaps, and a file whose overlaps total more than 16 per run and piece (plus an allowance) is read as text without paragraph structure or, for character runs, without character formatting or the deleted/hidden-text exclusion; each case is reported in `Metadata::warnings`. Overlap alone is never refused, since real fast-saved files share bytes between pieces. A Word 6/95 file with one long piece holding many small ones, which was refused, is now read with its tables. Table column spans count the grid edges in a cell's interval by binary search instead of scanning the whole grid for every cell (100,000 wide paragraphs took 42 s).

- **Text in a table row-end paragraph is no longer dropped.** A row-end paragraph is empty in a well-formed file, and the table walk discarded its text. A damaged property list can mark a paragraph that holds text as the row end; that text is now the content of a cell instead of disappearing.

- **A `.doc` table is no longer cut in two at a row whose properties live in the `Data` stream ([#456](https://github.com/yfedoseev/office_oxide/issues/456); reported by @xugangqiang).** When a paragraph's property list is too large for its PAPX, Word stores a `sprmPHugePapx` ([MS-DOC] §2.6.2) that names a `PrcData` in the `Data` stream. The reader never followed it, so a row mark carrying a large table definition lost its in-table and row-end flags and its row definition: the table ended at that row, the row became a blank paragraph, and the rest started a second table. The reference is now expanded in place, one level deep and within `PrcData`'s size limit, under both opcode numbers readers use for it (`0x6646`, and `0x6645` from Word 97, which Apache POI and LibreOffice also accept). The total copied out of the `Data` stream is budgeted at twice its size plus 1 MiB, every copy being charged, so a crafted file cannot multiply its size into gigabytes (a 94 KB file did that before the budget); paragraphs with byte-for-byte the same property list (the row ends of a long table) share one expanded copy and spend the budget once, while lists that merely name the same `PrcData` are charged each, and a run's property list is decoded once however many paragraph marks it covers. A reference that cannot be resolved (outside the stream, negative or oversized length, or over budget) is reported in `Metadata::warnings` instead of being dropped silently.
- **`.doc` header rows now set `TableRow::is_header` and `repeat_as_header`, as `.docx` does.** `sprmTTableHeader` ([MS-DOC] §2.6.3) was not decoded at all. A nested table's header flag is not carried over: nested rows are flattened into the outer table, where the flag would mark an arbitrary row.

  Measured against v0.1.15 on the 712 unique `.doc`/`.dot` files under `~/projects`: 673 are byte-identical in both the IR and the text output, 39 change, none gains an error, and none loses or gains a word, a row, or a character of text. The 39 are all tables: for example 307 tables become 5, 161 become 6, and 134 become 33. On those files the table count and the header-row count match Apache POI's HWPF (`Table`, `TableRow.isTableHeader()`) in every file POI opens, except for table counts in four: three where this fix changes nothing (nested-table handling) and one that moves toward POI's count.

- **A `.doc` table that contains a table keeps it inside the cell that holds it ([#465](https://github.com/yfedoseev/office_oxide/issues/465)).** A paragraph's `sprmPItap` is its table nesting depth, and the cells of a depth-2 table sit inside a cell of the depth-1 table. The reader built one flat table run, so a nested table's rows became rows of the outer table and the outer row that held it was cut apart: a 3-row table with a 2-row table in one cell came out as 5 rows, with the inner rows promoted. The walk now keeps one builder per open depth, and a table that closes becomes a block of its parent cell, so the outer table keeps its rows and cells and the nested `Table` is in that cell's content. The enclosing tables are opened when a nested table is the first thing in a cell.

  Word ends a cell of a nested table with an ordinary paragraph mark flagged `sprmPFInnerTableCell` (0x244B), and its row with `sprmPFInnerTtp` (0x244C), not with the `\x07` and `sprmPFTtp` of an outermost table; both are now decoded, so nested rows keep their cells (including empty ones) and their row definition. `sprmPItap` is read as the signed 4-byte integer it is and saturated, where it used to be truncated to a byte (257 became depth 1). A nested table's header-row flag now belongs to the nested table. Nesting deeper than 15 levels is folded into the 15th and reported once in `Metadata::warnings`; the earlier warning that nested tables were flattened is gone with the flattening.

  Because the tree now holds tables in cells, the writers and renderers that only looked at a cell's paragraphs were fixed to keep that text: converting to `.xlsx` or `.pptx` writes a nested table's cell text into the cell (it was dropped), and Markdown writes a nested table's rows as lines of the outer cell, its cells separated by an escaped pipe, instead of adding columns to the outer table. The nested text is kept whether or not the outer cell carries formatting (a bold run turns a spreadsheet cell into a rich string, which skipped it), is kept apart from the outer cell's own text in `.xlsx` and `.pptx`, and a nested Markdown row of `-` placeholders or with an empty cell keeps its place. A hyperlink inside a nested cell is not carried to `.xlsx` or `.pptx` (the text is). A negative `sprmPItap` is reported like one beyond the supported depth instead of passing for an outermost paragraph. `DocumentIR::plain_text()` of a `.doc` with a nested table now renders it as `.docx` already did.

  Measured against the build just before it on 712 unique `.doc`/`.dot` files: all 712 are identical in the text output; the IR changes in exactly the 20 files that carried the flatten warning (58 tables), and in no file is a word lost or gained.

- **A sparse `.xls` sheet that spends the padding allowance no longer empties later sheets that need no padding.** Once the per-document allowance of empty padding slots (`limits::set_max_grid_padding_cells`) was spent, even a row that needs no padding was refused, so every later dense or single-column sheet lost all its rows. A row that needs no padding now always fits; only a row that does need padding is left out (with the usual notice and `Metadata::text_truncated`).

- **A compressed `.ppt`/`.doc` picture left out for a limit is now reported, and a flat-colour bitmap is no longer refused for its ratio.** The declared-size guard was 100:1 (plus 1 MiB), which refused legitimate flat-colour bitmaps (a 1920 x 1080 x 32-bit frame is 8.3 MB stored in about 8 KB, near 1000:1); it is now the deflate format's own bound, 1,032:1, with the per-picture cap and the per-document allowance doing the bounding. A picture refused for its declared size, or for not fitting the document's inflation allowance (`limits::set_max_inflated_image_bytes`), was only logged; it is now counted in `Metadata::warnings` (naming the limit) and sets `Metadata::text_truncated` when the document is converted to the IR.

- **Which of several overlapping `.ppt` slide or notes containers is read no longer depends on the order the list names them.** A container that overlaps one already read was left out, so a nested container listed before the one holding it was kept and the outer one dropped. The containers of the slide and notes lists are now sorted by position before any is read and the outermost of each overlapping group wins, whatever the list order; the notice and `Metadata::text_truncated` are unchanged. A genuine deck's containers are disjoint and unaffected.

- **A `.ppt` read without a usable persist directory spends the document's text budget once, not once per container.** The fallback readers (every `Slide` container in the stream, the stray-text scan, and the masters' static text) each started with a fresh full budget, so a deck of thousands of slide containers could produce thousands of times `limits::set_max_text_chars`. They now draw on the one budget the deck is read under, share what the persist path already spent, and stop listing slides when it runs out, with the usual notice in `Metadata::warnings` and `Metadata::text_truncated`.

- **A `.doc` is no longer charged twice for the same header and footer text.** The piece table charged every decoded character to the text budget (`limits::set_max_text_chars`), the header/footer stories were charged again when sliced out of the header document, and again for the first IR copy of each, so a file whose text was within a few percent of the limit lost its headers and footers. The first copy of each story is now covered by the piece-table charge; only copies beyond it (stories whose ranges overlap, and the repeated per-section copies of a shared header) are charged, before they are made. The Word 6.0/95 reader also hands its spent budget on instead of starting the IR conversion with a fresh one.

### Added

- **Tables in Word 6.0/95 `.doc` files ([#466](https://github.com/yfedoseev/office_oxide/issues/466)).** The Word 6/95 reader took only the text: it read no paragraph properties, so the cell and row marks of a table came out as tabs and line breaks inside ordinary paragraphs and every table was lost, with no warning. It now reads the PAPX that carries them. Word 6/95 stores paragraph properties in the `WordDocument` stream in an encoding of its own: a bin table of 2-byte page numbers, FKP pages of 7-byte entries, and one-byte sprm numbers whose operand sizes come from a table (`sprmPFInTable` 24, `sprmPFTtp` 25, `sprmTTableHeader` 186, `sprmTDefTable` 190). The table sprms are re-encoded as the Word 97 ones and go through the same table walk, so a Word 6/95 table has rows, cells, column spans from its row definition, and header rows. Only a file that contains a table is read this way; every other Word 6/95 file keeps the text-only view it had. The bin table can list fewer pages than the FIB says exist (`cpnBtePap`); the rest are read from `pnPapFirst` on. Section breaks end their paragraph through the section table. Text no property run covers keeps the default properties, and runs are clipped to a disjoint set before the paragraph builder sees them, so a hostile file cannot make it multiply runs against pieces; a fast-saved file whose pieces would make the paragraph builder do more than a bounded amount of work per run and piece (pieces piled on the same bytes, or one long piece with many small ones inside it; real files share a byte or two) is not walked paragraph by paragraph, and says so. The translated property lists are budgeted against the size of the file. An unreadable bin table, property list or row definition, and anything over budget, is reported in `Metadata::warnings`.

  Measured against the build just before it on 804 unique `.doc`/`.dot` files, of which 38 are Word 6/95: all 766 others are byte-identical in the IR and the text, and the text output of all 38 is identical. The table-bearing Word 6/95 files gain their tables (nine of them hold 93 tables per LibreOffice), with the same table count and rows per table as LibreOffice in 7 of the 9 it opens; the other 2 are damaged fuzz-corpus files. The sprm size table was checked against 14,393 property lists in the 32 real Word 6/95 files of the corpus, 99.97% of which it walks to exactly their end.

- **`Paragraph::style`: the paragraph style a document applies ([#455](https://github.com/yfedoseev/office_oxide/issues/455); requested by @xugangqiang, with the legislation use case from @nils-blockkin).** A paragraph in Word's Caption style, or in a legislation style such as `section-e`, used to come out as a plain paragraph in the IR, so a consumer could not tell it from body text. The IR now carries the style's display name (`"style": "Caption"` in the JSON, omitted when absent) for body, table-cell and list-item paragraphs of `.docx` and `.doc`; `None` means the default style or no known name. Names follow python-docx: DOCX shows Word's lower-case built-ins capitalized (`caption` → `Caption`), and an id the style sheet does not define, or one naming a character, table or numbering style, is the default style (in `.doc` the style kind is `StdfBase.stk`). In `.doc`, Heading 1–9, Header, Footer and Caption are reported under those English names whatever language the file was saved in; other styles keep the stored name, which may be localized. The default style is recognised by `w:default` / `istd` 0, never by its name. A name that is blank, longer than 255 characters (Word's limit) or holds a control character is not exposed and is reported once in `Metadata::warnings`; the cap also stops one huge style name from being copied into every paragraph that uses it.

  **Source compatibility:** `ir::Paragraph` gained a public field and is not `#[non_exhaustive]`, so a struct literal that names every field no longer compiles; build it with `..Default::default()` (the pdf_oxide crate already does). `docx::styles::Style` and `docx::styles::StyleSheet` also gained non-public fields (the `w:default` flag, the resolved display name, a count of unusable names); neither is `#[non_exhaustive]`, so a struct literal of either from outside the crate no longer compiles; values parsed from a document are unaffected. The resolved display name is computed when the style sheet is parsed, so it does not follow later edits to the public `styles` map. The JSON is compatible both ways. The DOCX and PPTX writers do not emit the style.

  Measured against the unmodified build on 10,060 unique `.doc`/`.docx`/`.xls`/`.xlsx`/`.ppt`/`.pptx` files: with the new key removed from paragraphs the IR is identical for every file that opens (9,736 files, and 313 that fail identically in both builds).

- **A spreadsheet can be converted to the IR one sheet at a time.** `xlsx::XlsxDocument::sheets_ir()` returns an iterator of `xlsx::SheetIr` (the section, the warnings its conversion raised and whether it was cut), in the order of `to_ir().sections`: each worksheet, the chart text, the pivot caches, the sheets that could not be read. Nothing is converted until the iterator is advanced and nothing handed out is kept, so a caller that uses and drops each section holds about one sheet's IR however large the workbook is (one worksheet at a time, as a streaming reader does); `SheetIrs::metadata()` reports the properties, warnings and truncation flag as `to_ir().metadata` does. `XlsxDocument::sheet_ir(index)` converts one sheet. New public types: `xlsx::SheetIr`, `xlsx::SheetIrs`. New limits: `limits::set_max_sheet_ir_padding_bytes` (a sheet's allowance for the empty cells that pad its rows, 1 GiB by default; a sheet past it is cut with a notice naming it, a warning and `SheetIr::text_truncated`, and the next sheet is converted whole) and `limits::set_max_total_image_bytes` (the picture bytes all the sheets of one run may copy, 4 GiB; each sheet's own allowance stays `limits::set_max_image_bytes`).

- **Renderings can be written to a stream as they are produced.** `Document::write_plain_text`, `write_markdown`, `write_markdown_with`, `write_html` and `write_html_with` (and `xlsx::XlsxDocument::write_plain_text`, `write_markdown`) take a `std::io::Write`. For an `.xlsx` the output goes to the writer a sheet, row or section at a time and is never held whole; the other formats build the string first and write it. Their output is byte-for-byte that of the `String` methods.

### Changed

- **`doc::HeaderFooterStories` fields are now `Option<Arc<str>>`** (they were `Option<String>`), so the stories a section inherits are shared with the section they come from instead of copied. Callers that read a field as a string keep working through `Deref<Target = str>` (`as_deref()`, `&**s`); code that moved a `String` out of a field, or built the struct with `String`s, needs `Arc::from`.

- **`Document::to_html`, `to_html_with` and `to_markdown_with` of an `.xlsx` convert and render one sheet at a time.** Each sheet's IR is dropped as soon as it is rendered, so the peak is about one sheet's IR instead of the whole workbook's. The allowances for the IR's padding cells and pictures are therefore a sheet's (`limits::max_sheet_ir_padding_bytes`, `limits::max_image_bytes`, given back with the sheet), and a workbook of many sheets each within them is rendered completely where the 512 MiB document-wide padding allowance cut it (8 sheets of 10,000 rows with cells in `A` and `CV` was cut in the first sheet in HTML and Markdown-with-options, and complete in `plain_text` and `to_markdown`). What bounds the run as a whole is unchanged and still document-wide: the text budget (`limits::set_max_text_chars`), with the characters the padding adds, and, new, a total on picture bytes. For a workbook within the budgets the output is byte-for-byte what it was. `Document::to_ir` holds every sheet at once and keeps the document-wide `limits::max_ir_padding_bytes` and `limits::max_image_bytes`; its output and its IR JSON are unchanged.

- **`XlsxDocument::plain_text` and `to_markdown` no longer build each sheet's text, the list of its parts and the joined result in turn.** They write the text straight into one result, the cells of a row charged to the text budget before they are written, so the intermediate copies (about three times the output) are gone. The output, including where a budget cuts it, is byte-for-byte what it was.

- **The CLI's `text`, `markdown` and `html` commands and the MCP `extract` tool write a spreadsheet's rendering as it is produced.** The CLI writes to stdout through a buffer instead of building the string first; `extract` writes into its size-limited buffer, so a rendering over the 32 MiB limit stops at the limit instead of being built whole and then refused. Output and errors are unchanged. `info` and `ir` still convert the whole workbook.

## [0.1.15] - 2026-10-15

### Added

- **`.doc` heading levels come from the style sheet ([#132](https://github.com/yfedoseev/office_oxide/issues/132); contributed by @xugangqiang in [#127](https://github.com/yfedoseev/office_oxide/pull/127)).** A paragraph whose style is a built-in `Heading 1`–`Heading 9`, or a user-defined `Heading N`, now renders at that level; previously the level came from a line-shape guess that could only produce 1 or 2 and missed any heading that was not ALL-CAPS. A direct `sprmPOutLvl` still beats the style, in both directions (operand 9 is explicit body text), and the line-shape guess stays as the fallback for paragraphs with no level, switched off only by a paragraph that states its own.
- **A style that is not named `Heading N` can still be a heading.** A style's own outline level and the style it is based on (`istdBase`) are now followed, to a bounded depth, so styles such as "H2", "Level 1" or a "Section Head" based on `Heading 2` resolve to their real level. A cyclic or over-deep base chain resolves to no level and is reported once.

  Measured against v0.1.14 on the 634 unique `.doc` files under `~/projects` that an independent MS-DOC reader (FIB, piece table, PAPX, STSH with inheritance; sharing no code with the crate) could score: 990 of about 3,250 true headings found before, **3,216** now (misses 2,262 → 36), and 77% → **97.5%** of the matched headings at the correct level. No scored document gets worse. The one heading lost, in `tdf81705_outlineLevel.doc`, is a paragraph explicitly marked body text. All other formats, and the `text` output of every `.doc`, are byte-identical across all 10,170 unique files. A style sheet that defines only part of the hierarchy still falls back to the guess for the rest, and a heading-styled paragraph that is a whole sentence (a run-in head) becomes one long heading, as it is in the document.

### Fixed

- **A rejected or truncated `.doc` style sheet is reported instead of dropped.** An absent or out-of-range style sheet, a bad header or `cbSTDBaseInFile`, or a sheet that declares more styles than it carries now adds a warning to `Metadata::warnings`; the styles read before a cut are kept. Four of the 804 `.doc` files in the local corpus carry one; no other output changes.

- **A DOCX table with no known widths no longer collapses in Microsoft Word ([#433](https://github.com/yfedoseev/office_oxide/issues/433)).** The writer emitted `<w:gridCol w:w="0"/>` for every column of a table that carried no widths — every table from `DocumentIR::from_markdown` — on the belief that zero means "auto". `w:gridCol` has no auto value and Word sizes columns from it, so the columns collapsed, cell text wrapped after a few characters and a short table ran over several pages; LibreOffice lays the table out from its content, which is why it looked fine there. An unsized table is now written the way pandoc writes one: `w:tblW` of `5000` `pct` (the full text width) and equal non-zero `w:gridCol` values as hints. A table with an explicit `width_twips` but no column widths splits that width equally; explicit widths are written unchanged. The grid also counts `gridSpan`, so a row of merged cells no longer under-sizes it, and the column count stays bounded. The same applies to the `DocxWriter::add_table` builder. A `pct` width reads back as no width, so the written table re-reads exactly as before.
- **Tables from Markdown are readable in Word ([#405](https://github.com/yfedoseev/office_oxide/issues/405)).** `DocumentIR::from_markdown` (and so `create_from_markdown*` to DOCX; the PPTX and XLSX writers do not draw these properties and are unchanged) now gives a pipe table thin single borders on every edge and between every cell, a thicker rule under the header row, and marks the header row to repeat at the top of each page a table runs onto. Header text is deliberately not bolded: it would render as `**Col A**` and break the Markdown round trip, which stays byte-identical. Tables read from documents, and tables built through the IR directly, are written exactly as given.

  Both fixes were checked against the previous build on all 10,014 unique files under `~/projects`, each taken through file → IR → DOCX → re-read and through `to_markdown()` → `from_markdown()` → DOCX → re-read. Of 9,703 files that open in both builds, the re-read `plain_text`, `to_markdown()` and table count are identical on every one, and no file changes status. The writer now emits no zero `gridCol` anywhere (3,782 files, and every one of the 4,702 markdown-derived documents with a table, had them before); the only other difference is that the re-read IR of those 3,782 files now carries real column widths. The written packages pass the OOXML schema and OPC validators with 0 findings. A grid whose widths are all zero — what the reader produces from a source with fractional `w:gridCol` values — is treated as unsized instead of being written back.

### Changed

- The DOCX text extractor, DOCX writer and DOCX-to-IR converter now clamp heading depth through the IR's single `MAX_HEADING_DEPTH` constant instead of a literal 6 (no behaviour change).
- **Every dependency is current.** `zip` 9.0.3 (a major bump), `pyo3` 0.29.3, `serde_json` 1.0.154, `libc` 0.2.190 and the other compatible Rust updates in the root, bench and fuzz lockfiles; `koffi` 3.3.3, `ruff` 0.17 (the lint pin moves to `>=0.17,<0.18`), the C# test packages (`Microsoft.NET.Test.Sdk` 18.11.0, `xunit.runner.visualstudio` 4.0.1), and every pinned GitHub Action. The benchmark competitor pins stay put, as they must match the published tables.
  - `zip` 9 makes an entry name fallible and refuses an entry whose data exceeds its declared uncompressed size. An archive with a name that cannot be decoded is now rejected when it is opened, never skipped, so the duplicate-entry check cannot be dodged; and a forged declared size now fails inside the zip reader (an `Io` error) instead of at the package byte budget, which stays as the second line of defence.
  - Checked against the previous build on all 10,170 unique files under `~/projects` over `text`, `markdown`, `html` and `ir`: no file changes status, and every output is byte-identical except two deliberately corrupt archives, where the same sheets are unreadable as before and only the wording of the note changes ("Invalid checksum" becomes "File is larger than its declared uncompressed size"). The Python, Node, Go, C# and WebAssembly binding tests pass against the rebuilt library, and the crate still builds on its minimum supported Rust, 1.88.

## [0.1.14] - 2026-10-14

### Fixed

- **`DocumentIR` is source-compatible with 0.1.11 again.** 0.1.12 added a public `defined_names` field to `DocumentIR`, so any crate that builds it with a literal naming every field — including the published `pdf_oxide` 0.3.78 — failed to compile (`E0063`) as soon as Cargo resolved office_oxide 0.1.12 or 0.1.13. `DocumentIR` is back to exactly `{ metadata, sections }`, and the defined names (named ranges, print areas; XLSX and XLS) now live on `Metadata::defined_names`. The workaround of pinning `office_oxide = "=0.1.11"` is no longer needed. A guard test builds `DocumentIR` from those two fields alone.

### Changed

- **Migration for 0.1.12 / 0.1.13 users of `defined_names`:** read `ir.metadata.defined_names` instead of `ir.defined_names`. In the IR JSON the `defined_names` array moved from the top level into `metadata`. Nothing else in the IR changed: a 1,500-file comparison against 0.1.13 (250 each of `.docx/.xlsx/.pptx/.doc/.xls/.ppt`, four output surfaces) was identical once that field is mapped across.

## [0.1.13] - 2026-09-30

> Every finding of the 2026-09-27 audit — 19 issues, 165 checklist items — plus ten resource-bound defects reported privately, fixed test-first: each with a minimal synthetic document built in code, a test that fails before the fix and passes after. A two-arm sweep of 10,215 unique real Office documents against v0.1.12 then found nine regressions introduced by those fixes; all nine are fixed below, each with its own regression test, and the re-run sweep shows no crash, no timeout, no unintended new failure and no real slowdown.

### Verified

Every Office document under `~/projects` — 16,247 files across 13 projects, **10,215 unique** by content hash (`.docx` 3,446, `.xlsx` 3,039, `.xls` 1,177, `.pptx` 1,139, `.doc` 793, `.ppt` 413 and 208 macro/template/binary variants) — swept arm-to-arm with v0.1.12 on four surfaces (`text`, `markdown`, `html`, `ir`), 40,860 observations per arm, compared with `scripts/regression-sweep/compare_tree.py` (status transitions, word multisets with split/join awareness, timing).

| | v0.1.12 | 0.1.13 |
|---|---|---|
| crashes / timeouts | 0 / 0 | **0 / 0** |
| files newly failing | — | 11, all intended: 10 password-protected `.ppt` decks that v0.1.12 returned as empty text now fail with `Encrypted`, and one package with case-only duplicate part names (5 bytes under v0.1.12) is refused as POI refuses it |
| slowdowns (≥ 1.5× and ≥ 200 ms, re-timed serially) | — | **0** for the same output; 3 pairs are slower only because they now produce correct content v0.1.12 dropped (compressed WMF pictures inflated into valid metafiles, an external-source pivot's cached data, real per-cell fills and borders) |

The 0.1.13 sweeps found eleven regressions introduced by this release's own fixes (listed under Fixed); each was fixed with a regression test and the full sweep re-run after every fix, the last on the release head. Every remaining word-level difference was read and resolved to a deliberate class: first-page headers/footers of sections without `titlePg` (Word does not show them), PowerPoint master placeholder prompts, v0.1.12's synthetic `SheetN` names replaced by real BIFF5 names, the in-text nested-table notice moved to `Metadata::warnings`, relationship ids no longer written as markdown image targets, linked-picture URLs moving from text into `<img src>`, and superscripts rendered as `<sup>`. 1,476 file/surface pairs only gained words.

### Security

Resource bounds for untrusted input — each was reachable from a small file through every binding and the MCP server:

- **CFB (`.doc`/`.xls`/`.ppt`):** the header's mini-stream cutoff no longer sizes an allocation (the spec's fixed 4096 is used, as POI does), and mini-stream reads reserve at most the mini stream; FAT construction is bounded by the sectors the file holds, a FAT sector listed twice is an error, and DIFAT slots past the declared FAT count are ignored.
- **XLS / XLSX `to_ir()`:** merged-range and hyperlink-range expansion is clipped to the cells that exist and bounded by a per-sheet work budget — a 2 KB file declaring a full-grid merge no longer hangs; ranges past the budget are reported with a notice.
- **DOCX:** text boxes nested in text boxes are capped (16 levels, measured against a 2 MiB stack) with a visible truncation notice; one wide table row no longer makes every row cost the full width in `to_ir()` and the renderers.
- **IR renderers and IR writers:** recursion over a deserialized `DocumentIR` is depth-bounded on every surface (plain text, markdown, HTML, DOCX/XLSX/PPTX writers); a skipped subtree leaves a visible notice.
- **PPT:** table-grid clustering arithmetic is widened, so extreme shape anchors no longer overflow.
- **OPC packages:** the zip central directory is walked independently of the zip crate, so a ZIP64 archive or an understated EOCD count can no longer hide a duplicate entry; a package's entry count (100,000) and total decompressed bytes (2 GiB) are bounded (`limits::max_package_entries`, `limits::max_package_bytes`) with a new `core::Error::PackageLimit`.
- **FFI:** every `extern "C"` export contains panics, so a bug can no longer unwind into — and abort — the host process (Node, Go, .NET, CPython).

### Fixed

- **XLS** ([#422](https://github.com/yfedoseev/office_oxide/issues/422), [#411](https://github.com/yfedoseev/office_oxide/issues/411)): a defined name, conditional format or comment on row 65,536 no longer overflows (panic in debug, row 0 in release); BIFF5 sheet names, `FORMAT` and `STRING` records are decoded with the BIFF5 layouts; chart, dialog and VBA-module substreams are no longer read as worksheets full of plotted values; Excel 4.0 macro sheets keep their cells; encrypted BIFF2–4 files are refused; a truncated SST is flagged only when a cell loses its string, and a malformed SST header no longer fails the workbook; `MULRK`/`MULBLANK` column arithmetic stays in range; sheets are paired with their substreams by stream offset; macros are detected through the `_VBA_PROJECT_CUR` storage Excel actually writes.
- **XLSX / XLSB** ([#407](https://github.com/yfedoseev/office_oxide/issues/407)): sparse rows and columns are placed by cell reference in `plain_text`, `to_csv` and `to_markdown`; `set_cell` writes into `<sheetData/>` and resolves sheets through the workbook relationships; a sheet with a missing part is reported and hidden flags stay aligned; chartsheets and dialog sheets surface as named sections with their chart text and no cells, macro sheets with their cells; sheet headers/footers, table parts, pivot tables (and pivot cache records when the source is not in the workbook), `t="d"` dates, cell alignment, wrap text, fills and borders are read; defined names, sheet visibility and `xml:space="preserve"` are written.
- **DOCX** ([#406](https://github.com/yfedoseev/office_oxide/issues/406)): table rows and cells wrapped in `w:ins`/`w:moveTo`/`w:customXml` are kept; ruby guides stay out of the base text; fonts and sizes follow the script of the text (`eastAsia`/`hAnsi`/`cs`, `szCs`/`bCs`); full `lvlOverride` levels, level indents, marker patterns and alignment are read; comments keep their anchor and round-trip as comments; table styles apply their borders, shading and conditional run/paragraph formatting consistently in `to_ir()` and the direct renderers; hyperlink runs inside `w:ins`/`w:sdt`/`w:smartTag`, `<w:tab></w:tab>`, linked pictures, section page numbering and note settings are read; `replace_text` matches text split across runs and reaches headers, footers, notes and comments; nested lists are written under their parent item.
- **DOC** ([#410](https://github.com/yfedoseev/office_oxide/issues/410)): reference-mark and hyphen control characters no longer leak into text; every section has its own headers/footers and IR section; hidden, strike, super/subscript, highlight, colour and line-spacing SPRMs are decoded; negated `ilfo` values, piece `Prm` modifiers, `PICF` pictures (including all Word 6.0/95 pictures) are resolved; each text box is its own element, placed after the paragraph holding its anchor; a flattened nested table is a metadata warning instead of fabricated document text.
- **PPT** ([#412](https://github.com/yfedoseev/office_oxide/issues/412)): speaker notes are extracted from real files; encrypted decks fail with `PptError::Encrypted` instead of returning empty text; header/footer text follows its container, flags and per-slide overrides; typefaces, internal-jump hyperlinks, media placeholders and master text styles at every level are resolved, with masters found through the master list rather than a masked persist id; hyperlink and style ranges are measured in UTF-16 units; master boilerplate is excluded by structure, not English prompt strings; static master text and the shown notes/handout header and footer appear once per deck.
- **PPTX** ([#408](https://github.com/yfedoseev/office_oxide/issues/408)): real SmartArt text is read from the diagram data part; `replace_text` reaches every slide, notes, layout and master part; the full text-style inheritance chain is resolved; comment authors, linked pictures, media and OLE previews, paragraph spacing/indents/tabs and percent spacing are read; one policy for unreadable slide, notes and comment parts; static layout/master text appears once per deck, placeholder prompts never.
- **CFB** ([#409](https://github.com/yfedoseev/office_oxide/issues/409)): compressed EMF/WMF/PICT pictures are inflated (one shared BLIP decoder); the mini stream is bounded by the root entry's size; short stream reads are reported; directory names compare Unicode case-insensitively; SummaryInformation is found by FMTID.
- **OPC core** ([#414](https://github.com/yfedoseev/office_oxide/issues/414)): case-only duplicate part names are rejected; a repeated relationship Id resolves to its first occurrence; `_rels` filtering, the main-part fallback, root-close check and Strict namespace alternates are exact.
- **Metadata** ([#419](https://github.com/yfedoseev/office_oxide/issues/419)): `app.xml` keeps entity references; `lastModifiedBy`, revision, category, content status, language, company, manager, custom properties, signature presence and the thumbnail reach the IR, the CLI and MCP `info`, and are written back; every writer emits `docProps/app.xml`; legacy DocumentSummaryInformation is read.
- **Write path** ([#415](https://github.com/yfedoseev/office_oxide/issues/415)): speaker-note hyperlinks resolve in the notes part; ordered lists are written with `buAutoNum`; no XML-illegal characters in comments or relationship targets; the DOCX page background is displayed; `notesMasterIdLst` is declared; an empty text box holds a paragraph; markdown keeps `# Learn C#` and an ordered list's start number; `w:outlineLvl` is clamped.
- **Bindings** ([#416](https://github.com/yfedoseev/office_oxide/issues/416)): Python and JS writers store booleans as booleans; Go, C# and JS surface writer status codes as errors; FFI error codes follow the error variant; formulas can be written from every binding; `from_bytes` editing and markdown with images are available in Go, JS, C# and Python; integers beyond 2^53 are refused rather than rounded.
- **CLI / MCP** ([#417](https://github.com/yfedoseev/office_oxide/issues/417)): an empty search string is an error in the library (it used to overwrite the original through MCP); in-place saves are atomic; MCP no longer replies to notifications; stdout errors are reported instead of panicking; `info` reports every set property and the reader's warnings.
- **Parsed-but-inert** ([#418](https://github.com/yfedoseev/office_oxide/issues/418)): hyperlink tooltips, `title_page`, row height rules, the gutter, the decorative-picture flag, table styles, numbering marker text, shading patterns, XLSX fills/borders and theme fonts now reach the IR (and the writers).
- **Text with entity references** keeps its spaces (`see A &amp; B` was read as `see A&B`) in every reader that trims text.
- Regressions the release sweep found in this release's own fixes, each with a regression test: a CDATA sheet header failed the whole sheet; a CRC-damaged part failed its whole document (it is now read and reported in `Metadata::warnings`) and an XLSX sheet whose part failed to read could fall back to another sheet's cells; the HTML renderer cut the deepest levels of a table the reader had kept; `sprmCFVanish` operand `0x81` hid ordinary visible text; macro sheets lost their cells and chart/dialog sheets their tab names; static slide-master text and shown notes/handout headers and footers vanished from `.ppt`; a first-page header a later `titlePg` section inherits was hidden; pivot caches were parsed once per pivot even when a workbook table made them unnecessary (4.7× slower open on a 17-pivot workbook).

### Performance

- DOC piece-table lookups binary-search in CP and FC ([#421](https://github.com/yfedoseev/office_oxide/issues/421)); PPTX shared layouts, images and charts are read once per deck; DOCX altChunk HTML is flattened in linear time, each paragraph is converted once and each style chain folded once per conversion; XLSX date styles are found once and custom number formats compiled once per render.

### Changed

- **`Metadata::warnings`** is new: what a reader worked around (a skipped or CRC-damaged part, a flattened structure, a compound file whose counts disagree), one line each; CLI and MCP `info` list them.
- A part whose bytes all arrived but fail their CRC-32 is read as stored and reported, instead of failing the document.
- Static slide-master/layout text is extracted once per presentation in a trailing "Slide Master" section; chart and dialog sheets appear as named sections; `Section::sheet_kind` marks worksheets, macro sheets, chart sheets and dialog sheets.
- Direct and IR markdown lay documents out alike (DOCX block spacing, PPTX lists, XLSX tables).
- Public structs gained fields and enums gained variants (`Metadata`, `Section`, `List`, `TableCell`, `TableRow`, `Image`, `PageSetup`, `xls::Sheet`, `xlsx::Worksheet`, PPTX shape types, `GraphicContent`, `BulletStyle`); code that builds them with struct literals or matches them exhaustively needs `..Default::default()` / a wildcard arm.
- Python 3.10+ (abi3-py310), matching the versions CI tests.
- **Dependencies current**: `cargo update` across the workspace, bench and fuzz lockfiles (`encoding_rs` 0.8.42, `thiserror` 2.0.21, `wasm-bindgen` 0.2.129; no direct dependency has a newer major), `koffi` 3.3.2, `markitdown` 0.1.8 bench pin, `uv.lock` (`ruff` 0.16.9), CodeQL action 4.38.2, `setup-uv` 10.2.0, `install-action` 2.87.22, and a release-workflow action pinned to its tagged release instead of `main`. Swept separately (the release branch before vs after the update, same 10,215 files): all 40,860 outputs byte-identical, no status change, no slowdown.
- CI now runs the workspace tests (CLI and MCP included), builds and briefly fuzzes the fuzz target, tests the `mmap`/`parallel` features, generates the wasm-bindgen glue, builds the whole workspace at MSRV, and pins the OOXML schema downloads to commits and SHA-256s; the fuzz target renders every document it parses.

## [0.1.12] - 2026-09-21

> Read-path fidelity across all six formats, and the first real fidelity work on the legacy binaries. 141 issues closed, every dependency at its latest release, and a corpus verification that compares against independent readers, not just the previous version.
>
> The legacy `.doc`/`.xls`/`.ppt` readers went from "extracts the text" to structured: per-note footnotes, endnotes, comments and headers/footers in `.doc`; character and paragraph formatting from CHPX/PAPX; list identity via `PlfLst`/`PlfLfo`; `.ppt` direct formatting, master inheritance, hyperlinks, hidden slides, OLE objects and grid-of-shapes tables; `.xls` comments, hyperlinks, data validation, chart text and the 1904 date system. On the OOXML side the recurring shape was *parsed-then-discarded*: the parser had it, a converter or one of the four renderers dropped it — rich-text runs, merged cells, defined names, formula text, speaker-note formatting, OMML equations, SmartArt, embedded charts, shape click actions.
>
> Verified against **10,599 real Office documents** — arm-to-arm with v0.1.11 (5 crashes and 5 timeouts in the baseline, **0** here) and against ten independent readers, which is how the largest content-loss defect in the crate's history (80 `.xls` files at 2–58 % of every reference) was found and fixed *before* this shipped.

### Verified

Two passes over `~/projects/office_oxide_tests/` — **10,599 real Office documents**, all six formats plus their macro/template/strict/encrypted variants, 38 sources.

**Pass 1 — release regression sweep, after every dependency was bumped.** v0.1.11 and this branch, each over every file on four surfaces (`text`, `markdown`, `html`, `ir`) — 42,396 observations per arm, compared on status, on word multisets with split/join awareness (a `\w+` tokenizer misreads restored entities as loss), and on timing, using `scripts/regression-sweep/sweep_tree.py` + `compare_tree.py`.

| | v0.1.11 | 0.1.12 |
|---|---|---|
| `ok` observations | 40,202 | **40,716** |
| crashes (SIGKILL/abort) | 5 | **0** |
| timeouts (120 s) | 5 | **0** |
| `err` | 2,184 | 1,680 |

Every `ok → err` (20 files) replaced a 0-byte "success": 7 encrypted `.pptx` now say so ([#296](https://github.com/yfedoseev/office_oxide/issues/296), [#320](https://github.com/yfedoseev/office_oxide/issues/320)), 12 `.ppt` files with no readable PowerPoint stream (PowerPoint 4.0/95-only files and fuzzer artefacts) now say so ([#351](https://github.com/yfedoseev/office_oxide/issues/351)), and POI's own duplicate-entry `.xlsx` is refused rather than silently resolved to one copy. Every `err → ok` (146 files) is a macro-enabled or template extension that used to be rejected ([#282](https://github.com/yfedoseev/office_oxide/issues/282)). The sweep's first run was *not* clean — it found the `ir` memory regression ([#349](https://github.com/yfedoseev/office_oxide/issues/349)), the TOC-entries regression ([#350](https://github.com/yfedoseev/office_oxide/issues/350)) and two smaller ones ([#355](https://github.com/yfedoseev/office_oxide/issues/355), [#356](https://github.com/yfedoseev/office_oxide/issues/356)) that this release had introduced, all fixed before this section was written.

**Pass 2 — quorum check against independent readers.** The same corpus through Apache Tika 3.3.2 (POI), python-docx 1.2.0, pandoc, openpyxl 3.1.5, python-calamine, xlrd 2.0.2, catdoc, antiword, xls2csv and catppt (`scripts/regression-sweep/refpanel.py` + `quorum.py`). A file is flagged when at least two references agree with each other while office_oxide recovers under 70 % of their words. This is the pass a two-arm diff cannot do, and it found the defects present in v0.1.11 too: **80 `.xls` files at 2–58 % of every reference** — one shared-string parsing bug ([#353](https://github.com/yfedoseev/office_oxide/issues/353)) — plus [#352](https://github.com/yfedoseev/office_oxide/issues/352), [#357](https://github.com/yfedoseev/office_oxide/issues/357), [#358](https://github.com/yfedoseev/office_oxide/issues/358). Flagged files went from 149 to 64; the `.xls` odd-one-out count from 80 to **0**. What remains flagged is explained and deliberate: Word 6.0/95 `.doc` rejected with a clear error (23), hidden text (`w:vanish`, including style-inherited) and tracked deletions the reference tools show and Word does not, `.xlsx` sheet names ([#359](https://github.com/yfedoseev/office_oxide/issues/359)), and fuzzer artefacts.

**Pass 3 — performance and rendering, after the fixes above.** The same corpus through `perf_sweep.py` on both arms (zero status regressions; every timing flag re-measured serially), the quorum re-run against Tika, the python libraries and the C tools (74 flagged files, every one explained: hidden text and tracked deletions the references show, formulas vs cached values, references that returned nothing, Word 6/95 and `.xlsb`), and a structural comparison of our markdown against pandoc's GFM over the 3,521 `.docx` files, which found the three renderer defects listed under *Fixed*.

**Pass 4 — one harness, one verdict per file** ([#401](https://github.com/yfedoseev/office_oxide/issues/401)). The passes above became `scripts/regression-sweep/harness.py`: both arms, the reference panel, a write→reread round trip and the markdown-structure comparison over the corpus, reduced to a per-file verdict with a stated decision model — the consensus of references that agree with each other is the expected output; a two-arm loss counts only when the consensus has the words and the new arm does not; an arm that sides with a different reference than the other arm is a policy split, not a regression; `w:vanish` runs and field vocabulary explain a loss on their own. It added two axes a two-arm diff and a reference quorum are both blind to: our own text, markdown and html of one file compared with each other (the *dual-renderer* class), and the words of the original IR against those of the IR read back from what we wrote. Those found the defects under *Fixed* from [#381](https://github.com/yfedoseev/office_oxide/issues/381) on: every `docx/headers_footers/*` file lost its headers on a round trip, every picture below body level was dropped on write, formulas became constants, `.doc` fast-saved files had their paragraphs re-cut at piece boundaries and their tracked deletions shown on the structured surfaces only, and no markdown renderer escaped document text.

**Test-coverage audit.** Every one of the milestone's 162 issues was mapped to the tests its fix commits added (plus a topic search for the ones committed without a number) and judged against the cases its report enumerates. Gaps found and closed: endnote and comment bodies in `plain_text()`/`to_markdown()` (only footnotes were asserted), the DOCX side of numbered headings, paragraph-style-inherited `w:vanish`, `FORMTEXT` default values, CR-only runs' `xml:space`, `docProps/app.xml` and `has_macros` on PPTX and XLSX and a legacy `_VBA_PROJECT` storage, XLSX `plain_text()` sheet names, and a README-vs-clap consistency test for the CLI section (which found the undocumented `replace --output`). The `has_macros` gap turned out to be a real defect ([#380](https://github.com/yfedoseev/office_oxide/issues/380)).

Plus: 1,380 library and workspace tests green (`--all-features`), clippy `-D warnings` clean, rustdoc `-D warnings` clean, `cargo fmt --check` clean, `cargo audit` and `cargo deny` clean, C# (xUnit runner 4.0) and Node (koffi 3.3.1) suites green, and the `house-rules` gate green.

### Added

- **Word 6.0/95 text** ([#374](https://github.com/yfedoseev/office_oxide/issues/374)). The family was refused outright while catdoc, antiword, LibreOffice and POI read it. Its own FIB layout, no table stream, one-byte characters in the document's code page (or two-byte units under `fExtChar`, the Macintosh and Far East builds), non-complex and fast-saved files, subdocuments included; 15 of the 17 openable corpus files at 0.90–1.00 recall against catdoc, antiword and Tika. Text only — no paragraph or table structure.
- **Excel Binary Workbooks (`.xlsb`)** ([#375](https://github.com/yfedoseev/office_oxide/issues/375)). The BIFF12 parts — sheet list, shared strings, number formats and cell formats, every cell and formula-value record, merged ranges — decoded into the same model as `.xlsx`, so every converter and renderer applies; 29 of the 35 openable corpus files at 1.00 recall against python-calamine, the rest differing only in value representation. Cached values only (no formula text); fonts, fills and borders not read.
- **Excel 2.x–4.0 (BIFF2/3/4) worksheets** ([#376](https://github.com/yfedoseev/office_oxide/issues/376)). The pre-OLE2 record stream — refused by name before — is read as a one-sheet workbook: the version's own cell records, `CODEPAGE` text, formula results via `STRING`; 1.00 recall against xlrd on the corpus files. Number formats are not applied.
- **Legacy `.doc` structure.** Footnotes and endnotes as one IR element each ([#286](https://github.com/yfedoseev/office_oxide/issues/286)); comments split per comment via `PlcfandTxt`/`PlcfandRef` with authors from `GrpXstAtnOwners` ([#345](https://github.com/yfedoseev/office_oxide/issues/345), [#298](https://github.com/yfedoseev/office_oxide/issues/298)); header/footer stories split via `PlcfHdd` ([#285](https://github.com/yfedoseev/office_oxide/issues/285)); CHPX character formatting and PAP alignment/indent/spacing ([#287](https://github.com/yfedoseev/office_oxide/issues/287)); list identity through `PlfLst`/`PlfLfo` for `start_number` and ordered/unordered ([#250](https://github.com/yfedoseev/office_oxide/issues/250)); embedded OLE objects under `ObjectPool` ([#284](https://github.com/yfedoseev/office_oxide/issues/284)); subdocument bodies in `plain_text()`/`to_markdown()` ([#248](https://github.com/yfedoseev/office_oxide/issues/248)); `HYPERLINK` field targets with field instructions stripped ([#249](https://github.com/yfedoseev/office_oxide/issues/249)).
- **Legacy `.ppt` structure.** Direct character/paragraph formatting from `StyleTextPropAtom` ([#254](https://github.com/yfedoseev/office_oxide/issues/254)); placeholder formatting inherited from the master's `TextMasterStyleAtom` ([#335](https://github.com/yfedoseev/office_oxide/issues/335)); hyperlinks from `InteractiveInfo`/`ExHyperlink` ([#257](https://github.com/yfedoseev/office_oxide/issues/257)); grid-of-shapes tables reconstructed as `Element::Table` ([#255](https://github.com/yfedoseev/office_oxide/issues/255)); hidden slides ([#297](https://github.com/yfedoseev/office_oxide/issues/297)); header/footer/date text ([#308](https://github.com/yfedoseev/office_oxide/issues/308)); embedded and linked OLE objects ([#337](https://github.com/yfedoseev/office_oxide/issues/337)).
- **Legacy `.xls` structure.** Cell comments from `NOTE`/`TXO`/`OBJ` ([#307](https://github.com/yfedoseev/office_oxide/issues/307)); hyperlinks from `HLINK` ([#306](https://github.com/yfedoseev/office_oxide/issues/306)); embedded chart text via `SeriesText` ([#246](https://github.com/yfedoseev/office_oxide/issues/246)); pre-OLE2 BIFF2/3/4 files named instead of failing at the CFB layer ([#227](https://github.com/yfedoseev/office_oxide/issues/227)) — and, later in this release, read ([#376](https://github.com/yfedoseev/office_oxide/issues/376)).
- **Across formats.** `Metadata::has_macros` for all six formats ([#283](https://github.com/yfedoseev/office_oxide/issues/283)); `SummaryInformation` title/author/dates for `.doc`/`.xls`/`.ppt` ([#244](https://github.com/yfedoseev/office_oxide/issues/244)); `docProps/app.xml` read on open ([#245](https://github.com/yfedoseev/office_oxide/issues/245)); conditional formatting as a new IR type for `.xlsx`/`.xls` ([#252](https://github.com/yfedoseev/office_oxide/issues/252)); data validation rules ([#275](https://github.com/yfedoseev/office_oxide/issues/275), [#343](https://github.com/yfedoseev/office_oxide/issues/343)); defined names reach `to_ir()` ([#251](https://github.com/yfedoseev/office_oxide/issues/251)); placeholder role (subtitle/date/footer/…) on PPTX and PPT shapes ([#258](https://github.com/yfedoseev/office_oxide/issues/258)); PPTX placeholder formatting inherited from `<p:txStyles>` ([#291](https://github.com/yfedoseev/office_oxide/issues/291)).
- **Writers.** Hyperlinks in XLSX and PPTX output — a total, unconditional loss before ([#262](https://github.com/yfedoseev/office_oxide/issues/262)); speaker-note formatting and list structure carried through read, write and every renderer ([#290](https://github.com/yfedoseev/office_oxide/issues/290)); table-cell character formatting written and read back from inline rich strings ([#346](https://github.com/yfedoseev/office_oxide/issues/346)); cell comments written as real comments ([#344](https://github.com/yfedoseev/office_oxide/issues/344)).
- **Renderers.** `to_html_with()` with base64 image embedding ([#318](https://github.com/yfedoseev/office_oxide/issues/318)); underline, super/subscript, highlight, colour, font and caps in HTML ([#314](https://github.com/yfedoseev/office_oxide/issues/314)); `List.start_number` honoured in HTML and markdown ([#315](https://github.com/yfedoseev/office_oxide/issues/315)); adjacent same-style lists merged ([#317](https://github.com/yfedoseev/office_oxide/issues/317)).
- `examples/corpus_validate`, the per-step JSONL harness behind the verification above.
- `scripts/house-rules-check.sh` and its CI job: every test function is named `test_*`, no `allow(dead_code)`, no issue numbers in code or comments.

### Fixed

- **DOCX read.** Footnote/endnote/comment bodies in `plain_text()`/`to_markdown()` ([#240](https://github.com/yfedoseev/office_oxide/issues/240)); footnote/endnote/comment *reference markers* reach `to_ir()`, so a note is anchored to where it was cited ([#241](https://github.com/yfedoseev/office_oxide/issues/241)); `HYPERLINK` field codes via `w:fldChar`/`w:instrText` and `w:fldSimple` ([#267](https://github.com/yfedoseev/office_oxide/issues/267)); a hyperlink carrying both `r:id` and `w:anchor` keeps its external target ([#242](https://github.com/yfedoseev/office_oxide/issues/242)); OMML equation text ([#270](https://github.com/yfedoseev/office_oxide/issues/270)); WordArt `v:textpath` text ([#274](https://github.com/yfedoseev/office_oxide/issues/274)); legacy VML `v:imagedata` images ([#268](https://github.com/yfedoseev/office_oxide/issues/268)); form-field checkbox/dropdown/default state from `w:ffData` ([#276](https://github.com/yfedoseev/office_oxide/issues/276)); embedded chart title/category/series/data text ([#273](https://github.com/yfedoseev/office_oxide/issues/273)); SmartArt diagram text ([#271](https://github.com/yfedoseev/office_oxide/issues/271)); embedded native OOXML packages ([#304](https://github.com/yfedoseev/office_oxide/issues/304)); OMML equation text and `mc:AlternateContent` ([#272](https://github.com/yfedoseev/office_oxide/issues/272)); `w:vanish` hidden runs excluded ([#305](https://github.com/yfedoseev/office_oxide/issues/305)); sdt-wrapped cells/rows kept, tracked-change `cellDel` excluded ([#277](https://github.com/yfedoseev/office_oxide/issues/277), [#266](https://github.com/yfedoseev/office_oxide/issues/266)); text box with `prstGeom` no longer double-extracted ([#263](https://github.com/yfedoseev/office_oxide/issues/263)); interrupted numbered lists keep counting ([#243](https://github.com/yfedoseev/office_oxide/issues/243)); `<w:startOverride>` read back ([#260](https://github.com/yfedoseev/office_oxide/issues/260)); numbered headings emitted as `Heading`, not `ListItem` ([#223](https://github.com/yfedoseev/office_oxide/issues/223)); outline-level headings keep their other paragraph properties, cell `<w:jc>` maps to `text_align`, thematic breaks become `Element::ThematicBreak` ([#215](https://github.com/yfedoseev/office_oxide/issues/215)); nested-table recursion bounded ([#329](https://github.com/yfedoseev/office_oxide/issues/329)); the default `to_markdown()` renderer increments ordered-list numbers instead of repeating the start value ([#316](https://github.com/yfedoseev/office_oxide/issues/316)); `truncated_subtrees()` reports content dropped past the depth bound ([#218](https://github.com/yfedoseev/office_oxide/issues/218)); malformed `dcterms:created`/`modified` normalised or dropped ([#217](https://github.com/yfedoseev/office_oxide/issues/217)).
- **DOCX write.** Hyperlink relationships scoped to the part that uses them ([#293](https://github.com/yfedoseev/office_oxide/issues/293)); `r:id` + `w:anchor` both kept ([#292](https://github.com/yfedoseev/office_oxide/issues/292)); custom footnote marks round-trip ([#219](https://github.com/yfedoseev/office_oxide/issues/219)); table caption no longer duplicates each round trip ([#311](https://github.com/yfedoseev/office_oxide/issues/311)); grid width from `col_span` sums ([#265](https://github.com/yfedoseev/office_oxide/issues/265)); nested lists share one `numId` across levels but each nested-in-container list gets its own ([#261](https://github.com/yfedoseev/office_oxide/issues/261), [#339](https://github.com/yfedoseev/office_oxide/issues/339)); title heading not duplicated ([#338](https://github.com/yfedoseev/office_oxide/issues/338), [#259](https://github.com/yfedoseev/office_oxide/issues/259)); tab/CR/LF-only runs get `xml:space="preserve"` ([#294](https://github.com/yfedoseev/office_oxide/issues/294)); image and font parts get a `Default` content type ([#295](https://github.com/yfedoseev/office_oxide/issues/295)); pretty-print indentation removed — it was Θ(depth²) memory on the write path ([#220](https://github.com/yfedoseev/office_oxide/issues/220)); rarely-populated sub-structs boxed to cut per-element memory ([#328](https://github.com/yfedoseev/office_oxide/issues/328)).
- **XLSX.** Shared-string run-level rich text ([#303](https://github.com/yfedoseev/office_oxide/issues/303)); inline-string whitespace ([#269](https://github.com/yfedoseev/office_oxide/issues/269)); inter-run spaces and scatter/bubble points ([#280](https://github.com/yfedoseev/office_oxide/issues/280), [#281](https://github.com/yfedoseev/office_oxide/issues/281)); formula text through the IR and renderer ([#279](https://github.com/yfedoseev/office_oxide/issues/279)); shared-formula follower cells keep their formula text ([#278](https://github.com/yfedoseev/office_oxide/issues/278)); `plain_text()` includes the chart text `to_markdown()` already had ([#331](https://github.com/yfedoseev/office_oxide/issues/331)); merged cells reach `to_ir()` ([#235](https://github.com/yfedoseev/office_oxide/issues/235)); threaded comments preferred over legacy boilerplate ([#301](https://github.com/yfedoseev/office_oxide/issues/301)); in-cell rich-value images instead of a fabricated `#VALUE!` ([#302](https://github.com/yfedoseev/office_oxide/issues/302)); date-serial conversion bounded, UTF-16 XML parts decoded ([#225](https://github.com/yfedoseev/office_oxide/issues/225), [#229](https://github.com/yfedoseev/office_oxide/issues/229)); accounting negatives, trailing literals and scientific exponents in number formats ([#234](https://github.com/yfedoseev/office_oxide/issues/234)); a cell keeps its column after a row-spanning merge ([#340](https://github.com/yfedoseev/office_oxide/issues/340)).
- **XLS.** Embedded chart's nested `EOF` no longer closes the parent sheet ([#237](https://github.com/yfedoseev/office_oxide/issues/237)); `DATEMODE` 1904 threaded through ([#233](https://github.com/yfedoseev/office_oxide/issues/233)); declared `CODEPAGE` applied to BIFF5 text ([#309](https://github.com/yfedoseev/office_oxide/issues/309)); record-parsing budget raised and truncation surfaced ([#236](https://github.com/yfedoseev/office_oxide/issues/236)); date walk bounded and format overrides honoured ([#225](https://github.com/yfedoseev/office_oxide/issues/225)); merged cells ([#235](https://github.com/yfedoseev/office_oxide/issues/235)); hidden and very-hidden sheets are kept and flagged rather than dropped, as XLSX already did ([#231](https://github.com/yfedoseev/office_oxide/issues/231)).
- **PPTX.** Speaker notes routed to `Section.speaker_notes` ([#238](https://github.com/yfedoseev/office_oxide/issues/238)); PowerPoint Sections' `p14:sldIdLst` no longer overwrites the slide list ([#228](https://github.com/yfedoseev/office_oxide/issues/228)); embedded chart text ([#239](https://github.com/yfedoseev/office_oxide/issues/239)); merged-cell tables rendered as a proper grid ([#289](https://github.com/yfedoseev/office_oxide/issues/289)); shape click actions and non-text AutoShape alt text reach the IR ([#299](https://github.com/yfedoseev/office_oxide/issues/299), [#300](https://github.com/yfedoseev/office_oxide/issues/300)); body/title placeholder content flows instead of being wrapped in a positioned `TextBox` ([#342](https://github.com/yfedoseev/office_oxide/issues/342)); `TextBox` written as its own positioned shape ([#264](https://github.com/yfedoseev/office_oxide/issues/264)); table cells and bullet items get real run formatting on write ([#341](https://github.com/yfedoseev/office_oxide/issues/341)).
- **PPT.** `TextTypeEnum` mapping from 3 upward — title/subtitle were swapped ([#253](https://github.com/yfedoseev/office_oxide/issues/253)); picture shapes resolve to their own slide via `pib` ([#256](https://github.com/yfedoseev/office_oxide/issues/256)); multi-paragraph text inside one text atom splits on PPT's bare `\r` delimiter ([#334](https://github.com/yfedoseev/office_oxide/issues/334)).
- **DOC.** `FibRgLw97` field offsets, shifted one slot for every `.doc` ([#247](https://github.com/yfedoseev/office_oxide/issues/247)); piece-table/FIB text-length gap surfaced instead of silently truncating ([#230](https://github.com/yfedoseev/office_oxide/issues/230)); FIB language id applied to compressed-text decoding ([#310](https://github.com/yfedoseev/office_oxide/issues/310)); deleted revision-mark text filtered out ([#288](https://github.com/yfedoseev/office_oxide/issues/288)); line-shape heading guess tightened against junk ([#224](https://github.com/yfedoseev/office_oxide/issues/224)).
- **Container / API.** CFB `find_entry`/`open_stream` resolve only the root storage's direct children ([#226](https://github.com/yfedoseev/office_oxide/issues/226)); macro-enabled and template extensions accepted ([#282](https://github.com/yfedoseev/office_oxide/issues/282)); encrypted files give a friendly error from every entry point ([#232](https://github.com/yfedoseev/office_oxide/issues/232), [#320](https://github.com/yfedoseev/office_oxide/issues/320), [#319](https://github.com/yfedoseev/office_oxide/issues/319)); `Element` gets an iterative `Drop`, closing the stack-overflow-on-drop gap ([#218](https://github.com/yfedoseev/office_oxide/issues/218)); `OfficeArtFBSE`-wrapped BLIPs in a Pictures/Data stream are no longer skipped, which also misaligned every later image index ([#336](https://github.com/yfedoseev/office_oxide/issues/336)); `with_parse_stack` bounds its in-flight parse threads instead of spawning without limit under concurrent load ([#313](https://github.com/yfedoseev/office_oxide/issues/313)); the C FFI documents its thread-safety contract for opaque handles ([#312](https://github.com/yfedoseev/office_oxide/issues/312)).
- **CLI / MCP.** `--version` flag ([#321](https://github.com/yfedoseev/office_oxide/issues/321)); an empty `replace` find string is rejected instead of inserting the replacement between every character ([#322](https://github.com/yfedoseev/office_oxide/issues/322)); the `ir` JSON projection carries `Note::marker` and run/paragraph formatting fields, and no longer silently swallows new IR variants ([#332](https://github.com/yfedoseev/office_oxide/issues/332), [#333](https://github.com/yfedoseev/office_oxide/issues/333), [#221](https://github.com/yfedoseev/office_oxide/issues/221)); JSON-RPC batch requests processed and the envelope validated ([#326](https://github.com/yfedoseev/office_oxide/issues/326), [#327](https://github.com/yfedoseev/office_oxide/issues/327)).
- **Found by the release regression sweep and quorum check** (v0.1.11 vs this branch over 10,599 files, then office_oxide vs Tika/POI, python-docx, pandoc, openpyxl, python-calamine, xlrd, catdoc, antiword, xls2csv and catppt). Regressions introduced by this release: the CLI's `ir` command built a `serde_json::Value` projection that cost 5–10 GB and OOM kills on large spreadsheets — it now streams the IR's serde form, the same bytes the Python/MCP/C/WASM surfaces emit ([#349](https://github.com/yfedoseev/office_oxide/issues/349)); fields nested inside another field's *result* (every TOC entry is a `HYPERLINK` field) were hidden along with the instruction text, so a table of contents became blank lines ([#350](https://github.com/yfedoseev/office_oxide/issues/350)); `plain_text()` for `.doc` picked up the footnote separator/continuation-notice stories ([#355](https://github.com/yfedoseev/office_oxide/issues/355)); style-inherited `w:vanish` was honoured by `to_html()` but not `plain_text()`/`to_markdown()` ([#356](https://github.com/yfedoseev/office_oxide/issues/356)). Defects present in v0.1.11 too: `.xls` shared strings cut by a `CONTINUE` boundary were misparsed, so any workbook with more than ~8 KB of unique strings lost most of its text cells — 80 corpus files at 2–58 % of what every reference extracts ([#353](https://github.com/yfedoseev/office_oxide/issues/353)); an embedded chart's cached series values overwrote the worksheet's own A1:Cn ([#352](https://github.com/yfedoseev/office_oxide/issues/352)); a text box wrapped in a content control was dropped whole ([#357](https://github.com/yfedoseev/office_oxide/issues/357)); PowerPoint 95 dual-storage files read the unreadable PPT 95 stream instead of `PP97_DUALSTORAGE/PowerPoint Document` ([#358](https://github.com/yfedoseev/office_oxide/issues/358)); a `.ppt` without its main stream opened as an empty deck with `Ok` ([#351](https://github.com/yfedoseev/office_oxide/issues/351)); `.xls`/`.xlsx` cell comments reached `to_ir()`/`to_html()` but not `plain_text()`/`to_markdown()` ([#354](https://github.com/yfedoseev/office_oxide/issues/354)). A 42 KB `.xls` declaring a 65,536 × 256 grid with four corner cells cost 5.1 GB in `to_ir()` (every empty row kept a 256-cell buffer behind it after its cells were popped) and 50 MB of markdown; now 3 MB and 3 KB ([#361](https://github.com/yfedoseev/office_oxide/issues/361)). `.xlsx` `plain_text()` names its sheets, as `.xls`, `to_markdown()` and every other spreadsheet reader do ([#359](https://github.com/yfedoseev/office_oxide/issues/359)); `Sheet::rows` is jagged instead of padded to the declared extent — `open()` on that 42 KB file drops from 392 MB to 12 MB ([#362](https://github.com/yfedoseev/office_oxide/issues/362)); the slot-leak test is deterministic ([#363](https://github.com/yfedoseev/office_oxide/issues/363)). Every fix ships a regression test that fails on the previous code.
- **Found by the release harness's cross-surface and round-trip axes** (`harness.py`, pass 4). Write side: every section's headers and footers were referenced from the last `sectPr` only, a section owning headers but no page setup got no `sectPr` at all, and the reader dropped an element-less trailing section — `docx/headers_footers/*` lost every header on a round trip ([#383](https://github.com/yfedoseev/office_oxide/issues/383)); a picture in a table cell, text box, list item, header, footer or note was skipped by the nested converter, and the reader never loaded pictures from header/footer/notes parts (their `r:id`s resolve against those parts' own rels) ([#384](https://github.com/yfedoseev/office_oxide/issues/384)); `.xlsx` image alt text was not written and every formula was written as its cached value — or as nothing, without one — now `<f>` plus `<v>` through `CellData::FormulaWithValue` ([#385](https://github.com/yfedoseev/office_oxide/issues/385)); a section title lifted from a heading that is not the first element rendered twice ([#386](https://github.com/yfedoseev/office_oxide/issues/386)). Read side: a compound file is read by the document stream it holds, so a workbook saved as `.doc` opens as a spreadsheet ([#387](https://github.com/yfedoseev/office_oxide/issues/387)); Word for Windows 1.x/2.0 flat files are read ([#388](https://github.com/yfedoseev/office_oxide/issues/388)); `.doc` paragraphs are cut at paragraph marks with each PAPX/CHPX run intersected with every piece — a fast-saved file's runs are not one CP range, so a paragraph could repeat its neighbour or start nowhere — and tracked deletions are left out of the structured view as they are of the flat text ([#389](https://github.com/yfedoseev/office_oxide/issues/389)); `.doc` `to_markdown()` renders from the structured view, with headings, lists and tables, and headings are no longer forced bold ([#390](https://github.com/yfedoseev/office_oxide/issues/390)); legacy `SummaryInformation` strings decode in the set's `PID_CODEPAGE` ([#391](https://github.com/yfedoseev/office_oxide/issues/391)); every markdown renderer escapes document text and writes a cell's line breaks as `<br>` ([#392](https://github.com/yfedoseev/office_oxide/issues/392)); `.ppt` soft returns are line breaks on the IR surfaces and the direct surfaces no longer leak `\r`/`\x0B` ([#393](https://github.com/yfedoseev/office_oxide/issues/393)); `.xlsx` prose-mode rows keep every cell ([#394](https://github.com/yfedoseev/office_oxide/issues/394)); a spreadsheet or document comment's cell and author reach the IR surfaces. A package with one damaged part no longer fails as a whole: an unreadable header/footer part is skipped with a warning, an unreadable sheet is skipped, named in every renderer's output and in `Metadata::text_truncated` ([#395](https://github.com/yfedoseev/office_oxide/issues/395)); a `.ppt` title placeholder holding several paragraphs is a heading and paragraphs, not one fused bold heading, and a free text box's paragraphs are not fused either ([#396](https://github.com/yfedoseev/office_oxide/issues/396)); a content control inside a `.docx` table cell keeps the cell's text, a `.doc` that ends inside a table cell keeps that cell, an `.xlsx` formula whose cached result is empty shows nothing on the IR surfaces as on `plain_text()`, and comments, embedded objects and picture alt text carry one label on every surface (`Comment (Author): …`, `[object]`, `[alt]`) ([#397](https://github.com/yfedoseev/office_oxide/issues/397)). A `.doc`'s declared title reaches `Metadata::title` only, not the section, and a section title is the whole first heading ([#399](https://github.com/yfedoseev/office_oxide/issues/399)). An `.xlsx` merge over rows the sheet does not store spans the stored rows it covers, not the rows that follow, and a `.pptx` image with a description but no bytes is written as a description-only shape and read back as the same image ([#398](https://github.com/yfedoseev/office_oxide/issues/398)). Earlier in the same pass: text after a hard break in the same paragraph reached `plain_text()` but not `to_ir()` ([#381](https://github.com/yfedoseev/office_oxide/issues/381)); drawings in bordered empty paragraphs, XLSX drawn text, PPTX slide comments and XLS comments on empty sheets each reached one renderer and not another ([#382](https://github.com/yfedoseev/office_oxide/issues/382)).
- **Found by the markdown-structure comparison against pandoc and the "peers read it, we don't" list** (this release's second verification pass, `render_compare.py` / `unparsed_vs_peers.py`). `to_markdown()` ignored list numbering inherited from the paragraph style — Word's own "List Bullet"/"List Number" — while `to_ir()`/`to_html()` resolved it, so a template with 157 bullet items rendered none ([#377](https://github.com/yfedoseev/office_oxide/issues/377)); the same renderer dropped bold/italic inherited from paragraph and character styles, and every span inside a table cell ([#378](https://github.com/yfedoseev/office_oxide/issues/378)); a heading style's own `<w:b/>` no longer renders as `<h1><strong>` / `# **…**` on any surface; a table nested in a one-cell layout frame is rendered in place of the frame on both markdown surfaces instead of being flattened into it — 77 rows where there was 1 ([#379](https://github.com/yfedoseev/office_oxide/issues/379)). The legacy template, show and add-in extensions (`.dot`, `.pot`, `.pps`, `.xlt`, `.xla`) open like their `.doc`/`.ppt`/`.xls` siblings ([#373](https://github.com/yfedoseev/office_oxide/issues/373)); a file shorter than a CFB header is reported as "not a compound file"; `office-oxide text f.docx | head` no longer panics on the closed pipe ([#372](https://github.com/yfedoseev/office_oxide/issues/372)). `Metadata::has_macros` recognises the `vbaProject` relationship type Office actually writes (the `schemas.microsoft.com` one) — it was `false` for every real `.docm`/`.xlsm`/`.pptm`, 0 of 70 corpus files; the earlier test had been built from the crate's own wrong constant ([#380](https://github.com/yfedoseev/office_oxide/issues/380)).
- Workspace test suite repaired ([#330](https://github.com/yfedoseev/office_oxide/issues/330)); the crate docs, three binding READMEs, the Rust getting-started guide and the root README stop overclaiming write/edit support for legacy formats and list the full CLI/MCP surface ([#323](https://github.com/yfedoseev/office_oxide/issues/323), [#324](https://github.com/yfedoseev/office_oxide/issues/324), [#325](https://github.com/yfedoseev/office_oxide/issues/325)).

### Performance

The first profiling pass over the corpus: every file on four surfaces with per-process CPU and peak RSS (`scripts/regression-sweep/perf_sweep.py`), synthetic documents that scale one structure at a time to expose superlinear paths (`gen_scaling.py` + `scale_run.py`), `perf` with DWARF call graphs on the outliers, and the same files through the reference readers for a baseline. Eight issues, each fixed with a regression test that fails on the old code — allocation counts from a counting global allocator where a timing assertion would be flaky ([#364](https://github.com/yfedoseev/office_oxide/issues/364)–[#371](https://github.com/yfedoseev/office_oxide/issues/371)).

- **Package open was quadratic in the number of parts** ([#364](https://github.com/yfedoseev/office_oxide/issues/364)). `OpcReader::has_part` and the `read_zip_entry` fallback rescanned the zip's central directory — allocating a normalised copy of every name — on every call, and the XLSX loader probes optional parts per sheet, so each *miss* was a full scan. A name index built once per archive: 8,000 one-cell sheets **56 s → 0.19 s**, 8,000 slides **19.9 s → 0.49 s**, 8,000 header parts 6.6 s → 0.24 s, 8,000 images 4.9 s → 0.22 s.
- **IR JSON no longer spends 17 bytes per image byte** ([#365](https://github.com/yfedoseev/office_oxide/issues/365)). `Image.data` went out as a pretty-printed number array, one line per byte: a 1.2 MB `.docx` was a 140 MB, 8.3-million-line `ir`. It is base64 now (11 MB, 1.09 s → 0.13 s), on every JSON surface; the old array still deserialises.
- **Attribute reads no longer rescan the tag or allocate** ([#366](https://github.com/yfedoseev/office_oxide/issues/366), [#368](https://github.com/yfedoseev/office_oxide/issues/368)). Every `optional_attr_str` re-parsed the attribute list from its start and paid quick-xml's per-call duplicate-check `Vec`, so every `w:val` read was a malloc and a four-attribute border edge cost sixteen attribute parses. One allocation-free pass (`xml::attrs`) under every attribute helper; the worksheet cell loop keeps the reference and type borrowed and parses a numeric `<v>` from the borrowed text. XLSX numeric cells: 3.78 → 0.12 heap allocations each.
- **A DOCX table cell no longer costs 4 KB** ([#367](https://github.com/yfedoseev/office_oxide/issues/367)). `BlockElement` was 536 bytes (an unboxed `Table`), `RunContent` 224 (an unboxed `DrawingInfo`), and every one-element content `Vec` kept the four slots `push` reserves. Boxed and shrunk: a 128,000-cell table opens in **90 MB instead of 524 MB** and converts to IR in 140 MB instead of 241 MB; an 8,000-row table's `text` 1.05 s → 0.14 s.
- **Compound-file streams are read in runs, not sectors** ([#369](https://github.com/yfedoseev/office_oxide/issues/369)). `read_chain` allocated, seeked and read once per 512-byte sector of an unbuffered `File` — ~28,000 syscalls for a 14 MB `.doc` — then copied the stream again. Consecutive sectors are fetched with one read each into a buffer reserved once.
- **Pictures decode on first request, not at `open()`** ([#370](https://github.com/yfedoseev/office_oxide/issues/370)). DOC's `Data` stream, PPT's `Pictures` stream and — byte by byte — XLS's whole Workbook stream were scanned for BLIPs even for `plain_text()`: 10–17 % of a text extraction on picture-heavy files. Lazy now; XLS reads only the `MsoDrawingGroup`/`MsoDrawing` payloads its record walk already passes over, where the format keeps them.
- **XLS grid rows are sized once and numbers format with one allocation** ([#400](https://github.com/yfedoseev/office_oxide/issues/400)). `build_grid` grew each row cell by cell and `format_commas` allocated three times per value, which made million-cell workbooks 2–4× slower than v0.1.11 on the harness's timing axis; `govdocs1_000032.xls` (29 MB) `text` 1.8 s → 0.46 s.
- **XLS no longer copies every record and every cell's text** ([#371](https://github.com/yfedoseev/office_oxide/issues/371)). `RecordIter` copied each BIFF record into a `Vec`, `parse_cell_record` returned a `Vec` per cell, and a rendered-text grid held every cell's string at `open()` — a string cell three times over, and the allocator half of a large workbook's runtime. Records borrow the stream, cells append, and text renders on demand (see *Changed*). `text` on the largest corpus workbooks 2–2.5× faster at 40 % less memory.

Corpus totals (CPU over every file that opens, this branch before → after): XLSX **0.41×**, PPT `ir` 0.15×, PPTX `ir` 0.35×, DOC `ir` 0.42×, XLS `text` 0.52×, DOCX 0.7×; the largest DOCX peak RSS 104 → 51 MB. Against the reference readers on the same files office_oxide is now the fastest in the panel: 0.59× calamine's total CPU (1.36× before), 0.51× catppt, 0.49× xls2csv, 0.48× antiword, 0.03× openpyxl, 0.05× pandoc. Each scaling dimension probed is linear (log₂ slope ≤ 1.2 at 8,000–32,000 elements).

### Security

- **Spreadsheets: a per-document text budget bounds shared-string fan-out.** A shared string referenced from every cell is rendered once per cell, and nothing bounded that product: a 110 KB `.xlsx` holding one 1 MB string in 12,000 cells rendered to 393 MB on every surface, and the same string from a million rows would have been 32 GB from a 300 KB file — through every binding and the MCP server. Every XLSX/XLSB renderer, the IR converter and the `.xls` reader (which copies the string into each cell at `open()`) now charge the characters they emit against `office_oxide::limits::max_text_chars()` — 256 Mi by default, the shape of POI's `ZipSecureFile.setMaxTextSize` — and stop with a visible `[output truncated: …]` notice (`.xls` also sets `Metadata::text_truncated`). `set_max_text_chars` raises it for processes that read larger workbooks. The default keeps every genuine workbook in the 10,599-file corpus (the largest renders 181 MB of text) and changed exactly the two proof-of-concept files.
- **XLSX/XLSB: a cell reference past the grid sized the IR table to it.** `CellRef::parse` accepted any run of column letters, so `ZZZZZZ1` was column 321,272,405 and `to_ir()` allocated every table row that wide — a 95 GB allocation abort from a 2 KB `.xlsx`, through every binding and the MCP server. References past `XFD1048576` are refused (the cell takes the implied column), the `.xlsb` reader drops such cells, and the converter caps the grid width regardless. Found by the mutation pass over the new `.xlsb` reader.
- **CFB (`.doc`/`.xls`/`.ppt`): a cyclic DIFAT chain hung `Document::open` forever.** The DIFAT walk in `CfbReader::read_fat` was the one chain-follower in the reader without a cycle guard; a 1.5 KB file whose DIFAT sector pointed at itself never returned, with no crash signal for a supervisor to act on. The walk is now bounded by the sectors the file holds and reports `DIFAT chain cycle detected`.
- **CFB: a cyclic directory tree overflowed the stack.** `find_in_tree` recursed along sibling pointers without a visited set; since this release routes `find_entry`/`open_stream` through it (the embedded-object shadowing fix), a malformed directory aborted every legacy open. The search is iterative and visits each entry once.
- **XLSX: the fast-path reader skipped the duplicate-zip-entry refusal** that `OpcReader` applies for DOCX/PPTX (the CVE-2025-31672 shape: a scanner reads one copy of `sheet1.xml`, the consumer renders the other). All three OOXML readers now open their archive through one shared check.

### Changed

- Ticket numbers removed from code comments (they belong here), 40 `allow(dead_code)` suppressions removed with the dead code they hid (including an unreachable 190-line XLSX OPC parser), every test function named `test_*`. Enforced by the new `house-rules` CI job.
- `.research/` review notes are no longer tracked.
- **IR JSON omits fields at their default.** Every `None` and every `false` (and `TableRow::allow_break` at `true`) is left out of the serialized IR on every surface — Python, MCP, C FFI, WASM and the CLI's `ir` — and absent keys deserialize to the same value. A 14 MB `.xls` went from 909 MB to 207 MB of pretty-printed JSON. Consumers that index a key rather than `.get()` it must use the default when it is absent ([#360](https://github.com/yfedoseev/office_oxide/issues/360)).
- **IR JSON carries image bytes as a base64 string** (`"data": "iVBORw0KGgo…"`) on every surface instead of an array of numbers; both shapes deserialise ([#365](https://github.com/yfedoseev/office_oxide/issues/365)).
- **`office_oxide::limits`** is new: `max_text_chars()` / `set_max_text_chars()` and `DEFAULT_MAX_TEXT_CHARS` (256 Mi). A spreadsheet whose rendered text exceeds the budget is truncated with a notice on every surface; raise the limit once at startup to read larger workbooks whole.
- **`xls::Sheet::display` is gone; use `Sheet::display_text(row, col)`**, which returns the cell's text as Excel shows it (number formats and dates applied) as a `Cow<str>`, formatted on demand from the new `Sheet::xf` grid and the workbook's shared `NumberFormats`. `CellValue::as_text_cow` is the borrowing form of `as_text` ([#371](https://github.com/yfedoseev/office_oxide/issues/371)). `docx::BlockElement::Table` and `docx::RunContent::Drawing` hold a `Box` ([#367](https://github.com/yfedoseev/office_oxide/issues/367)).
- **`xlsx::write::CellData::FormulaWithValue { formula, cached }`** is new (a formula with its last computed value, written as `<f>` plus `<v>`), and `SheetImage` gains `alt_text` with `add_image_with_alt` beside `add_image` ([#385](https://github.com/yfedoseev/office_oxide/issues/385)). **`doc::DocDocument::to_markdown()` renders from the structured view** — headings, lists, tables, headers/footers and notes — instead of the flat text split into blocks ([#390](https://github.com/yfedoseev/office_oxide/issues/390)). **Every `to_markdown()` escapes document text** (`\|`, `\<`, `\*`, `\[` …) and writes a table cell's line breaks as `<br>` ([#392](https://github.com/yfedoseev/office_oxide/issues/392)); a picture's `![alt]` is one line. Comments carry their cell and author on the IR surfaces (`A1 (Jane): …`).
- **`.xlsx` `plain_text()` starts each sheet with its name**, like `.xls` and `to_markdown()` ([#359](https://github.com/yfedoseev/office_oxide/issues/359)). `xls::Sheet::rows` is jagged — index with `.get()` ([#362](https://github.com/yfedoseev/office_oxide/issues/362)).
- **quick-xml 0.41 → 0.42** — the parser's names, attributes and text events are now `&str`; invalid UTF-8 in an XML part is replaced by U+FFFD (with a warning) rather than failing the part, since 0.42 validates UTF-8 per event.
- Every other dependency and pinned GitHub Action bumped to its latest release: `cargo update` across the workspace and bench lockfiles, bench_rust `calamine` 0.36, `koffi` 3.3.1, `Microsoft.NET.Test.Sdk` 18.10.1 and `xunit.runner.visualstudio` 4.0.0, `uv.lock` (`ruff` 0.16.8), the bench pins `markitdown` 0.1.7 / `python-docx` 1.2.0 / `xlrd` 2.0.2 (the BENCHMARKS.md table predates these pins), and SHA-pinned `action-gh-release` 3.0.3, `codeql-action` 4.37.4, `setup-uv` 10.1.0, `codecov-action` 7.1.1, `rust-cache` 2.9.2, `install-action` 2.87.16, `rust-toolchain` current. `cargo audit` and `cargo deny` clean.

## [0.1.11] - 2026-09-11

> Write-path correctness. 16 issues closed and the OOXML validation gate that found them.
>
> v0.1.10 shipped with a green suite and a write path that produced documents Word and PowerPoint refuse to open: nothing checked that the files we *write* are valid OOXML, because generated output was only round-tripped through our own deliberately lenient parser. #199 came in from a user; chasing it with a schema validator turned up fifteen more.
>
> Verified against 6,062 real Office documents arm-to-arm with v0.1.10 — **0 extraction regressions**, and `save_as` conversions that schema-validate go from **2,387/5,921 (40.3%) to 5,918/5,921 (99.9%)**. The sweep itself found nine further defects the unit suite could not see, four of them introduced by fixes in this release.

### Verified

Swept **6,062 real Office documents** (LibreOffice QA, Apache POI, OpenXML SDK, python-pptx, ClosedXML, calamine, PhpSpreadsheet and others) arm-to-arm against v0.1.10, on two axes.

**Extraction — no regressions.** 6,062 files × 4 surfaces = 24,248 observations per arm. **0 ok→not-ok**, 3 recoveries. The only text change is one file where a time that rounded up to a whole day now carries into the date instead of rendering hour 24. Three anomalous status transitions were re-measured quiescently and were load artifacts, not regressions — the sweep README warns about exactly that.

**Conversion — the axis this release changes.** Every file through `Document::save_as`, then schema-validated:

| source | v0.1.10 | v0.1.11 |
|---|---|---|
| `.doc` → `.docx` | 92 / 201 | **201 / 201** |
| `.docx` → `.docx` | 60 / 2509 | **2507 / 2509** |
| `.ppt` → `.pptx` | 0 / 176 | **176 / 176** |
| `.pptx` → `.pptx` | 0 / 791 | **791 / 791** |
| `.xls` → `.xlsx` | 469 / 475 | **475 / 475** |
| `.xlsx` → `.xlsx` | 1766 / 1769 | **1768 / 1769** |
| **total** | **2387 / 5921 (40.3%)** | **5918 / 5921 (99.9%)** |

Zero regressions. The three remaining files fail identically under v0.1.10: two carry a malformed `dcterms:modified` in the source (`2015sss-06-20T07:40:00Z`) that is copied through verbatim, and one is a deeply nested table.

The sweep found **nine defects the unit suite could not see**, all fixed here — including three element-ordering defects of the same shape (`w:pPr`, `w:tblPr`, `w:tcPr`), two namespace declarations missing from parts that use them, and one duplicate attribute I introduced while fixing a missing one.

### Added

- **An OOXML validation gate ([#201](https://github.com/yfedoseev/office_oxide/issues/201)).** Nothing checked that the documents this library *writes* are valid — generated files were only round-tripped through our own lenient parser, so a document Word refuses to open passed the whole suite. `scripts/ooxml-validate/` now fetches the ISO/IEC 29500-4 schemas (not committed, per CONTRIBUTING #4), `cargo run --example gen_validation_corpus` produces 30 packages across the markdown path, the builder APIs, all nine `save_as` conversion pairs and out-of-range values, and both validators run in CI. Writing it immediately found three more defects: `w:titlePg` emitted in the wrong `CT_SectPr` position, `xml:space` on an XLSX `<t>` where it is not allowed, and an empty `p:txBody` on a slide whose only content was a table.

### Security

- **A sub-1 KB `.docx` could abort the process (`CWE-770`).** `w:gridSpan` was parsed as an unbounded `u32` and summed straight into an allocation, so a 978-byte file could request 34 GB and abort. Measured: 1e6 → 11 MB, 8e6 → 64 MB, `u32::MAX` → abort, i.e. ~35,000,000× amplification. This ran on `to_ir()`, which backs `save_as`, `to_markdown`, the MCP `extract` tool and every binding, and an `abort()` is not catchable — no caller could defend against it. Spans are now clamped where the IR cell is built, in the converter's column count, and defensively in `ir_render::table_grid`, since `DocumentIR` is `Deserialize` and can arrive from an untrusted source.
- **A hostile `.pptx` could hang `save_as` indefinitely.** `col_span`/`row_span` drove the DOCX writer's grid loop with the bounds check *inside* the loop body, so the iteration still ran `row_span × col_span` times — up to 1.8×10¹⁹ — doing nothing. Nothing allocated, so nothing ever stopped it. The check is now hoisted and the spans clamped, and PPTX spans are clamped at parse time.
- **A deeply nested IR overflowed the writer's stack.** The readers have bounded nesting via `DepthGuard`/`MAX_NESTING_DEPTH`; the writers had no equivalent, so nesting past a few thousand levels aborted. The same guard now covers both DOCX writer recursions.

- **A negate overflow panicked the writer.** `first_line_indent_twips = i32::MIN` reached `(-v).to_string()` and `-i32::MIN` overflows; the reader had the mirror problem on `w:hanging="-2147483648"`. Both now use the magnitude, which is what `ST_TwipsMeasure` wants — the attribute is unsigned.

Known and not fixed: dropping a deeply nested `Element` is itself recursive and can overflow a small stack independently of the writers. That affects any consumer deserialising such an IR and needs an iterative `Drop`.

### Fixed

- **Two regressions caught by the v0.1.10 → v0.1.11 corpus sweep**, both introduced by fixes in this release. Speaker notes reached `plain_text` and markdown but not HTML, because moving them out of `Section::elements` for [#203](https://github.com/yfedoseev/office_oxide/issues/203) dropped them from the one renderer that was never updated; HTML now emits them in a labelled `<aside class="speaker-notes">`, which is better than v0.1.10, where they were indistinguishable from slide body text. And [#207](https://github.com/yfedoseev/office_oxide/issues/207)'s built-in format resolution widened `apply_format`'s custom-code branch, so a built-in code was fed to `apply_custom` and rendered literally — a cell formatted `mm:ss.0` came out as `mm:ss0.6`. The render paths now pass the workbook's *declared* code only.

- **The PPTX writer flattened structure into text ([#214](https://github.com/yfedoseev/office_oxide/issues/214)).** Tables were joined with literal tabs and newlines into a single run, losing the grid and every cell boundary; they are now written as a real `a:tbl` in a `p:graphicFrame`. Bullet lists were emitted at level 0 with no `marL`/`indent`, so nesting was invisible and the glyph sat at the same x as its text. Footnote and endnote bodies fell into a catch-all and were dropped entirely.

- **The markdown front end dropped fenced code blocks and list nesting ([#215](https://github.com/yfedoseev/office_oxide/issues/215)).** A ` ```rust ` fence produced no `Element::CodeBlock` at all — the language leaked into the text as an ordinary paragraph — and `ListItem::nested` was hardcoded to `None`, so every bullet flattened to level 0. The second of those is also why the earlier "3-level nested lists" coverage could not have caught [#211](https://github.com/yfedoseev/office_oxide/issues/211): the markdown path never populated the field.

- **Every binding discarded the writer's "value not written" signal ([#209](https://github.com/yfedoseev/office_oxide/issues/209)).** `sheet_set_cell` and the PPTX slide setters return `bool` precisely so a silent no-op is detectable — the Rust API was fixed for that regression and all six bindings reintroduced it. The C FFI entry points now return a status, Python raises `IndexError`, and C# throws; an unrecognised FFI `value_type` is rejected instead of erasing the existing cell; and `CellData::Boolean`/`Formula`, previously unreachable from any binding, have `value_type` 3 and 4. Go silently wrote an **empty** cell for `int64`, `float32`, `uint` and a styled `bool`, and C# used a locale-dependent `ToString()` that turned a `decimal` into text (`"1,5"` under `de-DE`); both now cover the full numeric set, with C# using the invariant culture.

- **`replace_text` reported success for XLSX, where it is not implemented ([#210](https://github.com/yfedoseev/office_oxide/issues/210)).** It returned `0`, which is indistinguishable from "the text was not present", so the CLI exited 0, the MCP tool returned a non-error result, and the file was rewritten anyway — with the MCP schema advertising XLSX support. It now returns an error naming the unsupported format, and the MCP tool description no longer claims XLSX. The edit path is also byte-deterministic again: parts, part relationships and content-type overrides were iterated out of `HashMap`s, so saving an unchanged document produced a different byte stream every time.

- **IR fields the API accepted were never emitted ([#213](https://github.com/yfedoseev/office_oxide/issues/213)).** `TextSpan::hyperlink` was read and explicitly discarded, so a markdown link's URL was not recoverable from the output at all — DOCX now emits `w:hyperlink` with an external relationship, and run coalescing treats the URL as part of a run's identity so a link is no longer swallowed by an adjacent plain run. PPTX dropped `InlineContent::LineBreak` (joining the words on either side), underline and strikethrough, and image alt text; XLSX and PPTX omitted `xml:space="preserve"`, so leading and trailing spaces were stripped by conformant consumers; `Paragraph::tabs` reached no writer, costing dot-leader tables of contents both their leaders and their alignment; and `frame_position`, `Section::background_rgb` and `Table::caption` were likewise dropped — the caption was emitted only as a `Caption`-styled paragraph, so it was lost on round-trip while accumulating a phantom body paragraph on every cycle.

- **The XLSX write bridge discarded typed cell data ([#212](https://github.com/yfedoseev/office_oxide/issues/212)).** It re-parsed the *rendered* string instead of using the `data_type`, `raw_number` and `number_format` the reader had recorded, so on any XLSX round trip `"007"` became `7`, currency and date cells became text, and a cell whose text reads `inf` or `NaN` became an Excel error cell. `ir_to_xlsx` also ended in a catch-all that swallowed lists, code blocks, notes and everything inside a text box — and since PPTX wraps slide bodies in a text box, PPTX → XLSX lost almost all of its text. `CellStyle::number_format_code` is new, so a format read from a source document is carried through instead of flattened to `General`.

- **Nested content was silently dropped on write ([#211](https://github.com/yfedoseev/office_oxide/issues/211)).** `ListItem::nested` was consumed by the renderers but by no writer, so every list item below level 0 vanished from the output while the API reported success — on any document opened and re-saved. A heading nested in a table cell, text box, header or note also lost its alignment, unlike the identical heading at top level.

- **Generated DOCX carried dangling cross-part references ([#208](https://github.com/yfedoseev/office_oxide/issues/208)).** These are schema-valid and still make Word report unreadable content, so neither the XSDs nor any existing test caught them. A list in a header, footer or note emitted `w:numId`/`ListParagraph` while the numbering part and the style were gated on the body alone; the note-reference character styles were gated on a note part existing, though a run can carry a reference without one; `footnotes.xml` never contained the separator notes Word expects; a `w:type="first"` header was written without `w:titlePg` and a `w:type="even"` one without `settings.xml`, so neither was ever shown; a multi-section document emitted duplicate `w:type` values in the final `sectPr`; and an ordered list nested in a cell, text box or header used the bullet definition, disagreeing with the same list at top level.

- **A workbook's own `numFmt` override was ignored for built-in ids ([#207](https://github.com/yfedoseev/office_oxide/issues/207)).** [ECMA-376] §18.8.30 lets a workbook redefine ids 0-163, but the date check consulted the id before the declaration, so a number formatted `0.00" kg"` under id 14 was read back as `1900-01-02`. `StyleSheet::number_format_for` now also resolves built-in ids instead of returning `None` for the majority of real formatted cells, with `number_format_override_for` exposing the explicit declaration alone. Separately, a time fraction that rounds up to a whole day now carries into the date rather than producing hour 24, which no date library accepts.

- **The XLSX writer produced unparseable, invalid or silently wrong output on unusual input ([#206](https://github.com/yfedoseev/office_oxide/issues/206)).** A control character in a sheet name made `xl/workbook.xml` unparseable — `sanitize_xml_text` guarded element text but had never been applied to an attribute. `merge_cells` overflowed `usize` and panicked inside `save()`. `add_row` bypassed the grid check `set_cell` enforces and wrote cells past column XFD. Colours, column widths and font sizes were written through unvalidated (`inf`, `FFred`, `sz="0"`); half-point font sizes were truncated by integer division; a zero-sheet workbook emitted an invalid empty `<sheets/>`; a styled empty cell silently dropped its style; a formula kept a caller's leading `=`; and `styles.xml` was non-deterministic because it iterated a `HashMap`.

- **The XLSX edit path corrupted workbooks ([#205](https://github.com/yfedoseev/office_oxide/issues/205)).** A self-closing `<row .../>` — what Excel writes for a row carrying only a custom height — made `set_cell` insert the new cell into the **next** row, so a cell whose reference said row 1 ended up inside row 2. The same fix class had already been applied to self-closing `<c>` and was never carried over to `<row>`. Cells and rows are now also inserted in ascending order as [ECMA-376] §18.3.1.73 requires, control characters are stripped, and non-finite numbers become error cells instead of raw `NaN`/`inf` — both guards the writer already had and the editor lacked.

- **PPTX attribute values were written without range validation ([#204](https://github.com/yfedoseev/office_oxide/issues/204)).** `p:sldSz`, `a:rPr/@sz`, `a:srgbClr/@val` and `a:ext` all took caller input straight into restricted XSD types. `font_size(f64::NAN)` became `sz="0"` through a saturating float cast, and `.color("#FF0000")` — the form every CSS-adjacent API accepts — produced an invalid file. The slide-size case needed no explicit API call: the IR bridge converts a source page size straight to EMU, so any page under an inch produced an invalid deck.

- **Speaker notes were promoted onto the visible slide ([#203](https://github.com/yfedoseev/office_oxide/issues/203)).** The PPTX converter appended notes to the slide's `elements` as an ordinary paragraph, so every writer treated them as body text and a round trip published a presenter's private notes to the audience with no marker. Notes now live in `Section::speaker_notes` and round-trip into `ppt/notesSlides/`; the renderers still surface them, labelled.

- **DOCX documents were schema-invalid in six ways ([#200](https://github.com/yfedoseev/office_oxide/issues/200)).** `w:tblGrid` was written only when explicit column widths were known, so every markdown table was invalid; the simple `add_table` writer emitted neither `w:tblPr` nor `w:tblGrid`; `numbering.xml` interleaved `w:abstractNum` and `w:num` instead of writing all of the former first; both floating-anchor writers put `wp:docPr` before the wrap element; `w:pgMar` omitted the required `w:gutter`; and `w:pPr` children were emitted in the wrong order. The `w:pPr` fix covers more than the reported `spacing`/`ind` pair — `w:numPr`, `w:pBdr` and `w:shd` were also misplaced, and the whole block now follows the `CT_PPrBase` sequence.

- **PPTX packages were missing the theme, `presProps.xml` and the layout→master relationship ([#202](https://github.com/yfedoseev/office_oxide/issues/202)).** The slide layout had no `_rels` part at all, so it was orphaned from its master — a hard `shall` in [ISO/IEC 29500-1] §13.3.9 and a likely reason PowerPoint's repair failed rather than succeeded. There was also no theme, which left [#199](https://github.com/yfedoseev/office_oxide/issues/199)'s colour map naming slots that did not exist.

- **PPTX slide masters were missing the required `p:clrMap` ([#199](https://github.com/yfedoseev/office_oxide/issues/199)).** `CT_SlideMaster` is a strict sequence — `cSld`, `clrMap`, `sldLayoutIdLst` — and the colour map was never written at all, so every deck this library produced was schema-invalid and PowerPoint's repair had no colour mapping to recover.

## [0.1.10] - 2026-09-09

> Correctness release. 69 issues closed, concentrated in one defect shape: **the parser read a value correctly and the converter then dropped it**. Every format is affected; DOCX most of all. Also closes six security-relevant robustness gaps, adds editing to the WASM/MCP/CLI surfaces, and replaces several silent empty-successes with named errors. No breaking API changes; some previously-empty fields are now populated, and some previously-`Ok(empty)` reads are now `Err`.
>
> Verified against a 3,036-file corpus of real Office documents (Apache POI, Tika, LibreOffice QA, ClosedXML, OpenXML SDK test suites) compared arm-to-arm against v0.1.9: zero panics, zero crashes, zero timeouts, and no content regression. The harness that did it ships in `scripts/regression-sweep/`.

### Added

- **Editing on every surface ([#62](https://github.com/yfedoseev/office_oxide/issues/62)).** `replace_text` and save now reach the CLI, MCP server and WASM build, not just the Rust and Python APIs.
- **Inline base64 images in markdown ([#100](https://github.com/yfedoseev/office_oxide/issues/100)).** `to_markdown_with(MarkdownOptions { image_embed: ImageEmbed::Base64 })` emits `[image-base64:...]` at each image's position in the flow; the MCP `extract` tool exposes it as the `markdown-with-images` format.
- **Per-platform npm packages ([#61](https://github.com/yfedoseev/office_oxide/issues/61)).** The JS package now uses `optionalDependencies` for its six platform binaries instead of shipping all of them to every install.
- **Document metadata is read ([#174](https://github.com/yfedoseev/office_oxide/issues/174)).** `ir::Metadata` title, author and dates were always `None` even though `CoreProperties::parse` existed and worked.
- **Footnotes, endnotes, comments and speaker notes ([#142](https://github.com/yfedoseev/office_oxide/issues/142), [#143](https://github.com/yfedoseev/office_oxide/issues/143), [#165](https://github.com/yfedoseev/office_oxide/issues/165), [#195](https://github.com/yfedoseev/office_oxide/issues/195)).** DOCX `footnotes.xml`/`endnotes.xml`/`comments.xml`, XLSX `xl/comments*.xml`, PPTX `ppt/comments/*` and `notesSlides`, and the `.doc` subdocuments the FIB's `ccp*` lengths had been parsed for and never used.
- **Legacy images reach the IR ([#186](https://github.com/yfedoseev/office_oxide/issues/186)).** `.doc`/`.xls`/`.ppt` extracted pictures were being dropped during conversion.

### Fixed

#### Content silently dropped on read

- **XML entities and character references were deleted from text ([#156](https://github.com/yfedoseev/office_oxide/issues/156)).** `AT&T` extracted as `ATT`. quick-xml emits entity references as `Event::GeneralRef`, a separate event from `Event::Text`, and 30 handlers across every format matched only the latter. Attribute values were affected too — the root cause behind [#147](https://github.com/yfedoseev/office_oxide/issues/147). This one change recovers apostrophes, ampersands and non-ASCII characters across the whole corpus.
- **DOCX text-box content ([#140](https://github.com/yfedoseev/office_oxide/issues/140), [#102](https://github.com/yfedoseev/office_oxide/issues/102)) and transparent paragraph wrappers ([#141](https://github.com/yfedoseev/office_oxide/issues/141)).** `w:txbxContent` was never read; `w:ins`, `w:fldSimple`, `w:ruby` and `w:smartTag` text was dropped on the floor. Both returned `Ok`.
- **DOCX `w:cr`, `w:noBreakHyphen` and `w:sym` ([#152](https://github.com/yfedoseev/office_oxide/issues/152)).** `e-mail` extracted as `email`.
- **DOCX `altChunk`, image alt text and HYPERLINK field URLs ([#155](https://github.com/yfedoseev/office_oxide/issues/155)); internal `w:anchor` hyperlinks and `w:tblCaption` ([#189](https://github.com/yfedoseev/office_oxide/issues/189)).**
- **XLSX cell hyperlinks ([#164](https://github.com/yfedoseev/office_oxide/issues/164)).** Parsed into `Worksheet::hyperlinks`, then never rendered — while DOCX and PPTX both emitted links.
- **PPTX SmartArt and chart text ([#149](https://github.com/yfedoseev/office_oxide/issues/149)); `a:tbl` merge continuation cells ([#148](https://github.com/yfedoseev/office_oxide/issues/148)).** `graphicFrame` handled only the table URI, and dropped `vMerge`/`hMerge` cells shifted every cell to their right one column left.
- **DOCX `mc:AlternateContent` extracted both `Choice` and `Fallback` ([#163](https://github.com/yfedoseev/office_oxide/issues/163)) — shape text came out twice.**

#### Formatting parsed and then discarded

- **DOCX style-based formatting was never applied ([#194](https://github.com/yfedoseev/office_oxide/issues/194)).** Only direct `w:rPr`/`w:pPr` survived, so an ordinarily-styled document lost all of its formatting.
- **DOCX run and paragraph properties ([#181](https://github.com/yfedoseev/office_oxide/issues/181), [#182](https://github.com/yfedoseev/office_oxide/issues/182)).** Indent, spacing, line spacing and keep/page-break flags were lost on read — only alignment survived; underline, highlight, `vertAlign`, caps and smallCaps left half of `TextSpan`'s fields permanently empty.
- **DOCX table geometry ([#180](https://github.com/yfedoseev/office_oxide/issues/180)) and borders ([#190](https://github.com/yfedoseev/office_oxide/issues/190)).** `tblW`, `gridCol`, `tcW`, `tblCellMar` and `trHeight` were all lost, so the library could not read back the widths of tables it had written; `w:pBdr` was narrowed to a boolean and `tblBorders`/`tcBorders` were not parsed at all.
- **DOCX `w:themeColor` never resolved and discarded the `w:val` fallback ([#175](https://github.com/yfedoseev/office_oxide/issues/175)).** `resolve_color` existed and was never called.
- **DOCX headings, lists and breaks ([#154](https://github.com/yfedoseev/office_oxide/issues/154), [#187](https://github.com/yfedoseev/office_oxide/issues/187), [#188](https://github.com/yfedoseev/office_oxide/issues/188), [#185](https://github.com/yfedoseev/office_oxide/issues/185), [#135](https://github.com/yfedoseev/office_oxide/issues/135)).** Headings were identified only via `outlineLvl`, missing style-name and `Heading1..9` conventions; `w:numId="0"` (explicitly *no* numbering) was turned into a bullet list; `List::start_number` and `List::style` were never populated, so a list starting at 5 rendered as 1 and `a) b) c)` was indistinguishable from `1. 2. 3.`; page breaks were reported as `ThematicBreak` and column breaks dropped; `w:outlineLvl="9"` (body text) rendered as a heading.
- **DOCX sections ([#177](https://github.com/yfedoseev/office_oxide/issues/177), [#178](https://github.com/yfedoseev/office_oxide/issues/178), [#191](https://github.com/yfedoseev/office_oxide/issues/191)).** Column layout was written in full and read back with only the count; first-page and even-page headers were merged into one and every section received every section's headers; `Section::break_type` was hardcoded from the section index, with continuous/nextPage inverted.
- **XLSX ([#184](https://github.com/yfedoseev/office_oxide/issues/184), [#179](https://github.com/yfedoseev/office_oxide/issues/179), [#193](https://github.com/yfedoseev/office_oxide/issues/193)) and PPTX ([#183](https://github.com/yfedoseev/office_oxide/issues/183), [#150](https://github.com/yfedoseev/office_oxide/issues/150), [#192](https://github.com/yfedoseev/office_oxide/issues/192)).** Cell font formatting reached the IR in prose mode but not table mode — the same cell rendered differently depending on the shape of the sheet around it; landscape orientation was not applied to `paperSize` dimensions; hidden sheets and slides were extracted with no indication; PPTX run underline, font name, baseline, caps and spacing were never parsed; level-0 bullets rendered as plain paragraphs and `a:buAutoNum` numbering was lost; `TableRow::is_header` was set from row position rather than `a:tblPr/@firstRow`.

#### Wrong values

- **XLSX number formats ([#147](https://github.com/yfedoseev/office_oxide/issues/147)).** Quoted literals in custom formats were read as date tokens, so `12,500,000` rendered as `36123-11-01`. Format sections are now selected by their condition rather than by position, `?` is treated as a digit placeholder, and a format's code is no longer printed in place of its value.
- **XLSX cell placement ([#146](https://github.com/yfedoseev/office_oxide/issues/146)).** The column reference was discarded, so sparse or unordered rows put every value under the wrong header. Cells whose `r` attribute is absent now take the next column, and rows without `r` are numbered in document order, per ECMA-376 18.3.1.4 and 18.3.1.73.
- **XLS dates ([#172](https://github.com/yfedoseev/office_oxide/issues/172)).** `FORMAT` and `XF` records were never parsed, so dates extracted as raw serials (`38971` instead of a date).
- **`.doc` outline levels ([#139](https://github.com/yfedoseev/office_oxide/issues/139)).** The line-shape heading heuristic now defers to real outline data where it exists.

#### Silent success on unreadable input

- **Word 6.0/95 files were accepted and parsed with Word 97 offsets ([#167](https://github.com/yfedoseev/office_oxide/issues/167)).** 426,450 characters extracted as an empty string with `Ok`. Now a named error.
- **Encrypted legacy files ([#169](https://github.com/yfedoseev/office_oxide/issues/169)) and encrypted OOXML packages ([#119](https://github.com/yfedoseev/office_oxide/issues/119), partial).** Both returned empty text or a corrupt-zip failure; both now say the file is password-protected. Reading and writing ECMA-376 agile encryption is **not** implemented — see *Known gaps*.
- **Valid real-world `.doc`/`.ppt` files extracted as an empty string with `Ok` ([#168](https://github.com/yfedoseev/office_oxide/issues/168)).** The reported cause — a three-piece piece table — parses correctly and now has a test pinning it. The real defect was three silent empty-with-`Ok` returns in the `.doc` reader, which now name the structure that failed.
- **OOXML robustness ([#145](https://github.com/yfedoseev/office_oxide/issues/145), [#176](https://github.com/yfedoseev/office_oxide/issues/176)).** Truncated XML, wrong-format packages and duplicate parts all returned `Ok`; a malformed theme part made the whole document unreadable while a missing theme was fine.
- **`with_parse_stack` converted every internal panic into "unsupported format" ([#161](https://github.com/yfedoseev/office_oxide/issues/161)),** hiding real bugs and blinding the fuzz target.

#### Security and robustness

- **Unbounded parser recursion ([#151](https://github.com/yfedoseev/office_oxide/issues/151)).** A 1.6 KB `.docx` aborted the process with a stack overflow no caller could catch. Nesting is now capped at an empirically calibrated depth, with the excess reported rather than silently truncated.
- **No decompression limit on OOXML parts ([#144](https://github.com/yfedoseev/office_oxide/issues/144)).** A 2 MB `.docx` expanded to 2 GiB in memory and returned `Ok`. Parts are now bounded.
- **Emitted URLs were not scheme-filtered and markdown output was unescaped ([#157](https://github.com/yfedoseev/office_oxide/issues/157)).** `javascript:` reached `href`.
- **The write path emitted XML-illegal control characters ([#158](https://github.com/yfedoseev/office_oxide/issues/158)),** producing documents Word, Excel and LibreOffice all reject.
- **`replace_text` corrupted the document on `&` ([#159](https://github.com/yfedoseev/office_oxide/issues/159)),** never matched escaped text, and injected raw markup.
- **Panic on a malformed CFB header ([#138](https://github.com/yfedoseev/office_oxide/issues/138)).** An unvalidated sector shift overflowed in `cfb/header.rs`.
- **Element dispatch ignored XML namespaces ([#162](https://github.com/yfedoseev/office_oxide/issues/162)),** so foreign-namespace text was extracted as document content.
- **XLSX/PPTX writers validated nothing ([#173](https://github.com/yfedoseev/office_oxide/issues/173), [#160](https://github.com/yfedoseev/office_oxide/issues/160)).** An out-of-range sheet index silently dropped data, illegal sheet names and out-of-grid cells were accepted, and non-finite numbers were written as `<v>inf</v>`/`<v>NaN</v>` — a schema-invalid workbook, returned as `Ok`.

#### Renderers disagreeing with each other

- **DOCX had two markdown renderers that disagreed ([#137](https://github.com/yfedoseev/office_oxide/issues/137), [#166](https://github.com/yfedoseev/office_oxide/issues/166), [#136](https://github.com/yfedoseev/office_oxide/issues/136), [#134](https://github.com/yfedoseev/office_oxide/issues/134)),** and the IR path duplicated each section's first heading. Headers and footers appeared only in markdown, `<img>` was emitted without a source, and markdown `---` leaked into `plain_text`. `Heading::level`'s documented invariant is now enforced, and the `outline_level` doc comment no longer states its value space backwards.
- **Adjacent emphasis runs were not coalesced ([#153](https://github.com/yfedoseev/office_oxide/issues/153)),** producing literal `****` in the output, and `vertAlign` was lost.
- **PPT slide-master placeholder prompts were extracted as content ([#170](https://github.com/yfedoseev/office_oxide/issues/170)).** "Click to edit Master title style" is no longer document text.
- **IR round-trip was unbounded ([#171](https://github.com/yfedoseev/office_oxide/issues/171)).** Image alt text was written back as body text, growing the document on every generation.

#### Packaging

- **`go/cmd/install` 404'd on every run ([#110](https://github.com/yfedoseev/office_oxide/issues/110)).** The download URL put the version in the asset filename.

### Testing

- **Mechanical gates for parser correctness ([#133](https://github.com/yfedoseev/office_oxide/issues/133)).** SPRM opcode identity is now asserted against a table transcribed from [MS-DOC], so a wrong opcode fails a test rather than silently mis-parsing.
- **Release regression sweep (`scripts/regression-sweep/`).** Two arms — the previous tag and the release branch — run over a corpus of real Office documents, compared on two axes: status transitions and word-level content change. `vanished.py` is the check that decides content loss, because byte totals cannot: restoring a dropped `&` makes output longer and dropping 65,536 rows of grid padding makes it far shorter, and both swamp any aggregate. It found seven defects that 800+ synthetic tests could not see, including two introduced during this release.
- The suite grew from 792 to 818 tests, all fixtures built in code.

### Known gaps

- **[#119](https://github.com/yfedoseev/office_oxide/issues/119) is deliberately partial.** An encrypted OOXML file now reports that it is password-protected instead of failing as a corrupt zip. Reading and writing ECMA-376 agile encryption is not implemented: it needs four cryptography dependencies plus a CFB writer that does not exist here, and neither half can be verified in this repository without an Office-produced encrypted fixture. Shipping encryption that does not interoperate with Word would be the same confidently-wrong failure the rest of this release removes.
- **No external reference panel.** Every judgement in the regression sweep is 0.1.10 against v0.1.9, so a defect both versions share does not show up. Comparing against LibreOffice, Tika or pandoc remains future work.


## [0.1.9] - 2026-09-01

> Legacy binary `.doc` structure release: tables, lists, and tab stops are now recovered into the IR, closing the long-standing fidelity gap against the `.docx` path. Plus an IR list-nesting fix, green lint gates, and a dependency refresh. No breaking API changes.
>
> Both `.doc` features in this release were contributed by **[@xugangqiang](https://github.com/xugangqiang)** — thank you.

### Added

- **`.doc`: tables are extracted into the IR ([#115](https://github.com/yfedoseev/office_oxide/issues/115), [#116](https://github.com/yfedoseev/office_oxide/pull/116)).** For legacy binary Word files, `Document::to_ir()` (and therefore `to_ir_json()`, `to_markdown()`, and every binding built on them) previously emitted only `Paragraph` and heuristic `Heading` elements — table structure that exists in the file was silently dropped, while the `.docx` path emitted `Table` elements from the same API. Tables are now rebuilt by walking the structured paragraphs from the PAPX FKP of the table stream: a row ends at a `sprmTDefTable` row-terminator paragraph, a cell at a `\x07`-terminated paragraph. `col_span` is computed from the shared column-grid edges (the column-boundary union, Apache POI `buildTableCellEdgesArray` style) with a 4-twip snap tolerance, since Word rounds per row; `row_span` is driven by the **2-bit** `TCGRF.fVertMerge` field (`fvmClear` / `fvmMerge` / `fvmRestart`), including a guard for Word's whole-TAP-copy quirk where continuation rows wrongly carry `fvmRestart`. Nested tables (`itap > 1`) are flattened with a visible notice rather than silently mis-rendered. Carries shared parser hardening that underpins the walk: non-monotonic CP now returns `Err`, a `cw == 0` FKP re-read off-by-one is fixed, `fc_to_cp` is overflow-protected, and `decode_cp_range` is byte-clamped against a DoS. Every fixture is a minimal synthetic `.doc` built in code. Thanks @xugangqiang.
- **`.doc`: lists and tab stops are extracted into the IR ([#115](https://github.com/yfedoseev/office_oxide/issues/115), [#120](https://github.com/yfedoseev/office_oxide/pull/120)).** Building on the table walker. **Lists** key membership on the list format override id (`ilfo`, set by `sprmPIlfo` `0x460B`) per [MS-DOC] §2.4.6.3 rather than on the level SPRM alone, read as a signed `i16`: `0x0000` and `0xF801` mean "not in a list"; `0x0001`–`0x07FE` index into `PlfLfo.rgLfo`; `0xF802`–`0xFFFF` are the negation of a 1-based index and are still list items. Consecutive list paragraphs group into a nested `Element::List`, with nesting driven by `ilvl`. **Tab stops** are decoded from `sprmPChgTabs` (`0xC615`) and `sprmPChgTabsPapx` (`0xC60D`) — both carry a 1-byte `cb`, with `cb == 255` an escape for `0xC615` that derives the operand length from the payload's own `cDel`/`cAdd` counts, and the two opcodes differing only in their delete block (`PchgTabsDelClose`, `1 + 4·cDel`, vs `PchgTabsDel`, `1 + 2·cDel`) — and surfaced on the paragraph's `tabs`. The `0xXXXX` opcode identities and the `cb`/delete-block byte contracts are asserted directly, so the tests double as a spec-conformance gate. Two known gaps remain, documented rather than papered over: a negated `ilfo` is not yet resolved to a positive index, and a list whose `ilfo` is inherited only through a paragraph style is not yet detected. Thanks @xugangqiang.

### Fixed

- **IR: a list whose items all sit at the same non-zero depth no longer loses every item but the first ([#122](https://github.com/yfedoseev/office_oxide/pull/122)).** `ir::build_nested_list` decided each item's children by comparing the following items against the run's `base_level` instead of against that item's own level. When every item in a run was deeper than `base_level`, the first item claimed the whole run as its child range, discarded it (because the item itself was not shallow enough to be a parent), and still advanced the cursor past the range — silently dropping the rest. Both call sites pass a hard-coded `base_level` of 0, so **any list whose items all sit at level ≥ 1 lost all but its first item**: a PowerPoint placeholder whose bullets are all `<a:p lvl="1">` is an ordinary shape, and it measured 3 bullets in → 1 list item out. DOCX reached the same path through `<w:numPr w:ilvl>`. Grouping now uses each item's own depth and always advances to the end of the consumed range.

### CI / tooling (dev-only, not shipped)

- **Clippy: constant-size `chunks_exact` rewritten as `as_chunks` ([#124](https://github.com/yfedoseev/office_oxide/pull/124)).** Rust 1.98's clippy added `chunks_exact_to_as_chunks`; two call sites (`src/cfb/reader.rs`, `src/ppt/text.rs`) tripped it, and both the Clippy and Lint jobs run with `-D warnings`, so `main` itself failed. `as_chunks::<N>()` discards the remainder exactly as `chunks_exact` did, so behaviour is identical, and it hands `from_le_bytes` a `[u8; N]` directly. `as_chunks` stabilised in Rust 1.88, this workspace's declared `rust-version`, so the MSRV job is unaffected.
- **Python lint: 31 ruff findings cleared and the linter pinned to a minor series ([#123](https://github.com/yfedoseev/office_oxide/pull/123)).** `ruff>=0.4` was unpinned, so CI resolved whatever ruff published most recently and each release turned on new default rules — the gate changed under the repo with no commit, reddening every open PR including bumps that touch no Python at all. 29 of the findings were in the `_native.pyi` stub (PYI044/PYI026/PYI034/PYI036, PEP 604/585 unions — valid in a stub regardless of the `requires-python = ">=3.8"` runtime floor, since stubs are never evaluated at runtime), the other 2 import and `__all__` ordering. `ruff` is now pinned `>=0.16,<0.17`.
- **Dependency refresh**: `pyo3` 0.29.0 → 0.29.2, `thiserror` 2.0.19 → 2.0.20, `clap` 4.6.4 → 4.6.6, `log` 0.4.33 → 0.4.34, `fast-float2` 0.2.3 → 0.2.4, `koffi` 3.1.2 → 3.1.6 (JS binding), plus a bulk in-semver lockfile refresh ([#99](https://github.com/yfedoseev/office_oxide/pull/99)). Pinned GitHub Actions: `ossf/scorecard-action` 2.4.3 → 2.4.4, `pypa/gh-action-pypi-publish` 1.14.1 → 1.14.2.

## [0.1.8] - 2026-07-22

> PPT97 (legacy binary `.ppt`) text-extraction correctness release. No breaking API changes.

### Fixed

- **PPT: `PlainText()` / `ToMarkdown()` / `ToIRJSON()` no longer return stale or embedded-metadata text instead of real slide content on legacy binary `.ppt` files ([#85](https://github.com/yfedoseev/office_oxide/issues/85)).** PPT97 uses incremental ("fast") saves — each save appends new/changed records to the end of the "PowerPoint Document" stream, leaving superseded copies of earlier records (e.g. a pre-edit `Slide` container) behind in the stream with nothing to compact them away. The extractor scanned front-to-back and trusted whichever slide-shaped data it found first, so stale leftovers could win over current content, and in the worst case a corrupted/oversized declared record length (the record iterator bounded child data against the whole stream rather than its own enclosing container) could derail the scan and swallow unrelated later records. Slide content is now resolved through the PPT97 persist object directory instead — `Current User` stream → `UserEditAtom` chain (`offsetLastEdit`) → merged `PersistDirectoryAtom` entries, per [MS-PPT] 2.1.2 — so the *current* copy of each slide's records is used regardless of scan order, and the record iterator now bounds each record against its own container. Also resolves `OutlineTextRefAtom` indirection, where a placeholder shape stores text only as an index back into the slide's outline-text cache rather than embedding it directly — needed for correctness against a real-world corpus, not just the reported bug. Verified against 176 real-world `.ppt` files (LibreOffice QA data, Apache POI/Tika test resources, OSS-Fuzz regression cases): zero crashes either side vs. v0.1.7, and every file with changed output gained previously-missing content (stale duplicates collapsed, indexed placeholder text now resolved) with no garbled output observed. Thanks to the reporter of #85.

## [0.1.7] - 2026-07-19

> Round-trip and robustness release: the write↔parse IR cycle is now a fixed point for DOCX/PPTX, oversized XLSX sheets are bounded instead of stalling, plus the earlier XLSX cell-type/DOCX-truncation/XLSX-edit fixes, a dependency refresh, and CI-hardening + IP-governance groundwork. No breaking API changes.

### Added

- **IR spreadsheet cells now carry their semantic type and number format ([#72](https://github.com/yfedoseev/office_oxide/issues/72)).** Previously `Document::to_ir()` (and therefore the WASM `toIr()` surface) flattened every XLSX cell to a display string, so consumers could not tell a number or date cell from text — a `42` and a `"42"` looked identical, and a date serial under a date format was indistinguishable from a number. `TableCell` gains four optional, `serde`-defaulted fields populated on the XLSX grid path: `data_type` (`text` / `number` / `date` / `boolean` / `error`, via the new `CellDataType` enum), `raw_number` (the underlying `f64` — dates as an Excel serial, booleans as `1.0`/`0.0`), `number_format` (the format code, e.g. `"#,##0.00"`), and `number_format_id` (built-in or custom `numFmtId`, e.g. `14` for a date). The fields are `#[serde(skip_serializing_if = "Option::is_none")]`, so DOCX/PPTX table cells and existing JSON consumers are unaffected. Format metadata is surfaced on **every** non-empty cell — including text cells — so a caller can also tell that a text cell sits in a date-formatted column (a real case from the reporter's workbook, where a "date" column held typed-in text). Built-in format IDs (which never appear in `<numFmts>`) resolve to their canonical code via a new `numfmt::builtin_format_code`, so `number_format` is populated for built-ins like `14` → `"m/d/yyyy"` and `4` → `"#,##0.00"`, not just custom formats. Reuses the existing date-detection and number-format resolution logic rather than duplicating it. Thanks to the reporter for the clear request.

### Fixed

- **DOCX: `to_markdown()` / `docx_to_ir` no longer silently truncate the body on paragraphs with a bordered `<w:pBdr>` ([#71](https://github.com/yfedoseev/office_oxide/issues/71)).** A paragraph carrying Word's horizontal-rule idiom (an empty paragraph whose `pPr` holds a bottom-border-only `<w:pBdr>`) desynced the reader: after detecting the self-closing `<w:bottom/>` edge, the `pBdr` scanner issued a stray extra read that swallowed the `</w:pBdr>` close tag when no whitespace separated them (i.e. in real, non-pretty-printed Word output). With the border's end tag consumed and the loop keyed to break only on `</w:pBdr>`, the shared reader ran to EOF — dropping every subsequent paragraph and table and returning `Ok` with a partial document. The scanner now walks the `<w:pBdr>` subtree with a balanced depth counter and exits at the matching close, so trailing content is preserved. Thanks to the reporter for the precise measurements.
- **XLSX edit: `set_cell` no longer strips a cell's attributes, and no longer deletes the next cell when writing into an empty-but-styled one.** `set_cell_in_xml` rebuilt the `<c>` element from scratch — emitting only `r` and `t` — so every other attribute of the overwritten cell was discarded, most importantly `s`, the style index: a currency or date cell came back with no number format. Worse, the replacement located the end of the cell by searching for `</c>` *before* checking whether the opening tag was self-closing; an empty-but-styled cell is written `<c r="A1" s="5"/>` with no `</c>`, so the search ran on and found the **next** cell's closing tag, and the splice swallowed the cell in between. Both defects converge on the same real workflow — injecting values into a pre-formatted spreadsheet template, whose data rows are made entirely of styled, empty cells: the write silently destroyed a neighbouring cell and stripped the formatting off the one it wrote. `set_cell` now carries every attribute of the existing cell through verbatim (except `r` and `t`, which it owns) and delimits the opening tag before deciding how the element ends.
- **DOCX/PPTX: `create_from_ir` (and `create_from_markdown`) round-trips are now idempotent ([#79](https://github.com/yfedoseev/office_oxide/issues/79)).** Writing an IR and reading it back drifted on every cycle: a document/slide title is surfaced by the parsers as **both** `section.title` **and** a leading heading, and the writer re-emitted the title on top of that heading. For DOCX this materialised a duplicate H1 that grew monotonically each write→parse cycle; for PPTX the title was re-emitted as body text and — because `emit_pptx_element` had **no arm for `Element::TextBox`**, the wrapper the parser puts slide-body content in — the real body (bullets, paragraphs) was silently dropped, leaving only the title echo. The writer now skips a standalone title when it is already the section's leading heading (both formats) and emits `Element::TextBox` content into the slide body, so `parse(write(x))` is a fixed point. Covered by a new `ir_roundtrip_idempotence` test for all three OOXML formats.

### Changed

- **XLSX: worksheets are capped at 10,000 rows when materialised into the IR ([#29](https://github.com/yfedoseev/office_oxide/issues/29)).** A worksheet is converted eagerly into a single in-memory `Table`, which downstream cannot render to an unbounded PDF, so a pathological sheet (e.g. a 200k-row fixture) exploded memory and stalled rendering. Beyond the cap the remaining rows are dropped and a **visible truncation notice** (`[N of M rows not shown — worksheet truncated at 10000 rows]`) is appended, so the omission is obvious in text/markdown/PDF rather than silent or hanging. Ordinary large spreadsheets are unaffected.
- **Refreshed `Cargo.lock` to latest semver-compatible versions** (in-semver, no manifest changes): `clap` 4.6.1 → 4.6.2, `futures-*` 0.3.32 → 0.3.33, `portable-atomic` 1.13.1 → 1.14.0, `simd-adler32` 0.3.9 → 0.3.10, `thiserror` 2.0.18 → 2.0.19, `syn` 2.0.118 → 2.0.119, and others. `thiserror-impl` 2.0.19 now depends on `syn` 3.0, so the tree carries a second `syn` major transitively; `cargo deny` is clean.

### CI / tooling & governance (dev-only, except NOTICE/TRADEMARKS which ship in the crate)

- **CI hardening.** Least-privilege default `GITHUB_TOKEN` permissions on the Python workflow (OpenSSF Scorecard Token-Permissions); a `cargo-fuzz` target (`fuzz/`) that feeds arbitrary bytes to every format parser (must never panic/hang); deterministic PR-quality gates (Conventional-Commits title, no committed compiled/binary artifacts, AI-assistance labelling); and template-compliance + CLA-assistant workflows. The `pull_request_target` workflows only read event/API strings and never check out or run PR code.
- **IP governance.** Added [`CLA.md`](CLA.md) (contributor licence agreement — a licence, not an assignment), [`TRADEMARKS.md`](TRADEMARKS.md), [`NOTICE`](NOTICE), and [`AGENTS.md`](AGENTS.md); `NOTICE` and `TRADEMARKS.md` are now included in the published crate package. The code remains dual-licensed `MIT OR Apache-2.0`.

## [0.1.6] - 2026-07-13

> Fixes PPTX slides created by `PptxWriter` / `create_from_markdown` rendering blank in LibreOffice Impress.

### Fixed

- **PPTX placeholder shapes now carry explicit geometry ([#69](https://github.com/yfedoseev/office_oxide/issues/69)).** The generated title and body placeholders previously emitted an empty `<p:spPr/>` with no `<a:xfrm>`, relying on the slide layout for position/size. PowerPoint resolves that from the layout, but LibreOffice Impress does not — so slides rendered blank. The writer now emits an explicit `<a:xfrm>` (offset + extent) on each placeholder, derived from PowerPoint's default Office-theme rectangles and scaled to the presentation size (`set_presentation_size`). Thanks @holgerreppert for the detailed report.

## [0.1.5] - 2026-07-13

> Dependency-maintenance release. No API or behavior changes — all 424 tests pass and every feature build (default, `parallel`, `mmap`, `python`, `wasm`) compiles clean.

### Changed

- **Refreshed all Rust dependencies to their latest semver-compatible versions** (lockfile): `clap` 4.5 → 4.6.1, `wasm-bindgen` 0.2.114 → 0.2.126, `js-sys` 0.3.91 → 0.3.103, `libc` 0.2.182 → 0.2.186, `serde_json` 1.0.149 → 1.0.150, `crossbeam-deque` 0.8.6 → 0.8.7, `crossbeam-utils` 0.8.21 → 0.8.22, plus `quote`/`syn`/`itoa`/`either` and others. The crate's version requirements were already broad enough to permit these — intentionally kept loose so downstreams aren't over-constrained.
- **JS binding: `koffi` 2.9 → 3.1.1** (major). Uses only stable core APIs (`load`, `decode`, `disposable`, `func`); koffi 3.x ships native binaries as platform-specific optional packages. Node engine requirement (`>=18`) unchanged.

### CI / tooling (dev-only, not shipped)

- **Pinned GitHub Actions bumped to latest (SHA-pinned):** `checkout` v4 → v7.0.0, `codecov-action` v5 → v7.0.0, `codeql-action` v3 → v4.37.0, `attest-build-provenance` v2 → v4.1.1, `setup-go` v5 → v6.5.0, `setup-dotnet` v4 → v5.4.0, `setup-python` v6.2 → v6.3.0, `setup-uv` v8.1 → v8.3.2, `cargo-deny-action` → v2.0.20, `scorecard-action` → v2.4.3, `action-gh-release` v2 → v3.0.1, `rust-cache` → v2.9.1, `install-action` → v2.83.2, and `actions/cache` moved from a `@v5` tag to a SHA-pinned v6.1.0. Also unified a stray `download-artifact@v4` to v8.0.1 and corrected a stale `github-script` version comment. Subsumes Dependabot #40/#41/#43/#44/#45.
- **pre-commit hooks:** `cargo-deny` 0.19.0 → 0.20.2, `ruff-pre-commit` v0.15.5 → v0.15.21, `ty-pre-commit` v0.0.18 → v0.0.59.
- **C# test deps (test-only):** `Microsoft.NET.Test.Sdk` 17.10 → 18.7.0, `xunit` 2.9.0 → 2.9.3, `xunit.runner.visualstudio` 2.8.2 → 3.1.5.

### Notes

- No new advisories (`cargo audit` / `cargo deny` clean, 81 deps).
- Benchmark competitor pins in `scripts/bench-requirements.txt` are intentionally frozen until the next full benchmark re-run (bumping them would desync the numbers in BENCHMARKS.md), so Dependabot #53/#54 are out of scope here.
- `download-artifact` v4 → v8 and `attest-build-provenance` v2 → v4 run in the release workflow, which only fully executes on a tag; they are validated when v0.1.5 is tagged.

## [0.1.4] - 2026-07-13

> Build-compatibility fix that unblocks downstream crates from moving to quick-xml ≥ 0.41 (and clearing RUSTSEC-2026-0194 / RUSTSEC-2026-0195) when office_oxide shares a dependency tree with a crate that enables quick-xml's `encoding` feature.

### Fixed

- **quick-xml `encoding`-feature build break** — `unescape_attr_value` called `Attribute::unescape_value()`, which quick-xml gates behind `#[cfg(not(feature = "encoding"))]`. Cargo feature unification turns `encoding` on transitively whenever another crate in the tree enables it (e.g. `calamine ≥ 0.36`), so office_oxide failed to compile with `error[E0599]: no method named unescape_value`. This forced downstreams (e.g. `kreuzberg`) to pin `calamine 0.35`, keeping a vulnerable quick-xml < 0.41 in the tree. The helper now decodes the raw attribute bytes as UTF-8 and expands entities via `escape::unescape`, which is feature-independent. Thanks @waxzce (#63) and @Goldziher (#64).

### Changed

- **`zip` 8.1 → 8.6** and assorted transitive dependency refreshes. Via #64 (@Goldziher).

## [0.1.3] - 2026-07-04

> Security patch: clears every open RustSec advisory in the dependency tree — an untrusted-XML DoS (quick-xml), two pyo3 soundness/OOB fixes, and a memmap2 unsound-pointer fix — plus Linux aarch64 wheels on PyPI.

### Security

- **quick-xml 0.40 → 0.41 (RUSTSEC-2026-0194 / RUSTSEC-2026-0195)** — quick-xml 0.40's `NamespaceResolver::push` performs an unbounded per-`xmlns` heap allocation inside `NsReader` *before* the parse event is returned, so a crafted Office document (DOCX/XLSX/PPTX) declaring many namespaces could exhaust memory and crash the reader — a denial-of-service on untrusted input. Upgraded to quick-xml 0.41, which bounds the allocation. Thanks @smissingham (#58); tracked in #59.
- **pyo3 0.28 → 0.29 (RUSTSEC-2026-0176 / RUSTSEC-2026-0177)** — fixes an out-of-bounds read in `nth` / `nth_back` for `PyList` and `PyTuple` iterators, and a missing `Sync` bound on `PyCFunction::new_closure` closures, in the Python binding. Via Dependabot (#52).
- **memmap2 0.9.10 → 0.9.11 (RUSTSEC-2026-0186)** — fixes an unchecked pointer offset (unsoundness). Via Dependabot (#55).

### Packaging

- **Linux aarch64 wheels + source distribution on PyPI (#51)** — `pip install office-oxide` on Linux/arm64 (AWS Graviton, Apple-Silicon Docker, Raspberry Pi) failed with "no compatible wheel"; the release workflow now builds `manylinux` aarch64 wheels and an sdist alongside the existing x86_64/macOS wheels. Thanks @regularkevvv.

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
