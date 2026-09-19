use crate::doc::{
    ChpProps, DocDocument, DocParagraph, HyperlinkSpan, ListFormatting, PapProps, TapCellInfo,
    TapInfo,
};
use crate::format::DocumentFormat;
use crate::ir::*;

/// Convert a parsed legacy `.doc` into the intermediate representation.
///
/// When the FIB advertised a PlcfBtePapx, `DocDocument` carries structured
/// paragraphs with PAP flags and we walk them in order to rebuild tables
/// (and, eventually, lists). Otherwise we fall back to a line-based
/// heuristic over the sanitised text — the original `.doc` behaviour.
pub(crate) fn doc_to_ir(doc: &DocDocument) -> DocumentIR {
    let mut elements: Vec<Element> = Vec::new();

    let paragraphs = doc.paragraphs();
    // One whole-document decision, taken before the walk, gated on the
    // *outcome*: run the line-shape heading guess only when no paragraph
    // resolved a real outline level.
    //
    // Gating on "the document has a stylesheet" instead would be wrong —
    // a parsed style sheet is not evidence a document uses headings, and
    // letters, memos and forms with a valid STSH containing no heading
    // styles would silently lose their headings, and with them
    // `metadata.title` and `Section.title`, both of which are derived from
    // the first `Element::Heading` below.
    let has_structured_headings = paragraphs.iter().any(|p| p.props.outline_level.is_some());
    if !paragraphs.is_empty() {
        walk_paragraphs(paragraphs, has_structured_headings, &mut elements, doc.list_formatting());
    } else {
        line_heuristic(doc.plain_text_ref(), &mut elements);
    }

    let heading_title = elements.iter().find_map(|e| match e {
        Element::Heading(h) => h.content.first().and_then(|c| match c {
            InlineContent::Text(t) => Some(t.text.clone()),
            _ => None,
        }),
        _ => None,
    });
    // The file's own declared title (from `\x05SummaryInformation`) beats
    // a line-shape guess whenever both exist — a document can style
    // *some* headings and still use plain ALL-CAPS lines for others
    // (#224), but the declared title is never a guess (issue #244).
    let summary = doc.summary_properties();
    let title = summary
        .and_then(|s| s.title.clone())
        .filter(|t| !t.is_empty())
        .or(heading_title);

    let mut sections = vec![Section {
        title: title.clone(),
        elements,
        ..Default::default()
    }];

    // The header document's PlcfHdd-delimited stories for the first
    // section — this crate models only one `ir::Section` per DOC
    // document, so a document with 2+ sections only surfaces the first
    // one's headers/footers here (issue #285). When PlcfHdd yielded
    // nothing (older/malformed files), every field below is `None` and
    // the generic-TextBox fallback in the loop below still applies.
    let hf = doc.header_footer();
    if let Some(section) = sections.last_mut() {
        section.even_page_header = hf.even_header.as_deref().map(text_to_header_footer);
        section.header = hf.odd_header.as_deref().map(text_to_header_footer);
        section.even_page_footer = hf.even_footer.as_deref().map(text_to_header_footer);
        section.footer = hf.odd_footer.as_deref().map(text_to_header_footer);
        section.first_page_header = hf.first_header.as_deref().map(text_to_header_footer);
        section.first_page_footer = hf.first_footer.as_deref().map(text_to_header_footer);
    }
    let header_footer_structured = !hf.is_empty();

    // Footnotes, headers, comments, endnotes and text boxes live after the
    // main text in the same character space. Their `ccp*` lengths were
    // parsed and never used, so none of this reached a consumer.
    if let Some(section) = sections.last_mut() {
        let mut next_id = 0u32;
        for sub in doc.subdocuments() {
            // Already represented structurally on `Section.header`/
            // `.footer`/etc above — don't also dump the merged blob as a
            // generic TextBox (issue #285).
            if sub.kind == crate::doc::SubDocumentKind::HeadersFooters && header_footer_structured
            {
                continue;
            }
            // `PlcfandTxt`/`PlcfandRef` successfully split this document's
            // Comments substory into individual, correctly-attributed
            // comments — emit one `Element::Endnote` per comment instead
            // of falling through to the generic merged-substory path
            // below (issue #345).
            if sub.kind == crate::doc::SubDocumentKind::Comments && !doc.comments().is_empty() {
                for comment in doc.comments() {
                    let content: Vec<Element> = comment
                        .text
                        .lines()
                        .filter(|l| !l.trim().is_empty())
                        .map(|l| {
                            Element::Paragraph(Paragraph {
                                content: vec![InlineContent::Text(TextSpan::plain(l))],
                                ..Default::default()
                            })
                        })
                        .collect();
                    if content.is_empty() {
                        continue;
                    }
                    section.elements.push(Element::Endnote(Note {
                        id: next_id,
                        marker: Some(subdocument_label(sub.kind).to_string()),
                        content,
                        author: comment.author.clone(),
                    }));
                    next_id += 1;
                }
                continue;
            }
            // Footnote/endnote bodies are self-delimited: each one starts
            // with the literal auto-number reference-mark character
            // (`\u{2}`) in the substory's own text — confirmed on the full
            // local corpus (footnotes 49/51 files, endnotes 5/5 files that
            // had one). Comments carry no such marker in their substory
            // (0/12 files); they're split above via `PlcfandTxt` instead
            // when that PLC parses cleanly (issue #345). This is the
            // fallback path for a Comments substory whose `doc.comments()`
            // came back empty (PLC absent/malformed/mismatched) — it stays
            // merged into one Note, same as before #345.
            let splittable = matches!(
                sub.kind,
                crate::doc::SubDocumentKind::Footnotes | crate::doc::SubDocumentKind::Endnotes
            ) && sub.text.contains('\u{2}');

            let bodies: Vec<&str> = if splittable {
                sub.text.split('\u{2}').map(str::trim).filter(|s| !s.is_empty()).collect()
            } else {
                vec![sub.text.as_str()]
            };

            for body in bodies {
                let content: Vec<Element> = body
                    .lines()
                    .filter(|l| !l.trim().is_empty())
                    .map(|l| {
                        Element::Paragraph(Paragraph {
                            content: vec![InlineContent::Text(TextSpan::plain(l))],
                            ..Default::default()
                        })
                    })
                    .collect();
                if content.is_empty() {
                    continue;
                }
                // A single name in GrpXstAtnOwners unambiguously authored
                // every comment in the document — real-world common case.
                // Multiple names would need per-comment PlcfAtn/ATRD
                // correlation (not yet implemented) to attribute
                // correctly, so leave it unset rather than guess (issue
                // #298).
                let author = match sub.kind {
                    crate::doc::SubDocumentKind::Comments => match doc.comment_authors() {
                        [single] => Some(single.clone()),
                        _ => None,
                    },
                    _ => None,
                };
                let note = Note {
                    id: next_id,
                    marker: Some(subdocument_label(sub.kind).to_string()),
                    content,
                    author,
                };
                next_id += 1;
                section.elements.push(match sub.kind {
                    crate::doc::SubDocumentKind::Footnotes => Element::Footnote(note),
                    crate::doc::SubDocumentKind::HeadersFooters
                    | crate::doc::SubDocumentKind::TextBoxes
                    | crate::doc::SubDocumentKind::HeaderTextBoxes => Element::TextBox(TextBox {
                        content: note.content,
                        ..Default::default()
                    }),
                    _ => Element::Endnote(note),
                });
            }
        }
    }
    // Extracted pictures never reached the IR, so every image in a legacy
    // document was silently dropped on conversion even though the bytes
    // were already in hand.
    crate::convert_xls::append_legacy_images(&mut sections, doc.images());

    // Embedded OLE objects (issue #284) — at minimum, recognize each
    // object exists and surface its identity, even without extracting
    // its native payload. Mirrors #337's identical fix for legacy PPT:
    // a data-less Element::Image with a descriptive alt_text, the same
    // precedent #300 established for a data-less AutoShape placeholder.
    if !doc.ole_objects().is_empty() {
        if sections.is_empty() {
            sections.push(Section::default());
        }
        let last = sections.last_mut().expect("just ensured non-empty");
        for obj in doc.ole_objects() {
            last.elements.push(Element::Image(Image {
                alt_text: Some(obj.description.clone()),
                data: None,
                ..Default::default()
            }));
        }
    }

    DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Doc,
            title,
            author: summary.and_then(|s| s.author.clone()).filter(|s| !s.is_empty()),
            subject: summary.and_then(|s| s.subject.clone()).filter(|s| !s.is_empty()),
            keywords: summary
                .and_then(|s| s.keywords.as_deref())
                .map(crate::convert_docx::split_keywords)
                .unwrap_or_default(),
            description: summary.and_then(|s| s.comments.clone()).filter(|s| !s.is_empty()),
            created: summary.and_then(|s| s.created.clone()),
            modified: summary.and_then(|s| s.modified.clone()),
            has_macros: doc.has_macros(),
            text_truncated: !doc.text_complete(),
            ..Default::default()
        },
        sections,
        defined_names: Vec::new(),
    }
}

// ---------------------------------------------------------------------------
// Structured-paragraph walk (tables)
// ---------------------------------------------------------------------------
// Known limitation: the row-mark `sprmTDefTable` operand carries `itap` (the
// nesting depth), but this walk builds a single flat table run. Nested tables
// are therefore flattened into their containing table rather than represented
// as nested `TableRow`/`TableCell` blocks. This is a documented gap, not a
// deliberate silent merge of cells; merged-cell spans are still computed from
// `rgdxaCenter` and the vertical-merge flags as described below.

/// A completed table row awaiting span resolution: its cells plus the row
/// definition (TAP) from the row-terminator paragraph, when one was present.
struct PendingRow {
    /// Cells of the row; each cell is its block elements.
    cells: Vec<Vec<Element>>,
    /// Parsed `sprmTDefTable` of the row mark. `None` for rows closed without
    /// an explicit terminator, which then get `col_span = row_span = 1`.
    tap: Option<TapInfo>,
    /// Table nesting depth from `sprmPItap` (1 = top-level). `> 1` means a
    /// nested table, which this walk flattens.
    itap: u8,
}

/// Accumulator for the table currently being built from a run of `fInTable`
/// paragraphs. A row ends at a table-trailing-mark paragraph (one carrying
/// `sprmTDefTable`); a cell ends at a `\x07`-terminated paragraph within a row.
struct TableBuilder {
    /// Whether we are currently inside a table run.
    in_table: bool,
    /// Completed rows of the table under construction.
    rows: Vec<PendingRow>,
    /// Cells of the row under construction; each cell is its block elements.
    row_cells: Vec<Vec<Element>>,
    /// Block elements of the cell under construction (for multi-paragraph
    /// cells where interior paragraphs are `\r`-terminated).
    cell: Vec<Element>,
}

impl TableBuilder {
    fn new() -> Self {
        Self {
            in_table: false,
            rows: Vec::new(),
            row_cells: Vec::new(),
            cell: Vec::new(),
        }
    }

    /// Begin a new table run if not already inside one.
    fn ensure_open(&mut self) {
        if !self.in_table {
            self.in_table = true;
            self.rows.clear();
            self.row_cells.clear();
            self.cell.clear();
        }
    }

    /// Add a cell paragraph (`fInTable`, not a row mark).
    ///
    /// A `\x07` terminator closes the current cell; any other terminator
    /// (e.g. `\r`) is an interior paragraph break within the same cell.
    fn add_cell_paragraph(&mut self, p: &DocParagraph) {
        self.ensure_open();
        if !p.text.is_empty() {
            self.cell.push(Element::Paragraph(Paragraph {
                content: inline_content_for(&p.text, &p.hyperlinks, &p.chp_runs),
                // `tabs` is always empty in the tables-only build (it is
                // populated by the list/tab-stop PR); cloning keeps the IR
                // shape uniform with the list path.
                tabs: p.props.tabs.clone(),
                alignment: p.props.alignment.clone(),
                indent_left_twips: p.props.indent_left_twips,
                indent_right_twips: p.props.indent_right_twips,
                first_line_indent_twips: p.props.first_line_indent_twips,
                space_before_twips: p.props.space_before_twips,
                space_after_twips: p.props.space_after_twips,
                ..Default::default()
            }));
        }
        if p.terminator == '\u{7}' {
            self.row_cells.push(std::mem::take(&mut self.cell));
        }
    }

    /// A row-terminator paragraph: close the in-flight cell, then the row.
    fn end_row(&mut self, tap: Option<TapInfo>, itap: u8) {
        self.ensure_open();
        if !self.cell.is_empty() {
            self.row_cells.push(std::mem::take(&mut self.cell));
        }
        let cells = std::mem::take(&mut self.row_cells);
        self.rows.push(PendingRow { cells, tap, itap });
    }

    /// Flush any open table as an `Element::Table`. Called when prose (a
    /// non-`fInTable` paragraph) interrupts a table run, and at end-of-doc.
    fn flush(&mut self, elements: &mut Vec<Element>) {
        if !self.in_table {
            return;
        }
        // A row left without an explicit terminator — keep it, without a TAP.
        if !self.row_cells.is_empty() {
            let cells = std::mem::take(&mut self.row_cells);
            self.rows.push(PendingRow {
                cells,
                tap: None,
                itap: 0,
            });
        }
        if !self.rows.is_empty() {
            // Nested tables (itap > 1) are not yet represented as nested
            // `Table` blocks; the rows are flattened into the outer grid.
            // Surface that as a visible notice rather than silently emitting a
            // wrong structure (robustness contract: degrade gracefully).
            if self.rows.iter().any(|r| r.itap > 1) {
                elements.push(Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain(
                        "[nested table detected — not yet supported, flattened into the \
                         outer table]",
                    ))],
                    ..Default::default()
                }));
            }
            let rows = build_table_rows(&self.rows);
            elements.push(Element::Table(Table {
                rows,
                ..Default::default()
            }));
        }
        self.in_table = false;
    }
}

/// The vertical-merge state lives in a 2-bit field of the `TCGRF` `rgf`
/// (MS-DOC §2.9.185 `TCGRF`): bits 5-6, `fVertMerge`. The three legal values
/// are `fvmClear = 0` (not merged), `fvmMerge = 1` (continuation of the merge
/// started above), and `fvmRestart = 3` (first cell of a merge). As raw `rgf`
/// bits that is `fvmMerge = 0x0020` and `fvmRestart = 0x0060` (both bits set).
const FVM_CLEAR: u8 = 0;
const FVM_MERGE: u8 = 1; // rgf 0x0020
const FVM_RESTART: u8 = 3; // rgf 0x0060

/// Extract the 2-bit `fVertMerge` state from `TCGRF.rgf` (bits 5-6).
fn vert_merge_state(rgf: u16) -> u8 {
    ((rgf >> 5) & 0x03) as u8
}
/// When unifying column-grid edges across rows, boundaries within this many
/// twips are snapped together. Word rounds boundary positions per row, so
/// otherwise-near-identical edges inject a spurious grid edge and inflate
/// `col_span` for every cell spanning it (LibreOffice uses the same
/// `nTolerance = 4` in `FindMergeGroup`).
const EDGE_TOLERANCE_TWIPS: i32 = 4;

/// Resolve merged-cell spans for the accumulated rows.
///
/// The column grid is the sorted union of every row's `rgdxaCenter`
/// boundaries — the same array Apache POI's `buildTableCellEdgesArray`
/// computes. A cell's `col_span` is the number of grid edges inside its own
/// boundary interval `[centers[i], centers[i+1])`, which matches POI's
/// `getNumberColumnsSpanned`. `row_span` walks `fVertMerge`/`fVertRestart`
/// down the column; continuation cells are absorbed into the cell above and
/// not emitted (the same convention as the `.docx` converter).
fn build_table_rows(pending: &[PendingRow]) -> Vec<TableRow> {
    let mut edges: Vec<i32> = pending
        .iter()
        .filter_map(|r| r.tap.as_ref())
        .flat_map(|tap| tap.centers.iter().map(|&c| c as i32))
        .collect();
    edges.sort_unstable();
    // Dedup with tolerance: keep an edge only when it is more than
    // `EDGE_TOLERANCE_TWIPS` above the previous kept edge.
    let mut grid: Vec<i32> = Vec::with_capacity(edges.len());
    for e in edges {
        if grid
            .last()
            .is_none_or(|&last| e - last > EDGE_TOLERANCE_TWIPS)
        {
            grid.push(e);
        }
    }

    pending
        .iter()
        .enumerate()
        .map(|(row_idx, prow)| {
            let cells = match &prow.tap {
                Some(tap)
                    if tap.centers.len() == tap.cells.len() + 1
                        && tap.cells.len() == prow.cells.len() =>
                {
                    let (absorbed, spans) = resolve_row_spans(row_idx, pending, tap);
                    (0..tap.cells.len())
                        .filter_map(|col| {
                            if absorbed[col] {
                                return None; // vertical-merge continuation
                            }
                            Some(TableCell {
                                content: prow.cells[col].clone(),
                                col_span: count_grid_edges(&tap.centers, col, &grid),
                                row_span: spans[col],
                                ..Default::default()
                            })
                        })
                        .collect()
                },
                // No TAP or a cell-count mismatch — plain 1×1 cells.
                _ => prow
                    .cells
                    .iter()
                    .cloned()
                    .map(|content| TableCell {
                        content,
                        col_span: 1,
                        row_span: 1,
                        ..Default::default()
                    })
                    .collect(),
            };
            TableRow {
                cells,
                ..Default::default()
            }
        })
        .collect()
}

/// Per-column `(is_absorbed, row_span)` for one table row.
///
/// The 2-bit `fVertMerge` field (MS-DOC `TCGRF`) drives the logic:
/// `fvmRestart` starts a new merge, `fvmMerge` continues the merge above, and
/// `fvmClear` is an ordinary cell. A `fvmMerge` is absorbed into the cell above
/// and not emitted; the restart cell that begins each merge carries the
/// `row_span`. Every `fvmRestart` is a genuine new merge — it is never treated
/// as a continuation, so a `fvmRestart` that immediately follows another
/// `fvmRestart` opens its own separate merge. The only other representable
/// 2-bit value (the reserved `0x0040`) is treated as `fvmClear` rather than
/// panicking, keeping decoding robust against malformed input.
fn resolve_row_spans(
    row_idx: usize,
    pending: &[PendingRow],
    tap: &TapInfo,
) -> (Vec<bool>, Vec<u32>) {
    let mut absorbed = vec![false; tap.cells.len()];
    let mut spans = vec![1u32; tap.cells.len()];

    for (col, tc) in tap.cells.iter().enumerate() {
        match vert_merge_state(tc.rgf) {
            FVM_CLEAR => continue,             // ordinary cell, span 1, not absorbed
            FVM_MERGE => absorbed[col] = true, // continuation of the merge above
            FVM_RESTART => {
                // Genuine restart: count consecutive continuation cells below.
                // A `fvmRestart` is never a continuation, so an immediately
                // following `fvmRestart` opens a distinct merge (span 1 here).
                let mut span = 1u32;
                for rr in row_idx + 1..pending.len() {
                    if is_merge_continuation(pending, rr, col) {
                        span += 1;
                    } else {
                        break;
                    }
                }
                spans[col] = span;
            },
            // The reserved 2-bit value (0x0040) is malformed but representable;
            // treat it as an ordinary cell rather than panicking.
            _ => continue,
        }
    }
    (absorbed, spans)
}

/// The 2-bit `fVertMerge` state of the cell at `(row, col)`, or `FVM_CLEAR`
/// when the row has no TAP or no such column.
fn cell_state(pending: &[PendingRow], row: usize, col: usize) -> u8 {
    pending[row]
        .tap
        .as_ref()
        .and_then(|tap| tap.cells.get(col))
        .map(|tc: &TapCellInfo| vert_merge_state(tc.rgf))
        .unwrap_or(FVM_CLEAR)
}

/// Whether the cell at `(row, col)` continues the vertical merge above it.
///
/// Only a `fvmMerge` cell continues. A `fvmRestart` cell always opens a new
/// merge, never a continuation, so it stops the span begun by the cell above.
fn is_merge_continuation(pending: &[PendingRow], row: usize, col: usize) -> bool {
    cell_state(pending, row, col) == FVM_MERGE
}

/// Number of column-grid edges inside the cell's boundary interval
/// `[centers[col], centers[col+1])`. At least 1, since `centers[col]`
/// itself is always a grid edge.
fn count_grid_edges(centers: &[i16], col: usize, grid: &[i32]) -> u32 {
    let lo = centers[col] as i32;
    let hi = centers[col + 1] as i32;
    grid.iter().filter(|&&e| e >= lo && e < hi).count().max(1) as u32
}

/// Walk structured paragraphs in document order, emitting tables, lists,
/// and prose.
///
/// List handling is a first cut: consecutive paragraphs that carry a list
/// level SPRM (`0x460B` / `ilvl`) are grouped into one `Element::List`,
/// with nesting driven by `ilvl`. The ordered-vs-bullet distinction and
/// list-id (`ilfo`) grouping require the style table + PlcfLst/LSTF/LVL
/// chain, which is out of scope here; lists therefore default to bullet
/// (unordered). Lists whose `ilfo` is inherited only via a paragraph style
/// (no direct SPRM) are not yet detected and fall through as prose — a
/// graceful no-op rather than a regression.
///
/// `.doc` list levels are not guaranteed to start at 0 (Word writes the
/// level as stored in the list definition, which may begin at 1), so each
/// run's base level is taken as the minimum `ilvl` in that run — see
/// `flush_list`.
///
/// A paragraph is a list item iff its `ilfo` (sprmPIlfo, `0x460B`) is a valid
/// list index, per [MS-DOC] §2.4.6.3 ("If iLfoCur is zero, the paragraph is not
/// part of a list"). The operand is decoded as a signed `i16`:
/// `0x0000` / `0xF801` mean "not in a list"; `0x0001`–`0x07FE` are 1-based
/// indices into `PlfLfo.rgLfo`; `0xF802`–`0xFFFF` are the negation of a 1-based
/// index and are still list items. `None` (no sprmPIlfo) defaults to prose.
fn is_doc_list_item(ilfo: Option<i16>) -> bool {
    match ilfo {
        None | Some(0) | Some(-2047) => false, // 0x0000 / 0xF801: not in a list
        // TODO(ilfo-negated): 0xF802..=0xFFFF (i16 -2046..=-1) are list items
        // whose `ilfo` is the negation of a 1-based index; resolve to the
        // positive index when list-id grouping is implemented. Until then they
        // must still be emitted as list items, not dropped to prose.
        Some(v) if (1..=0x07FE).contains(&v) => true, // 0x0001..0x07FE normal
        Some(v) if (-0x07FE..=-1).contains(&v) => true, // 0xF802..0xFFFF negated
        _ => false,                                   // 0x07FF and other non-spec
    }
}

fn walk_paragraphs(
    paragraphs: &[DocParagraph],
    has_structured_headings: bool,
    elements: &mut Vec<Element>,
    list_formatting: &ListFormatting,
) {
    let mut table = TableBuilder::new();
    let mut list_items: Vec<(u8, Vec<InlineContent>)> = Vec::new();
    // The run's own `ilfo` (issue #250) — every item in one contiguous list
    // run shares the same `ilfo`/`ilvl`-derived list identity in practice
    // (a change of `ilfo` mid-run would itself interrupt list-item
    // membership via `is_doc_list_item`), so the first item's value is
    // enough to resolve `start_number`/`ordered` for the whole run.
    let mut list_ilfo: Option<i16> = None;

    for p in paragraphs {
        if p.props.is_table_trailing_mark {
            flush_list(&mut list_items, list_ilfo.take(), list_formatting, elements);
            table.end_row(p.props.tap.clone(), p.props.itap);
        } else if p.props.f_in_table {
            flush_list(&mut list_items, list_ilfo.take(), list_formatting, elements);
            table.add_cell_paragraph(p);
        } else if let Some(lvl) = p.props.outline_level {
            // Outline level wins over list membership, exactly like the
            // DOCX converter: Word's multilevel-list "Heading" gallery
            // attaches an ilfo to the heading styles themselves, so a
            // numbered heading ("1. Introduction") is the normal shape of
            // a heading in real documents. Checking ilfo first turned
            // every one of them into a list item and left no Headings in
            // the IR at all (issue #223).
            table.flush(elements);
            flush_list(&mut list_items, list_ilfo.take(), list_formatting, elements);
            emit_heading(&p.text, lvl + 1, elements, &p.hyperlinks, &p.chp_runs);
        } else if is_doc_list_item(p.props.ilfo) {
            // List membership is keyed on `ilfo` (sprmPIlfo, `0x460B`), not on
            // `ilvl`: per [MS-DOC] §2.4.6.3 a paragraph is a list item only when
            // its `ilfo` is a valid list index. `ilvl` still drives nesting.
            table.flush(elements);
            let ilvl = p.props.ilvl.unwrap_or(0);
            if list_ilfo.is_none() {
                list_ilfo = p.props.ilfo;
            }
            list_items.push((ilvl, inline_content_for(&p.text, &p.hyperlinks, &p.chp_runs)));
        } else {
            table.flush(elements);
            flush_list(&mut list_items, list_ilfo.take(), list_formatting, elements);
            if has_structured_headings {
                elements.push(Element::Paragraph(Paragraph {
                    content: inline_content_for(&p.text, &p.hyperlinks, &p.chp_runs),
                    tabs: p.props.tabs.clone(),
                    alignment: p.props.alignment.clone(),
                    indent_left_twips: p.props.indent_left_twips,
                    indent_right_twips: p.props.indent_right_twips,
                    first_line_indent_twips: p.props.first_line_indent_twips,
                    space_before_twips: p.props.space_before_twips,
                    space_after_twips: p.props.space_after_twips,
                    ..Default::default()
                }));
            } else {
                emit_prose(&p.text, &p.props, elements, &p.hyperlinks, &p.chp_runs);
            }
        }
    }
    table.flush(elements);
    flush_list(&mut list_items, list_ilfo.take(), list_formatting, elements);
}

/// Emit the accumulated list run as an `Element::List` and clear it.
///
/// `ilfo` is the run's own list identity (issue #250): resolved through
/// `list_formatting` to the declared start-at value and number format for
/// the run's base level. `None` (no `PlfLst`/`PlfLfo` data, or an `ilfo`
/// that doesn't resolve) degrades to the pre-#250 contract — bullet,
/// `start_number: None` — rather than erroring.
fn flush_list(
    items: &mut Vec<(u8, Vec<InlineContent>)>,
    ilfo: Option<i16>,
    list_formatting: &ListFormatting,
    elements: &mut Vec<Element>,
) {
    if items.is_empty() {
        return;
    }
    // `.doc` list levels are not guaranteed to start at 0; use the shallowest
    // level in this run as the base so a run whose top level is e.g. 1 is not
    // collapsed by `build_nested_list` (which would otherwise treat every
    // item as a child of the first and drop the rest when base_level is 0).
    let base_level = items.iter().map(|(lvl, _)| *lvl).min().unwrap_or(0);
    let level = ilfo.and_then(|ilfo| list_formatting.level_for(ilfo, base_level));
    let ordered = level.is_some_and(|l| l.is_numbered());
    let start_number =
        level.filter(|l| l.is_numbered() && l.start_at != 1).map(|l| l.start_at as u32);

    let mut list = build_nested_list(ordered, items, base_level);
    list.start_number = start_number;
    elements.push(Element::List(list));
    items.clear();
}

/// Turn a paragraph's text into inline content, preserving soft line breaks.
///
/// `sanitize_text` maps the Word soft-line-break control (`0x0B`) to `'\n'`;
/// within a single structured paragraph's inner text that is the only source
/// of `'\n'`, so each `'\n'` becomes an `InlineContent::LineBreak` instead of
/// being flattened into one run.
///
/// `hyperlinks` are `HYPERLINK` field display-text spans as byte ranges into
/// `text` (issue #249); `chp_runs` are character-property run boundaries,
/// also byte ranges into `text` (issue #287) — both empty for callers with
/// no per-paragraph data (e.g. the line-shape heading heuristic, which
/// works over the flat, already-sanitized document text rather than a
/// single `DocParagraph`).
fn inline_content_for(
    text: &str,
    hyperlinks: &[HyperlinkSpan],
    chp_runs: &[(std::ops::Range<usize>, ChpProps)],
) -> Vec<InlineContent> {
    let mut out = Vec::new();
    let mut base = 0usize;
    for seg in text.split('\n') {
        push_segment(seg, base, hyperlinks, chp_runs, &mut out);
        base += seg.len() + 1; // +1 for the '\n' the split consumed
        out.push(InlineContent::LineBreak);
    }
    // Drop the trailing `LineBreak` appended after the final segment.
    if matches!(out.last(), Some(InlineContent::LineBreak)) {
        out.pop();
    }
    if out.is_empty() {
        // Genuinely empty input: keep a single empty run so the content
        // vector is not zero-length.
        out.push(InlineContent::Text(TextSpan::plain(text)));
    }
    out
}

/// Apply one `ChpProps` to a plain `TextSpan` built from `text`.
fn styled_span(text: &str, props: &ChpProps) -> TextSpan {
    TextSpan {
        bold: props.bold,
        italic: props.italic,
        underline: props.underline.clone(),
        color: props.color,
        font_size_half_pt: props.font_size_half_pt,
        ..TextSpan::plain(text)
    }
}

/// Push one line-break-free segment of text as one or more `TextSpan`s,
/// splitting on every hyperlink and character-property (issue #287) run
/// boundary that overlaps it. `base` is `seg`'s own byte offset within the
/// original (pre-split) text, so `hyperlinks`'/`chp_runs`' ranges (computed
/// against that same original text) line up correctly.
fn push_segment(
    seg: &str,
    base: usize,
    hyperlinks: &[HyperlinkSpan],
    chp_runs: &[(std::ops::Range<usize>, ChpProps)],
    out: &mut Vec<InlineContent>,
) {
    if seg.is_empty() {
        return;
    }
    let seg_start = base;
    let seg_end = base + seg.len();

    // Every point where either a hyperlink or a CHP run starts/ends within
    // this segment is a potential span boundary; walking the sorted, deduped
    // union of both (plus the segment's own bounds) guarantees each
    // resulting sub-slice has one unambiguous hyperlink state and one
    // unambiguous `ChpProps` — `resolve_chp_segments` already guarantees
    // `chp_runs` fully and contiguously covers the paragraph, so no
    // sub-slice here can straddle two different CHP runs.
    let mut breakpoints: Vec<usize> = vec![seg_start, seg_end];
    for h in hyperlinks {
        if h.range.start > seg_start && h.range.start < seg_end {
            breakpoints.push(h.range.start);
        }
        if h.range.end > seg_start && h.range.end < seg_end {
            breakpoints.push(h.range.end);
        }
    }
    for (r, _) in chp_runs {
        if r.start > seg_start && r.start < seg_end {
            breakpoints.push(r.start);
        }
        if r.end > seg_start && r.end < seg_end {
            breakpoints.push(r.end);
        }
    }
    breakpoints.sort_unstable();
    breakpoints.dedup();

    for pair in breakpoints.windows(2) {
        let (a, b) = (pair[0], pair[1]);
        if a >= b {
            continue;
        }
        let slice = &seg[a - seg_start..b - seg_start];
        if slice.is_empty() {
            continue;
        }
        let hyperlink =
            hyperlinks.iter().find(|h| h.range.start <= a && b <= h.range.end);
        let props = chp_runs.iter().find(|(r, _)| r.start <= a && b <= r.end).map(|(_, p)| p);
        let mut span = match props {
            Some(p) => styled_span(slice, p),
            None => TextSpan::plain(slice),
        };
        if let Some(h) = hyperlink {
            span.hyperlink = Some(h.url.clone());
        }
        out.push(InlineContent::Text(span));
    }
}

/// Shift `hyperlinks`' byte ranges by `-trim_start` (the number of bytes
/// `text.trim()` removed from the front) and clip them to
/// `[0, trimmed_len]`, dropping any span that trimming removed entirely.
/// Needed because `emit_heading`/`emit_prose` call `inline_content_for` on
/// `text.trim()`, not `text` itself, so the spans (computed against the
/// untrimmed paragraph text) would otherwise point at the wrong bytes.
fn shift_hyperlinks_for_trim(
    hyperlinks: &[HyperlinkSpan],
    trim_start: usize,
    trimmed_len: usize,
) -> Vec<HyperlinkSpan> {
    hyperlinks
        .iter()
        .filter_map(|h| {
            let start = h.range.start.saturating_sub(trim_start).min(trimmed_len);
            let end = h.range.end.saturating_sub(trim_start).min(trimmed_len);
            (start < end).then(|| HyperlinkSpan { range: start..end, url: h.url.clone() })
        })
        .collect()
}

/// As [`shift_hyperlinks_for_trim`], for CHP run boundaries (issue #287).
fn shift_chp_runs_for_trim(
    chp_runs: &[(std::ops::Range<usize>, ChpProps)],
    trim_start: usize,
    trimmed_len: usize,
) -> Vec<(std::ops::Range<usize>, ChpProps)> {
    chp_runs
        .iter()
        .filter_map(|(r, props)| {
            let start = r.start.saturating_sub(trim_start).min(trimmed_len);
            let end = r.end.saturating_sub(trim_start).min(trimmed_len);
            (start < end).then(|| (start..end, props.clone()))
        })
        .collect()
}

/// Classify a prose paragraph as a heading or paragraph and push it.
///
/// Mirrors the line-based heuristic so a PAPX-bearing document keeps the same
/// heading/title detection as the fallback path.
/// Emit a heading at an explicit level, honouring soft line breaks.
fn emit_heading(
    text: &str,
    level: u8,
    elements: &mut Vec<Element>,
    hyperlinks: &[HyperlinkSpan],
    chp_runs: &[(std::ops::Range<usize>, ChpProps)],
) {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return;
    }
    let trim_start = text.len() - text.trim_start().len();
    let hyperlinks = shift_hyperlinks_for_trim(hyperlinks, trim_start, trimmed.len());
    let chp_runs = shift_chp_runs_for_trim(chp_runs, trim_start, trimmed.len());
    let mut content = inline_content_for(trimmed, &hyperlinks, &chp_runs);
    for ic in &mut content {
        if let InlineContent::Text(t) = ic {
            t.bold = true;
        }
    }
    elements.push(Element::Heading(Heading {
        level: level.clamp(1, 6),
        content,
        ..Default::default()
    }));
}

/// Guess headings from line shape — short, non-sentence, ALL-CAPS lines, or
/// a short opening line.
///
/// This is a guess and is only reached for documents that carry no
/// `sprmPOutLvl` at all. Running it alongside real outline levels produced
/// two disagreeing answers for the same paragraphs in one document.
fn emit_prose(
    text: &str,
    props: &PapProps,
    elements: &mut Vec<Element>,
    hyperlinks: &[HyperlinkSpan],
    chp_runs: &[(std::ops::Range<usize>, ChpProps)],
) {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return;
    }

    let is_heading = trimmed.len() < 100
        && !trimmed.ends_with('.')
        && !trimmed.ends_with(',')
        && !is_heading_guess_junk(trimmed)
        && (trimmed
            .chars()
            .filter(|c| c.is_alphabetic())
            .all(|c| c.is_uppercase())
            || (elements.is_empty() && trimmed.len() < 60));

    if is_heading {
        // Honour soft line breaks inside headings too: split on `'\n'` (the
        // sanitised form of `0x0B`) and keep each segment bold.
        let trim_start = text.len() - text.trim_start().len();
        let shifted = shift_hyperlinks_for_trim(hyperlinks, trim_start, trimmed.len());
        let shifted_chp = shift_chp_runs_for_trim(chp_runs, trim_start, trimmed.len());
        let mut content = inline_content_for(trimmed, &shifted, &shifted_chp);
        for ic in &mut content {
            if let InlineContent::Text(t) = ic {
                t.bold = true;
            }
        }
        elements.push(Element::Heading(Heading {
            level: if elements.is_empty() { 1 } else { 2 },
            content,
            ..Default::default()
        }));
    } else {
        elements.push(Element::Paragraph(Paragraph {
            content: inline_content_for(text, hyperlinks, chp_runs),
            tabs: props.tabs.clone(),
            alignment: props.alignment.clone(),
            indent_left_twips: props.indent_left_twips,
            indent_right_twips: props.indent_right_twips,
            first_line_indent_twips: props.first_line_indent_twips,
            space_before_twips: props.space_before_twips,
            space_after_twips: props.space_after_twips,
            ..Default::default()
        }));
    }
}

/// Reject line shapes the ALL-CAPS / short-opening-line heading guess in
/// `emit_prose` otherwise misclassifies as headings: pure separator lines,
/// bare dates, UK postcodes, and "CCY - symbol" currency labels — all
/// confirmed junk from the 246-file `.doc` corpus sweep behind issue #224
/// (`______________________________________________`, `11/16/2016`,
/// `SW8 5NQ`, `GBP - £`). The guess otherwise stays as-is — no reference
/// implementation does line-shape heading detection at all, so this only
/// narrows an already-approximate fallback rather than trying to perfect
/// it (`DRAFT`, a single common ALL-CAPS word, is left uncaught).
fn is_heading_guess_junk(line: &str) -> bool {
    (!line.chars().any(|c| c.is_alphanumeric()))
        || is_bare_date(line)
        || is_uk_postcode(line)
        || is_currency_label(line)
}

/// A line that is *only* `D[D]/-M[M]/-Y[YYY]`-shaped — a whole date with no
/// other text around it. Deliberately narrow: real headings almost never
/// consist of exactly three all-digit, slash/dash-separated groups.
fn is_bare_date(line: &str) -> bool {
    let parts: Vec<&str> = line.split(['/', '-']).collect();
    parts.len() == 3
        && parts.iter().all(|p| !p.is_empty() && p.chars().all(|c| c.is_ascii_digit()))
        && parts[0].len() <= 2
        && parts[1].len() <= 2
        && matches!(parts[2].len(), 2 | 4)
}

/// A UK postcode shape: `<1-2 letters><digit>[letter/digit] <digit><2
/// letters>` (e.g. `SW8 5NQ`). Not a full validator, just narrow enough
/// that a genuine heading is unlikely to match it by accident.
fn is_uk_postcode(line: &str) -> bool {
    let Some((outward, inward)) = line.rsplit_once(' ') else {
        return false;
    };
    let mut inward_chars = inward.chars();
    let inward_ok = inward.chars().count() == 3
        && inward_chars.next().is_some_and(|c| c.is_ascii_digit())
        && inward_chars.all(|c| c.is_ascii_uppercase());
    if !inward_ok {
        return false;
    }
    let outward: Vec<char> = outward.chars().collect();
    (2..=4).contains(&outward.len())
        && outward[0].is_ascii_uppercase()
        && outward.iter().any(|c| c.is_ascii_digit())
        && outward.iter().all(|c| c.is_ascii_alphanumeric())
}

/// A `"CCY - symbol"`-shaped currency label (`GBP - £`, `EUR - €`): a
/// 3-letter uppercase code followed by nothing but a currency symbol.
fn is_currency_label(line: &str) -> bool {
    let Some((code, rest)) = line.split_once(' ') else {
        return false;
    };
    if code.len() != 3 || !code.chars().all(|c| c.is_ascii_uppercase()) {
        return false;
    }
    let rest = rest.trim().strip_prefix('-').unwrap_or(rest).trim();
    !rest.is_empty() && !rest.chars().any(|c| c.is_alphanumeric())
}

// ---------------------------------------------------------------------------
// Fallback: line-based heuristic over sanitised text
// ---------------------------------------------------------------------------

fn line_heuristic(text: &str, elements: &mut Vec<Element>) {
    for line in text.lines() {
        emit_prose(line, &PapProps::default(), elements, &[], &[]);
    }
}

/// Build a `HeaderFooter` from one `PlcfHdd`-delimited story's already
/// sanitized text (issue #285).
fn text_to_header_footer(s: &str) -> HeaderFooter {
    HeaderFooter {
        content: s
            .lines()
            .filter(|l| !l.trim().is_empty())
            .map(|l| {
                Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain(l))],
                    ..Default::default()
                })
            })
            .collect(),
    }
}

/// Human-readable label for a `.doc` subdocument, used as the note marker.
fn subdocument_label(kind: crate::doc::SubDocumentKind) -> &'static str {
    use crate::doc::SubDocumentKind as K;
    match kind {
        K::Footnotes => "footnote",
        K::HeadersFooters => "header/footer",
        K::Comments => "comment",
        K::Endnotes => "endnote",
        K::TextBoxes => "text box",
        K::HeaderTextBoxes => "header text box",
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::{DocParagraph, PapProps, TapCellInfo, TapInfo};
    use crate::ir::{
        Element, InlineContent, Paragraph, TabAlignment, TabLeader, TabStop, TextSpan,
    };

    /// Build a `TapInfo` from column-center boundaries (twips).
    fn tap(centers: &[i16]) -> TapInfo {
        let itc_mac = centers.len() - 1;
        TapInfo {
            itc_mac: itc_mac as u8,
            centers: centers.to_vec(),
            cells: vec![TapCellInfo::default(); itc_mac],
        }
    }

    /// Regression for Medium #1: boundaries that differ by only a few twips
    /// across rows (Word rounds per row) must be snapped together, otherwise
    /// a spurious grid edge inflates `col_span` for every cell spanning it.
    #[test]
    fn column_edge_tolerance_collapses_near_boundaries() {
        // Row 0 and row 1 share edges 0/2000; row 1's middle edge is 3 twips
        // off (1003 vs 1000) — within `EDGE_TOLERANCE_TWIPS`.
        let rows = vec![
            PendingRow {
                cells: vec![vec![], vec![]],
                tap: Some(tap(&[0, 1000, 2000])),
                itap: 1,
            },
            PendingRow {
                cells: vec![vec![], vec![]],
                tap: Some(tap(&[0, 1003, 2000])),
                itap: 1,
            },
        ];
        let out = build_table_rows(&rows);
        assert_eq!(out.len(), 2);
        // Without tolerance the extra edge (1003) makes this cell span 2
        // columns; with tolerance it is a single column.
        assert_eq!(
            out[0].cells[1].col_span, 1,
            "near-identical column edges must be snapped (Medium #1)"
        );
    }

    /// Build a paragraph with the given text and PAP properties.
    fn para(text: &str, props: PapProps) -> DocParagraph {
        DocParagraph {
            text: text.to_string(),
            terminator: '\r',
            props,
            hyperlinks: Vec::new(),
            chp_runs: Vec::new(),
        }
    }

    /// issue #249 — a `HYPERLINK` field's display text must reach
    /// `TextSpan::hyperlink` in the IR, with the surrounding plain text on
    /// either side kept as ordinary, unlinked runs.
    #[test]
    fn hyperlink_span_reaches_textspan_hyperlink_in_the_ir() {
        let text = "Before text; Hyperlink text; after text.".to_string();
        let link_start = text.find("Hyperlink text").unwrap();
        let link_end = link_start + "Hyperlink text".len();
        let p = DocParagraph {
            text,
            terminator: '\r',
            props: PapProps::default(),
            hyperlinks: vec![crate::doc::HyperlinkSpan {
                range: link_start..link_end,
                url: "http://testuri.org/".to_string(),
            }],
            chp_runs: Vec::new(),
        };
        let mut els = Vec::new();
        walk_paragraphs(&[p], false, &mut els, &ListFormatting::default());
        let Element::Paragraph(par) = &els[0] else {
            panic!("expected a paragraph, got {:?}", els[0]);
        };
        let spans: Vec<&TextSpan> = par
            .content
            .iter()
            .filter_map(|c| match c {
                InlineContent::Text(t) => Some(t),
                _ => None,
            })
            .collect();
        let linked = spans
            .iter()
            .find(|s| s.hyperlink.is_some())
            .expect("a hyperlinked span must be present");
        assert_eq!(linked.text, "Hyperlink text");
        assert_eq!(linked.hyperlink.as_deref(), Some("http://testuri.org/"));
        assert!(
            spans.iter().any(|s| s.hyperlink.is_none() && s.text.contains("Before text")),
            "surrounding plain text must stay unlinked: {spans:?}"
        );
        assert!(
            spans.iter().any(|s| s.hyperlink.is_none() && s.text.contains("after text")),
            "surrounding plain text must stay unlinked: {spans:?}"
        );
    }

    /// Regression (issue #287): real per-run character formatting
    /// (bold/italic/underline/color/font-size) from CHPX reaches the IR's
    /// `TextSpan`s, split at exactly the right byte boundaries — not the
    /// always-`false`/`None` defaults from before this fix.
    #[test]
    fn test_chp_runs_reach_textspan_formatting_split_at_the_right_boundaries() {
        let text = "Plain bold italic.".to_string();
        let bold_start = text.find("bold").unwrap();
        let bold_end = bold_start + "bold".len();
        let italic_start = text.find("italic").unwrap();
        let italic_end = italic_start + "italic".len();
        let p = DocParagraph {
            text,
            terminator: '\r',
            props: PapProps::default(),
            hyperlinks: Vec::new(),
            chp_runs: vec![
                (
                    bold_start..bold_end,
                    ChpProps { bold: true, color: Some([255, 0, 0]), ..Default::default() },
                ),
                (
                    italic_start..italic_end,
                    ChpProps {
                        italic: true,
                        underline: Some(crate::ir::UnderlineStyle::Single),
                        font_size_half_pt: Some(28),
                        ..Default::default()
                    },
                ),
            ],
        };
        let mut els = Vec::new();
        walk_paragraphs(&[p], false, &mut els, &ListFormatting::default());
        let Element::Paragraph(par) = &els[0] else {
            panic!("expected a paragraph, got {:?}", els[0]);
        };
        let spans: Vec<&TextSpan> = par
            .content
            .iter()
            .filter_map(|c| match c {
                InlineContent::Text(t) => Some(t),
                _ => None,
            })
            .collect();

        let plain = spans.iter().find(|s| s.text == "Plain ").expect("leading plain span");
        assert!(!plain.bold && !plain.italic, "unformatted text must stay unformatted");

        let bold = spans.iter().find(|s| s.text == "bold").expect("bold span");
        assert!(bold.bold);
        assert!(!bold.italic);
        assert_eq!(bold.color, Some([255, 0, 0]));

        let italic = spans.iter().find(|s| s.text == "italic").expect("italic span");
        assert!(italic.italic);
        assert!(!italic.bold);
        assert_eq!(italic.underline, Some(crate::ir::UnderlineStyle::Single));
        assert_eq!(italic.font_size_half_pt, Some(28));

        assert!(
            spans.iter().any(|s| s.text == " " && !s.bold && !s.italic),
            "the gap between the two formatted runs must stay a separate, unformatted span: {spans:?}"
        );
    }

    /// Regression (issue #287): paragraph alignment/indentation/spacing
    /// from PAP SPRMs reach the IR's `Paragraph`, not the always-`None`
    /// defaults from before this fix.
    #[test]
    fn test_pap_alignment_indent_and_spacing_reach_the_ir_paragraph() {
        let props = PapProps {
            alignment: Some(crate::ir::ParagraphAlignment::Center),
            indent_left_twips: Some(720),
            indent_right_twips: Some(360),
            first_line_indent_twips: Some(-360),
            space_before_twips: Some(200),
            space_after_twips: Some(100),
            ..Default::default()
        };
        let p = para("A centered, indented paragraph.", props);
        let mut els = Vec::new();
        walk_paragraphs(&[p], false, &mut els, &ListFormatting::default());
        let Element::Paragraph(par) = &els[0] else {
            panic!("expected a paragraph, got {:?}", els[0]);
        };
        assert_eq!(par.alignment, Some(crate::ir::ParagraphAlignment::Center));
        assert_eq!(par.indent_left_twips, Some(720));
        assert_eq!(par.indent_right_twips, Some(360));
        assert_eq!(par.first_line_indent_twips, Some(-360));
        assert_eq!(par.space_before_twips, Some(200));
        assert_eq!(par.space_after_twips, Some(100));
    }

    /// Medium #4: a soft line break (`0x0B`, which `sanitize_text` maps to
    /// `'\n'`) inside a paragraph must survive as an `InlineContent::LineBreak`
    /// rather than being flattened into one run.
    #[test]
    fn soft_line_break_becomes_inline_break() {
        // Trailing '.' keeps this out of the heading heuristic so it routes to
        // a Paragraph (where soft breaks are honoured).
        let p = para(
            "First line of the paragraph.\nSecond line of the paragraph.",
            PapProps::default(),
        );
        let mut els = Vec::new();
        walk_paragraphs(&[p], false, &mut els, &ListFormatting::default());
        let Element::Paragraph(par) = &els[0] else {
            panic!("expected a paragraph, got {:?}", els[0]);
        };
        assert!(
            par.content
                .iter()
                .any(|c| matches!(c, InlineContent::LineBreak)),
            "soft line break must produce an InlineContent::LineBreak"
        );
        let texts: Vec<&str> = par
            .content
            .iter()
            .filter_map(|c| match c {
                InlineContent::Text(t) => Some(t.text.as_str()),
                _ => None,
            })
            .collect();
        assert_eq!(
            texts,
            vec![
                "First line of the paragraph.",
                "Second line of the paragraph."
            ]
        );
    }

    /// Medium #2: a nested table (`itap > 1`) is flattened, not silently
    /// mis-rendered — a visible notice paragraph must accompany the table.
    #[test]
    fn nested_table_itap_emits_notice() {
        let props = PapProps {
            is_table_trailing_mark: true,
            itap: 2, // nested
            tap: Some(tap(&[0, 1000, 2000])),
            ..PapProps::default()
        };
        let row = para("", props);

        let mut els = Vec::new();
        walk_paragraphs(&[row], false, &mut els, &ListFormatting::default());

        assert!(
            els.iter().any(|e| matches!(e, Element::Table(_))),
            "the (flattened) table must still be emitted"
        );
        assert!(
            els.iter().any(|e| matches!(
                e,
                Element::Paragraph(p)
                    if p.content.iter().any(|c| matches!(
                        c,
                        InlineContent::Text(t) if t.text.contains("nested table")
                    ))
            )),
            "a nested-table notice must be emitted (no silent flattening)"
        );
    }

    /// List membership is keyed on `ilfo` (`sprmPIlfo`, `0x460B`), per
    /// [MS-DOC] §2.4.6.3. `0x0000` and `0xF801` mean "not in a list" and the
    /// paragraph must be ordinary prose even when an `ilvl` SPRM is present.
    #[test]
    fn ilfo_not_in_list_bands_are_prose() {
        // `0x0000` (0) and `0xF801` (-2047 as signed i16) are the two
        // documented "not in a list" markers. A non-spec value (2047) is also
        // prose because it falls outside every valid band.
        for ilfo in [0i16, -2047, 2047] {
            let props = PapProps {
                ilvl: Some(0),
                ilfo: Some(ilfo),
                ..PapProps::default()
            };
            let p = para("Not a list item.", props);
            let mut els = Vec::new();
            walk_paragraphs(&[p], false, &mut els, &ListFormatting::default());
            assert!(
                !els.iter().any(|e| matches!(e, Element::List(_))),
                "ilfo {ilfo:#06x} must not build a list"
            );
            assert!(
                els.iter().any(|e| matches!(e, Element::Paragraph(_))),
                "ilfo {ilfo:#06x} must be emitted as ordinary prose"
            );
        }
    }

    /// issue #223 — a paragraph with a real outline level (heading) that
    /// *also* carries a valid `ilfo` (Word's multilevel-list "Heading"
    /// gallery attaches numPr/ilfo to the heading styles themselves) must
    /// come out as a Heading, not a ListItem — this is the normal shape
    /// of a numbered heading ("1. Introduction") in real documents.
    #[test]
    fn test_numbered_heading_wins_over_list_membership() {
        let props = PapProps {
            ilvl: Some(0),
            ilfo: Some(1), // valid 1-based list index
            outline_level: Some(0), // Heading 1
            ..PapProps::default()
        };
        let p = para("1. Introduction", props);
        let mut els = Vec::new();
        walk_paragraphs(&[p], false, &mut els, &ListFormatting::default());

        assert!(
            els.iter().any(|e| matches!(e, Element::Heading(h) if h.level == 1)),
            "a numbered Heading 1 must be emitted as a Heading, got {els:#?}"
        );
        assert!(
            !els.iter().any(|e| matches!(e, Element::List(_))),
            "list membership must be dropped once the paragraph is a heading, got {els:#?}"
        );
    }

    /// `0xF802`–`0xFFFF` is the negation of a 1-based index and is still a list
    /// item (see TODO(ilfo-negated)); it must not be dropped to prose.
    #[test]
    fn ilfo_negated_band_is_list() {
        let props = PapProps {
            ilvl: Some(0),
            ilfo: Some(-2046), // 0xF802
            ..PapProps::default()
        };
        let p = para("A list item via the negated band.", props);
        let mut els = Vec::new();
        walk_paragraphs(&[p], false, &mut els, &ListFormatting::default());
        assert!(
            els.iter().any(|e| matches!(e, Element::List(_))),
            "0xF802 (negated index) must still be a list item"
        );
    }

    /// issue #250 — a list run's `ilfo` must resolve through `PlfLfo`'s
    /// `lsid` to the matching `PlfLst` entry's `LVL`, surfacing a real
    /// `start_number` and `ordered = true` instead of always `None`/bullet.
    #[test]
    fn list_start_number_and_ordered_resolve_from_list_formatting() {
        let list_formatting = crate::doc::ListFormatting::from_parts(
            vec![(
                0x44F53D09, // lsid
                vec![
                    crate::doc::ListLevel { start_at: 5, nfc: 0x00 }, // level 0: numbered, starts at 5
                    crate::doc::ListLevel { start_at: 1, nfc: 0xFF }, // level 1: bullet
                ],
            )],
            vec![0x44F53D09], // lfo_lsids[0] == lsid above, so ilfo=1 resolves to it
        );
        let props = PapProps {
            ilvl: Some(0),
            ilfo: Some(1),
            ..PapProps::default()
        };
        let p = para("First", props);
        let mut els = Vec::new();
        walk_paragraphs(&[p], false, &mut els, &list_formatting);

        let Element::List(list) = &els[0] else {
            panic!("expected a List element, got {els:#?}");
        };
        assert!(list.ordered, "nfc != 0xFF must render as an ordered list");
        assert_eq!(list.start_number, Some(5), "iStartAt = 5 must reach List::start_number");
    }

    /// A level whose `nfc == 0xFF` (a bullet level) must never report
    /// `start_number`, even when its `iStartAt` happens to be non-1.
    #[test]
    fn bullet_level_never_gets_a_start_number() {
        let list_formatting = crate::doc::ListFormatting::from_parts(
            vec![(1, vec![crate::doc::ListLevel { start_at: 7, nfc: 0xFF }])],
            vec![1],
        );
        let props = PapProps {
            ilvl: Some(0),
            ilfo: Some(1),
            ..PapProps::default()
        };
        let p = para("Bulleted", props);
        let mut els = Vec::new();
        walk_paragraphs(&[p], false, &mut els, &list_formatting);

        let Element::List(list) = &els[0] else {
            panic!("expected a List element, got {els:#?}");
        };
        assert!(!list.ordered);
        assert_eq!(list.start_number, None);
    }

    /// A declared `iStartAt == 1` (the default) must not set
    /// `start_number` — only an explicit override is worth surfacing,
    /// mirroring DOCX's identical contract.
    #[test]
    fn start_at_one_is_not_surfaced_as_an_override() {
        let list_formatting = crate::doc::ListFormatting::from_parts(
            vec![(1, vec![crate::doc::ListLevel { start_at: 1, nfc: 0x00 }])],
            vec![1],
        );
        let props = PapProps {
            ilvl: Some(0),
            ilfo: Some(1),
            ..PapProps::default()
        };
        let p = para("Numbered from 1", props);
        let mut els = Vec::new();
        walk_paragraphs(&[p], false, &mut els, &list_formatting);

        let Element::List(list) = &els[0] else {
            panic!("expected a List element, got {els:#?}");
        };
        assert!(list.ordered);
        assert_eq!(list.start_number, None, "start_at == 1 is the default, not an override");
    }

    /// An `ilfo` with no matching `PlfLfo`/`PlfLst` data (the common case
    /// for most existing tests, and for a real file with no `PlfLst` at
    /// all) must degrade to the pre-#250 contract — bullet, no
    /// start_number — not panic or produce a wrong-but-confident answer.
    #[test]
    fn missing_list_formatting_degrades_to_the_old_contract() {
        let props = PapProps {
            ilvl: Some(0),
            ilfo: Some(1),
            ..PapProps::default()
        };
        let p = para("Item", props);
        let mut els = Vec::new();
        walk_paragraphs(&[p], false, &mut els, &ListFormatting::default());

        let Element::List(list) = &els[0] else {
            panic!("expected a List element, got {els:#?}");
        };
        assert!(!list.ordered);
        assert_eq!(list.start_number, None);
    }

    /// `sprmPChgTabs` tab stops decoded onto `PapProps` must surface on the
    /// produced paragraph's `tabs`.
    #[test]
    fn pchg_tabs_surfaced_on_paragraph() {
        let props = PapProps {
            tabs: vec![TabStop {
                position_twips: 1440,
                alignment: TabAlignment::Center,
                leader: TabLeader::None,
            }],
            ..PapProps::default()
        };
        let p = para("Indented text carrying tab stops.", props);

        let mut els = Vec::new();
        walk_paragraphs(&[p], false, &mut els, &ListFormatting::default());
        let Element::Paragraph(par) = &els[0] else {
            panic!("expected a paragraph, got {:?}", els[0]);
        };
        assert_eq!(par.tabs.len(), 1, "decoded tab stops must reach the IR");
        assert_eq!(par.tabs[0].position_twips, 1440);
    }

    /// A table cell's content is a paragraph flagged `f_in_table` (but not a row
    /// terminator). `walk_paragraphs` must route it through
    /// `table.add_cell_paragraph` (convert_doc.rs:362-363), the `f_in_table`
    /// dispatch branch. A lone cell paragraph makes no row, so wrap it between
    /// two row-terminators the way a real `.doc` lays out a one-cell table.
    #[test]
    fn in_table_paragraph_becomes_cell() {
        let mark = |itap: u8| DocParagraph {
            text: String::new(),
            terminator: '\r',
            props: PapProps {
                is_table_trailing_mark: true,
                itap,
                ..PapProps::default()
            },
            hyperlinks: Vec::new(),
            chp_runs: Vec::new(),
        };
        let cell = DocParagraph {
            text: "cell text".into(),
            terminator: '\u{7}', // closes the cell
            props: PapProps {
                f_in_table: true,
                ..PapProps::default()
            },
            hyperlinks: Vec::new(),
            chp_runs: Vec::new(),
        };
        let paragraphs = [mark(1), cell, mark(1)];
        let mut els = Vec::new();
        walk_paragraphs(&paragraphs, false, &mut els, &ListFormatting::default());
        assert!(
            els.iter().any(|e| matches!(e, Element::Table(_))),
            "f_in_table cell paragraph must be emitted inside a table"
        );
    }

    /// Build a single-column `PendingRow` whose cell carries the given `rgf`
    /// (the MS-DOC 2-bit `TCGRF` vertical-merge field) and renders "cell".
    fn vmerge_row(rgf: u16) -> PendingRow {
        let cell = TapCellInfo {
            rgf,
            ..Default::default()
        };
        PendingRow {
            cells: vec![vec![Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain("cell"))],
                ..Default::default()
            })]],
            tap: Some(TapInfo {
                itc_mac: 1,
                centers: vec![0, 1000],
                cells: vec![cell],
            }),
            itap: 1,
        }
    }

    /// vertMerge 2-bit fix: two *independent* 2-row vertical merges encoded as
    /// `[fvmRestart, fvmMerge, fvmRestart, fvmMerge]` (MS-DOC 2-bit field:
    /// `fvmRestart = 0x0060`, `fvmMerge = 0x0020`) must not be folded into a
    /// single `row_span = 4` cell. Each `fvmRestart` starts its own merge;
    /// each merge is 2 rows; both restart cells must be rendered and rows 1&3
    /// (the `fvmMerge` continuations) absorbed.
    #[test]
    fn two_independent_two_row_merges_do_not_fold() {
        let rows = vec![
            vmerge_row(0x0060), // merge A restart
            vmerge_row(0x0020), // merge A continuation
            vmerge_row(0x0060), // merge B restart
            vmerge_row(0x0020), // merge B continuation
        ];

        let out = build_table_rows(&rows);
        assert_eq!(out.len(), 4, "all four rows must be present");

        // Merge A: restart cell rendered with span 2; continuation row absorbed.
        assert_eq!(out[0].cells.len(), 1, "row 0 emits its restart cell");
        assert_eq!(out[0].cells[0].row_span, 2, "merge A spans exactly 2 rows (not folded into B)");
        assert!(out[1].cells.is_empty(), "row 1 fvmMerge continuation must be absorbed");

        // Merge B: a *second* restart cell, also span 2 — must not be dropped.
        assert_eq!(out[2].cells.len(), 1, "row 2 emits its restart cell");
        assert_eq!(out[2].cells[0].row_span, 2, "merge B spans exactly 2 rows, independent of A");
        assert!(out[3].cells.is_empty(), "row 3 fvmMerge continuation must be absorbed");
    }

    /// Removing the "whole-TAP-copy" quirk heuristic: a `fvmRestart` (`0x0060`)
    /// always opens a brand-new merge, never a continuation. The run
    /// `[fvmRestart, fvmRestart, fvmMerge]` must therefore render row 1 as its
    /// own 2-row merge — the middle cell must not be silently absorbed into the
    /// merge above it, which is exactly the wrong output the old heuristic
    /// produced for two independent adjacent merges.
    #[test]
    fn restart_after_restart_is_distinct_merge() {
        let rows = vec![vmerge_row(0x0060), vmerge_row(0x0060), vmerge_row(0x0020)];

        let out = build_table_rows(&rows);
        assert_eq!(out.len(), 3, "all three rows must be present");
        // Row 0: a restart with no merge continuation below it → span 1.
        assert_eq!(out[0].cells.len(), 1, "row 0 emits its restart cell");
        assert_eq!(out[0].cells[0].row_span, 1, "row 0 merge has no continuation below");
        // Middle cell MUST be emitted, not absorbed into row 0.
        assert_eq!(out[1].cells.len(), 1, "row 1 MUST be emitted, not absorbed into row 0");
        assert_eq!(out[1].cells[0].row_span, 2, "row 1 merge spans 2 rows (row 1 + row 2)");
        // Row 2 is the continuation of row 1's merge.
        assert!(out[2].cells.is_empty(), "row 2 fvmMerge continuation must be absorbed");
    }

    /// Plain single 3-row merge via `[fvmRestart, fvmMerge, fvmMerge]`.
    #[test]
    fn single_three_row_merge_spans_three() {
        let rows = vec![vmerge_row(0x0060), vmerge_row(0x0020), vmerge_row(0x0020)];

        let out = build_table_rows(&rows);
        assert_eq!(out.len(), 3);
        assert_eq!(out[0].cells[0].row_span, 3, "one continuous 3-row merge");
        assert!(out[1].cells.is_empty());
        assert!(out[2].cells.is_empty());
    }

    // ── Line-shape heading guess tightening (issue #224) ────────────────────

    fn is_heading_guess(text: &str) -> bool {
        let mut els = Vec::new();
        emit_prose(text, &PapProps::default(), &mut els, &[], &[]);
        matches!(els.as_slice(), [Element::Heading(_)])
    }

    /// A pure separator/underscore line must never become a heading —
    /// found ×4 in a real corpus file (`ob_is.doc`) behind #224.
    #[test]
    fn underscore_separator_line_is_not_a_heading() {
        assert!(!is_heading_guess("______________________________________________"));
    }

    /// A bare `MM/DD/YYYY` date, found as a false-positive heading in a
    /// real corpus file behind #224, must not be promoted.
    #[test]
    fn bare_date_is_not_a_heading() {
        assert!(!is_heading_guess("11/16/2016"));
    }

    /// A UK postcode shape (`SW8 5NQ`), found as a false-positive heading
    /// in a real corpus file behind #224, must not be promoted.
    #[test]
    fn uk_postcode_is_not_a_heading() {
        assert!(!is_heading_guess("SW8 5NQ"));
    }

    /// `"CCY - symbol"` currency labels (`GBP - £`, `EUR - €`), found as
    /// false-positive headings behind #224, must not be promoted.
    #[test]
    fn currency_label_is_not_a_heading() {
        assert!(!is_heading_guess("GBP - £"));
        assert!(!is_heading_guess("EUR - €"));
    }

    /// The tightening must not touch real ALL-CAPS section headers the
    /// guess correctly caught before #224 (e.g. `parentinvguid.doc`'s
    /// un-styled section headings).
    #[test]
    fn genuine_all_caps_headings_still_promoted() {
        for heading in ["INTRODUCTION", "TABLE OF CONTENTS", "A. GENERAL INFORMATION"] {
            assert!(is_heading_guess(heading), "{heading:?} must still be promoted");
        }
    }

    /// The bare-date rule only rejects a line that is *entirely* a date —
    /// a date-shaped substring inside otherwise heading-shaped text must
    /// still be promoted.
    #[test]
    fn date_substring_inside_otherwise_heading_shaped_text_is_unaffected() {
        assert!(is_heading_guess("Report 11/16/2016 Summary"));
    }
}
