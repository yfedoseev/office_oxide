use crate::doc::{DocDocument, DocParagraph, TapCellInfo, TapInfo};
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
    if !paragraphs.is_empty() {
        walk_paragraphs(paragraphs, &mut elements);
    } else {
        line_heuristic(doc.plain_text_ref(), &mut elements);
    }

    let title = elements.iter().find_map(|e| match e {
        Element::Heading(h) => h.content.first().and_then(|c| match c {
            InlineContent::Text(t) => Some(t.text.clone()),
            _ => None,
        }),
        _ => None,
    });

    DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Doc,
            title: title.clone(),
            ..Default::default()
        },
        sections: vec![Section {
            title,
            elements,
            ..Default::default()
        }],
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
                content: inline_content_for(&p.text),
                // `tabs` is always empty in the tables-only build (it is
                // populated by the list/tab-stop PR); cloning keeps the IR
                // shape uniform with the list path.
                tabs: p.props.tabs.clone(),
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

fn walk_paragraphs(paragraphs: &[DocParagraph], elements: &mut Vec<Element>) {
    let mut table = TableBuilder::new();
    let mut list_items: Vec<(u8, Vec<InlineContent>)> = Vec::new();

    for p in paragraphs {
        if p.props.is_table_trailing_mark {
            flush_list(&mut list_items, elements);
            table.end_row(p.props.tap.clone(), p.props.itap);
        } else if p.props.f_in_table {
            flush_list(&mut list_items, elements);
            table.add_cell_paragraph(p);
        } else if is_doc_list_item(p.props.ilfo) {
            // List membership is keyed on `ilfo` (sprmPIlfo, `0x460B`), not on
            // `ilvl`: per [MS-DOC] §2.4.6.3 a paragraph is a list item only when
            // its `ilfo` is a valid list index. `ilvl` still drives nesting.
            table.flush(elements);
            let ilvl = p.props.ilvl.unwrap_or(0);
            list_items.push((ilvl, inline_content_for(&p.text)));
        } else {
            table.flush(elements);
            flush_list(&mut list_items, elements);
            emit_prose(&p.text, &p.props.tabs, elements);
        }
    }
    table.flush(elements);
    flush_list(&mut list_items, elements);
}

/// Emit the accumulated list run as an `Element::List` and clear it.
fn flush_list(items: &mut Vec<(u8, Vec<InlineContent>)>, elements: &mut Vec<Element>) {
    if items.is_empty() {
        return;
    }
    // `.doc` list levels are not guaranteed to start at 0; use the shallowest
    // level in this run as the base so a run whose top level is e.g. 1 is not
    // collapsed by `build_nested_list` (which would otherwise treat every
    // item as a child of the first and drop the rest when base_level is 0).
    let base_level = items.iter().map(|(lvl, _)| *lvl).min().unwrap_or(0);
    // `ordered = false` (bullet) — see `walk_paragraphs` for the limitation.
    let list = build_nested_list(false, items, base_level);
    elements.push(Element::List(list));
    items.clear();
}

/// Turn a paragraph's text into inline content, preserving soft line breaks.
///
/// `sanitize_text` maps the Word soft-line-break control (`0x0B`) to `'\n'`;
/// within a single structured paragraph's inner text that is the only source
/// of `'\n'`, so each `'\n'` becomes an `InlineContent::LineBreak` instead of
/// being flattened into one run.
fn inline_content_for(text: &str) -> Vec<InlineContent> {
    let mut out = Vec::new();
    for seg in text.split('\n') {
        if !seg.is_empty() {
            out.push(InlineContent::Text(TextSpan::plain(seg)));
        }
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

/// Classify a prose paragraph as a heading or paragraph and push it.
///
/// Mirrors the line-based heuristic so a PAPX-bearing document keeps the same
/// heading/title detection as the fallback path.
fn emit_prose(text: &str, tabs: &[TabStop], elements: &mut Vec<Element>) {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return;
    }

    let is_heading = trimmed.len() < 100
        && !trimmed.ends_with('.')
        && !trimmed.ends_with(',')
        && (trimmed
            .chars()
            .filter(|c| c.is_alphabetic())
            .all(|c| c.is_uppercase())
            || (elements.is_empty() && trimmed.len() < 60));

    if is_heading {
        // Honour soft line breaks inside headings too: split on `'\n'` (the
        // sanitised form of `0x0B`) and keep each segment bold.
        let mut content = inline_content_for(trimmed);
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
            content: inline_content_for(text),
            tabs: tabs.to_vec(),
            ..Default::default()
        }));
    }
}

// ---------------------------------------------------------------------------
// Fallback: line-based heuristic over sanitised text
// ---------------------------------------------------------------------------

fn line_heuristic(text: &str, elements: &mut Vec<Element>) {
    for line in text.lines() {
        emit_prose(line, &[], elements);
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
        }
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
        walk_paragraphs(&[p], &mut els);
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
        walk_paragraphs(&[row], &mut els);

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
            walk_paragraphs(&[p], &mut els);
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
        walk_paragraphs(&[p], &mut els);
        assert!(
            els.iter().any(|e| matches!(e, Element::List(_))),
            "0xF802 (negated index) must still be a list item"
        );
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
        walk_paragraphs(&[p], &mut els);
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
        };
        let cell = DocParagraph {
            text: "cell text".into(),
            terminator: '\u{7}', // closes the cell
            props: PapProps {
                f_in_table: true,
                ..PapProps::default()
            },
        };
        let paragraphs = [mark(1), cell, mark(1)];
        let mut els = Vec::new();
        walk_paragraphs(&paragraphs, &mut els);
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
}
