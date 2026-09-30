use crate::format::DocumentFormat;
use crate::ir::*;

/// Maximum worksheet rows materialised into the IR per sheet.
///
/// The same cap `convert_xlsx` applies, and for the same reason: a sheet is
/// converted eagerly into in-memory IR, so an unbounded one builds millions
/// of cell allocations. `convert_xls` had no cap, and a 31 KB `.xls`
/// declaring a huge used range reached 8.5 GB and was killed by the OOM
/// killer — a crash no caller can catch. Excess rows are dropped and
/// flagged with a visible notice.
/// Budget in materialised cells, not rows. The declared grid is padded to
/// the used range, so a row limit is measured against padding rather than
/// content: a sheet whose real data sits past the limit in a mostly-empty
/// grid loses it. Trailing empty cells are trimmed before anything counts
/// against this, so a 65,536-row sheet of padding costs almost nothing and
/// only genuinely dense sheets can reach the cap.
#[cfg(not(test))]
const MAX_CELLS_PER_SHEET: usize = 1_000_000;
/// Unit tests exercise the behaviour around the cap, not the constant, and a
/// fixture large enough to reach a million cells costs hundreds of megabytes
/// in a debug build — enough to exhaust a CI runner once the harness runs
/// tests in parallel.
#[cfg(test)]
const MAX_CELLS_PER_SHEET: usize = 5_000;

use crate::xlsx::range_sweep::{CellRange, RangeSweep};

/// `Sheet::merged_cells` ((row_first, row_last, col_first, col_last)
/// tuples from MERGEDCELLS) as a row sweep. Each range is kept whole and
/// consulted per materialised row — the anchor gets its span, covered
/// positions are dropped — the same sparse, span-driven TableRow model
/// `ir_render`'s `table_grid` expects. Expanding each range into the set of
/// positions it covers cost its declared area: one whole-grid entry was
/// billions of inserts.
fn merge_sweep(merged_cells: &[(u16, u16, u16, u16)]) -> RangeSweep<()> {
    RangeSweep::new(
        merged_cells
            .iter()
            .map(|&(r0, r1, c0, c1)| {
                CellRange::from_corners(r0.into(), c0.into(), r1.into(), c1.into())
            })
            .filter(|r| r.row_span() > 1 || r.col_span() > 1)
            .map(|r| (r, ()))
            .collect(),
    )
}

/// `Sheet::hyperlinks` (one entry per `HLINK` record, which can cover a
/// range, not just a single cell) as a row sweep; a later record wins
/// where ranges overlap.
fn hyperlink_sweep(hyperlinks: &[crate::xls::XlsHyperlink]) -> RangeSweep<&str> {
    RangeSweep::new(
        hyperlinks
            .iter()
            .map(|hl| {
                let range = CellRange::from_corners(
                    hl.row_first.into(),
                    hl.col_first.into(),
                    hl.row_last.into(),
                    hl.col_last.into(),
                );
                (range, hl.target.as_str())
            })
            .collect(),
    )
}

/// Clip `range`'s columns to a row of `width` cells: `None` when it lies
/// wholly past the row's last cell.
fn clip_cols(range: &CellRange, width: usize) -> Option<(usize, usize)> {
    let lo = range.col_lo as usize;
    if lo >= width {
        return None;
    }
    Some((lo, (range.col_hi as usize).min(width - 1)))
}

pub(crate) fn xls_to_ir(doc: &crate::xls::XlsDocument) -> DocumentIR {
    let mut sections = Vec::new();

    for sheet in &doc.sheets {
        let mut merges = merge_sweep(&sheet.merged_cells);
        let mut links = hyperlink_sweep(&sheet.hyperlinks);
        let mut rows = Vec::new();
        // Rows past the last one carrying data are padding; measuring the
        // sheet against them would report a truncation that dropped nothing.
        let total_rows = sheet
            .rows
            .iter()
            .enumerate()
            .rposition(|(_, row)| {
                // Matches the emptiness test the emit loop applies, but
                // without rendering every cell of the declared grid.
                row.iter()
                    .any(|cell_value| !matches!(cell_value, crate::xls::CellValue::Empty))
            })
            .map_or(0, |i| i + 1);
        // Widest row, for clipping a merge's column span to the table.
        let max_width = sheet
            .rows
            .iter()
            .take(total_rows)
            .map(Vec::len)
            .max()
            .unwrap_or(0);
        let mut budget = MAX_CELLS_PER_SHEET;
        // Rows reached before the budget ran out, counted separately from
        // `rows` so that trimming an all-empty tail is not reported as
        // truncation.
        let mut rows_scanned = 0usize;

        for (row_idx, row) in sheet.rows.iter().take(total_rows).enumerate() {
            if budget == 0 {
                break;
            }
            rows_scanned += 1;
            // Only the columns up to the last non-empty one are built. The
            // trailing padding used to be materialised as full IR cells and
            // then popped — but `Vec::pop` keeps the buffer, so every one of
            // a 65,536-row grid's empty rows retained a 256-cell allocation
            // (~75 KB): 4.9 GB for a 42 KB file whose IR is 6 MB.
            let width = row
                .iter()
                .rposition(|cell_value| !matches!(cell_value, crate::xls::CellValue::Empty))
                .map_or(0, |i| i + 1);
            let row_links = row_hyperlinks(&mut links, row_idx as u32, width);
            let mut cells = Vec::with_capacity(width);
            for col_idx in 0..width {
                // Number-format-aware rendering: a date cell is an ISO
                // date rather than its raw serial.
                let text = sheet
                    .display_text(row_idx, col_idx)
                    .map(std::borrow::Cow::into_owned)
                    .unwrap_or_default();
                let hyperlink = row_links
                    .as_ref()
                    .and_then(|l| l[col_idx])
                    .map(str::to_string);
                cells.push(TableCell {
                    content: vec![Element::Paragraph(Paragraph {
                        content: if text.is_empty() {
                            Vec::new()
                        } else {
                            let mut span = TextSpan::plain(text);
                            span.hyperlink = hyperlink;
                            vec![InlineContent::Text(span)]
                        },
                        ..Default::default()
                    })],
                    col_span: 1,
                    row_span: 1,
                    ..Default::default()
                });
            }

            // Apply merges: the anchor gets its real span, and every
            // position a range covers is dropped from the row entirely.
            // `sheet.rows` is dense, so `row_idx`/`col_idx` are already the
            // absolute positions `merged_cells` uses.
            if !merges.is_empty() {
                apply_merges(&mut merges, &mut cells, row_idx, total_rows, max_width);
            }

            // Drop trailing empty cells. A BIFF sheet reports the whole
            // declared grid — one file here is 65,536 x 256, essentially all
            // empty — and materialising the padding built 16.7M IR cells and
            // ran the process out of memory. `convert_xlsx` has always
            // trimmed; this path never did.
            while cells.last().is_some_and(|c: &TableCell| cell_is_empty(c)) {
                cells.pop();
            }
            budget = budget.saturating_sub(cells.len());

            rows.push(TableRow {
                cells,
                is_header: row_idx == 0,
                ..Default::default()
            });
        }

        // Trailing all-empty rows go the same way as trailing empty cells.
        while rows.last().is_some_and(|r| r.cells.is_empty()) {
            rows.pop();
        }

        let mut elements = if rows.is_empty() {
            Vec::new()
        } else {
            vec![Element::Table(Table {
                rows,
                ..Default::default()
            })]
        };

        // Truncation is stated in the content rather than left silent,
        // matching the XLSX path.
        if rows_scanned < total_rows {
            let omitted = total_rows - rows_scanned;
            elements.push(Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(format!(
                    "[{omitted} of {total_rows} rows not shown — worksheet truncated at \
                     {rows_scanned} rows]"
                )))],
                ..Default::default()
            }));
        }

        // Range lookups whose work budget ran out (only a file declaring
        // vast numbers of overlapping ranges gets there) say so.
        for (sweep_exhausted, what) in [
            (merges.exhausted(), "merged cell ranges"),
            (links.exhausted(), "hyperlink ranges"),
        ] {
            if sweep_exhausted {
                elements.push(Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain(format!(
                        "[some {what} not applied — too many overlapping ranges]"
                    )))],
                    ..Default::default()
                }));
            }
        }

        // Cell comments are document content, and were never surfaced at
        // all before — appended as endnotes so every
        // renderer sees them, the same convention convert_xlsx.rs uses
        // for its own comments.
        for (i, c) in sheet.comments.iter().enumerate() {
            let cell_ref = crate::xls::condfmt::cell_ref(c.row, c.col);
            let marker = match c.author.as_deref() {
                Some(a) => format!("{cell_ref} ({a})"),
                None => cell_ref,
            };
            elements.push(Element::Endnote(Note {
                id: i as u32,
                marker: Some(marker),
                author: c.author.clone(),
                content: vec![Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain(c.text.clone()))],
                    ..Default::default()
                })],
            }));
        }

        sections.push(Section {
            title: Some(sheet.name.clone()),
            elements,
            // A hidden sheet is kept and flagged, not dropped — the same
            // contract `convert_xlsx` already honours.
            hidden: sheet.hidden,
            conditional_formats: sheet.conditional_formats.clone(),
            data_validations: sheet.data_validations.clone(),
            ..Default::default()
        });
    }

    // Extracted pictures never reached the IR, so every image in a legacy
    // workbook was silently dropped on conversion. Append them to the last
    // section; BIFF drawings carry no reliable per-sheet anchor here.
    append_legacy_images(&mut sections, doc.images());

    // Charts weren't rendered at all, only their own text (series names,
    // trendline names/labels, axis/chart titles) was recovered from
    // `SeriesText` records — surfacing it as a dedicated section keeps
    // every human-meaningful word in the workbook reachable, the same
    // contract `convert_xlsx` already honours for its own charts.
    if !doc.chart_text().is_empty() {
        let mut chart_elements: Vec<Element> = Vec::new();
        for (i, text) in doc.chart_text().iter().enumerate() {
            chart_elements.push(Element::Heading(Heading {
                level: 3,
                content: vec![InlineContent::Text(TextSpan::plain(format!(
                    "Chart {}",
                    i + 1
                )))],
                ..Default::default()
            }));
            chart_elements.push(Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(text.clone()))],
                ..Default::default()
            }));
        }
        sections.push(Section {
            title: Some("Charts".to_string()),
            elements: chart_elements,
            ..Default::default()
        });
    }

    // Whatever made the workbook incomplete is stated in the content, as
    // the direct renderers state it, not only flagged in the metadata.
    // The title is read before this, so a notice-only section never
    // becomes the document's title.
    let first_title = sections.first().and_then(|s| s.title.clone());
    if !doc.notices().is_empty() {
        let notices = doc.notices().iter().map(|n| {
            Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(n.clone()))],
                ..Default::default()
            })
        });
        match sections.last_mut() {
            Some(last) => last.elements.extend(notices),
            None => sections.push(Section {
                elements: notices.collect(),
                ..Default::default()
            }),
        }
    }

    // The workbook's own declared title (from `\x05SummaryInformation`)
    // beats the first sheet's name — a sheet name is not a document
    // title, it's just the only thing that was ever there to fall back
    // to.
    let summary = doc.summary_properties();
    let title = summary
        .and_then(|s| s.title.clone())
        .filter(|t| !t.is_empty())
        .or(first_title);

    DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Xls,
            title,
            author: summary
                .and_then(|s| s.author.clone())
                .filter(|s| !s.is_empty()),
            subject: summary
                .and_then(|s| s.subject.clone())
                .filter(|s| !s.is_empty()),
            keywords: summary
                .and_then(|s| s.keywords.as_deref())
                .map(crate::convert_docx::split_keywords)
                .unwrap_or_default(),
            description: summary
                .and_then(|s| s.comments.clone())
                .filter(|s| !s.is_empty()),
            created: summary.and_then(|s| s.created.clone()),
            modified: summary.and_then(|s| s.modified.clone()),
            has_macros: doc.has_macros(),
            text_truncated: doc.truncated(),
        },
        sections,
        defined_names: doc
            .defined_names
            .iter()
            .map(|dn| DefinedName {
                name: dn.name.clone(),
                value: dn.value.clone(),
                local_sheet_id: dn.local_sheet_id,
                hidden: dn.hidden,
            })
            .collect(),
    }
}

/// Append extracted BLIP images to the last section of a converted legacy
/// document, or to a new section when there is none.
pub(crate) fn append_legacy_images(
    sections: &mut Vec<Section>,
    images: &[crate::cfb::blip::BlipImage],
) {
    if images.is_empty() {
        return;
    }
    if sections.is_empty() {
        sections.push(Section::default());
    }
    let last = sections.last_mut().expect("just ensured non-empty");
    for img in images {
        last.elements.push(Element::Image(Image {
            data: Some(img.data.clone()),
            format: ImageFormat::from_blip(&img.format),
            ..Default::default()
        }));
    }
}

/// Whether a converted cell carries no text.
fn cell_is_empty(cell: &TableCell) -> bool {
    cell.content.iter().all(|e| match e {
        Element::Paragraph(p) => p.content.iter().all(|c| match c {
            InlineContent::Text(t) => t.text.is_empty(),
            _ => false,
        }),
        _ => false,
    })
}

/// The hyperlink target for each of a row's `width` cells, or `None` when
/// no hyperlink range touches the row.
fn row_hyperlinks<'a>(
    links: &mut RangeSweep<&'a str>,
    row: u32,
    width: usize,
) -> Option<Vec<Option<&'a str>>> {
    if links.is_empty() || width == 0 {
        return None;
    }
    let mut out: Vec<Option<&'a str>> = Vec::new();
    let mut cost = 0u64;
    // Declaration order, so a later record overwrites an earlier one.
    for (range, target) in links.row(row) {
        let Some((lo, hi)) = clip_cols(range, width) else {
            continue;
        };
        if out.is_empty() {
            out = vec![None; width];
        }
        for slot in &mut out[lo..=hi] {
            *slot = Some(*target);
        }
        cost += (hi - lo + 1) as u64;
    }
    links.charge(cost);
    (!out.is_empty()).then_some(out)
}

/// Give each merge anchor in `cells` its span (clipped to the table's
/// `total_rows` x `max_width` extent) and drop every covered position.
fn apply_merges(
    merges: &mut RangeSweep<()>,
    cells: &mut Vec<TableCell>,
    row_idx: usize,
    total_rows: usize,
    max_width: usize,
) {
    let width = cells.len();
    if width == 0 {
        // Still advance the sweep so its row order stays monotonic.
        let _ = merges.row(row_idx as u32).count();
        return;
    }
    let mut covered = vec![false; width];
    let mut cost = 0u64;
    for (range, ()) in merges.row(row_idx as u32) {
        let Some((lo, hi)) = clip_cols(range, width) else {
            continue;
        };
        if range.row_lo as usize == row_idx {
            let anchor = &mut cells[lo];
            let rows_left = (total_rows - row_idx) as u64;
            let cols_left = (max_width - lo) as u64;
            anchor.row_span = range.row_span().min(rows_left) as u32;
            anchor.col_span = range.col_span().min(cols_left) as u32;
            covered[lo + 1..=hi].fill(true);
        } else {
            covered[lo..=hi].fill(true);
        }
        cost += (hi - lo + 1) as u64;
    }
    merges.charge(cost);
    let mut col = 0usize;
    cells.retain(|_| {
        let keep = !covered[col];
        col += 1;
        keep
    });
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xls::{CellValue, Sheet, XlsDocument};

    /// A sheet `rows` tall whose only populated row is `data_row`.
    fn sparse_sheet(rows: usize, cols: usize, data_row: usize, text: &str) -> Sheet {
        let mut grid = vec![vec![CellValue::Empty; cols]; rows];
        grid[data_row][0] = CellValue::String(text.to_string());
        Sheet {
            name: "S".into(),
            rows: grid,
            ..Default::default()
        }
    }

    fn cell_texts(ir: &DocumentIR) -> Vec<String> {
        ir.sections[0]
            .elements
            .iter()
            .flat_map(|el| match el {
                Element::Table(t) => t
                    .rows
                    .iter()
                    .flat_map(|r| r.cells.iter())
                    .flat_map(|c| c.content.iter())
                    .filter_map(|e| match e {
                        Element::Paragraph(p) => Some(p),
                        _ => None,
                    })
                    .flat_map(|p| p.content.iter())
                    .filter_map(|c| match c {
                        InlineContent::Text(t) => Some(t.text.clone()),
                        _ => None,
                    })
                    .collect::<Vec<_>>(),
                Element::Paragraph(p) => p
                    .content
                    .iter()
                    .filter_map(|c| match c {
                        InlineContent::Text(t) => Some(t.text.clone()),
                        _ => None,
                    })
                    .collect(),
                _ => Vec::new(),
            })
            .collect()
    }

    /// XLS half of the merged-cell gap — TableCell::col_span/row_span were
    /// hardcoded to 1 on every cell; the MERGEDCELLS record wasn't even
    /// parsed, so merge information was discarded before it was in
    /// memory, not just dropped at IR conversion.
    #[test]
    fn test_merged_cells_set_col_span_on_the_anchor_and_exclude_covered_cells() {
        let sheet = Sheet {
            name: "S".into(),
            rows: vec![
                vec![
                    CellValue::String("Header".to_string()),
                    CellValue::Empty,
                    CellValue::Empty,
                ],
                vec![
                    CellValue::String("a".to_string()),
                    CellValue::String("b".to_string()),
                    CellValue::String("c".to_string()),
                ],
            ],
            merged_cells: vec![(0, 0, 0, 2)], // row 0, cols 0..=2
            ..Default::default()
        };
        let ir = xls_to_ir(&XlsDocument::from_sheets(vec![sheet]));
        let table = ir.sections[0]
            .elements
            .iter()
            .find_map(|e| match e {
                Element::Table(t) => Some(t),
                _ => None,
            })
            .expect("expected a table element");

        let header_row = &table.rows[0];
        assert_eq!(
            header_row.cells.len(),
            1,
            "the 2 covered cells must be excluded, leaving only the anchor: {:?}",
            header_row.cells
        );
        assert_eq!(header_row.cells[0].col_span, 3);
        assert_eq!(header_row.cells[0].row_span, 1);

        let data_row = &table.rows[1];
        assert_eq!(data_row.cells.len(), 3, "an unmerged row must keep all 3 cells");
    }

    #[test]
    fn test_data_past_the_old_row_limit_survives_in_a_mostly_empty_grid() {
        // A BIFF sheet is padded to its declared used range, so a row limit
        // was measured against padding: a value at row 20,000 of an
        // otherwise-empty 30,000-row grid was dropped by a cap that exists
        // only to bound the padding.
        let ir =
            xls_to_ir(&XlsDocument::from_sheets(vec![sparse_sheet(30_000, 4, 20_000, "deep")]));
        assert!(
            cell_texts(&ir).iter().any(|t| t == "deep"),
            "the one populated row must survive"
        );
    }

    #[test]
    fn test_a_grid_of_padding_emits_no_rows_and_claims_no_truncation() {
        // Every row empty: there is nothing to show and nothing was dropped,
        // so a truncation notice would be a false report. The real files that
        // motivated this are 65,536 x 256; the shape is what matters here.
        let ir = xls_to_ir(&XlsDocument::from_sheets(vec![Sheet {
            name: "S".into(),
            rows: vec![vec![CellValue::Empty; 64]; 2_000],
            ..Default::default()
        }]));
        assert!(ir.sections[0].elements.is_empty());
    }

    /// Regression: the corner-cell shape of a real 42 KB file — a 65,536 x
    /// 256 grid with text only in its four corners. Every empty row used to
    /// allocate a 256-cell buffer, pop the cells, and keep the buffer
    /// (`Vec::pop` does not shrink): 4.9 GB of retained capacity behind a
    /// 6 MB IR. The IR must be small *and* must not hold that capacity.
    #[test]
    fn test_corner_cells_in_a_huge_grid_retain_no_padding_capacity() {
        let (rows, cols) = (4_096, 256);
        let mut grid = vec![vec![CellValue::Empty; cols]; rows];
        grid[0][0] = CellValue::String("Top Left".into());
        grid[0][cols - 1] = CellValue::String("Top Right".into());
        grid[rows - 1][0] = CellValue::String("Bottom Left".into());
        grid[rows - 1][cols - 1] = CellValue::String("Bottom Right".into());
        let ir = xls_to_ir(&XlsDocument::from_sheets(vec![Sheet {
            name: "S".into(),
            rows: grid,
            ..Default::default()
        }]));
        let Element::Table(t) = &ir.sections[0].elements[0] else {
            panic!("expected a table");
        };
        assert_eq!(t.rows.len(), rows);
        assert_eq!(t.rows[0].cells.len(), cols, "the top row keeps its far-right cell");
        assert_eq!(t.rows[rows - 1].cells.len(), cols);
        let retained: usize = t.rows[1..rows - 1].iter().map(|r| r.cells.capacity()).sum();
        assert_eq!(retained, 0, "empty rows must not keep the padding's allocation");
        assert!(cell_texts(&ir).iter().any(|s| s == "Bottom Right"));
    }

    #[test]
    fn test_a_sheet_denser_than_the_budget_is_capped_and_says_so() {
        // The cap still has to exist: an unbounded grid built 16.7M IR cells
        // and ran the process out of memory.
        let rows = MAX_CELLS_PER_SHEET / 100 + 50;
        let grid = vec![vec![CellValue::Number(1.0); 100]; rows];
        let ir = xls_to_ir(&XlsDocument::from_sheets(vec![Sheet {
            name: "S".into(),
            rows: grid,
            ..Default::default()
        }]));
        let notice = cell_texts(&ir)
            .into_iter()
            .find(|t| t.contains("not shown"))
            .expect("a truncation notice");
        assert!(notice.contains(&rows.to_string()), "notice: {notice}");
    }

    /// `Sheet::hyperlinks` (from `HLINK` records) must reach
    /// the cell's own `TextSpan::hyperlink`, the same IR shape
    /// `convert_xlsx.rs` already uses.
    #[test]
    fn test_hyperlink_reaches_the_cells_text_span() {
        let sheet = Sheet {
            name: "S".into(),
            rows: vec![vec![CellValue::String("Stacie@ABC.com".to_string())]],
            hyperlinks: vec![crate::xls::XlsHyperlink {
                row_first: 0,
                row_last: 0,
                col_first: 0,
                col_last: 0,
                target: "mailto:Stacie@ABC.com".to_string(),
            }],
            ..Default::default()
        };
        let ir = xls_to_ir(&XlsDocument::from_sheets(vec![sheet]));
        let Element::Table(t) = &ir.sections[0].elements[0] else {
            panic!("expected a table");
        };
        let Element::Paragraph(p) = &t.rows[0].cells[0].content[0] else {
            panic!("expected a paragraph");
        };
        let InlineContent::Text(span) = &p.content[0] else {
            panic!("expected a text span");
        };
        assert_eq!(span.hyperlink.as_deref(), Some("mailto:Stacie@ABC.com"));
    }

    /// A hyperlink covering a multi-cell range (rare, but the record
    /// format allows it) must apply to every cell in that range, not
    /// just the anchor.
    #[test]
    fn test_hyperlink_range_applies_to_every_covered_cell() {
        let sheet = Sheet {
            name: "S".into(),
            rows: vec![vec![
                CellValue::String("A".to_string()),
                CellValue::String("B".to_string()),
            ]],
            hyperlinks: vec![crate::xls::XlsHyperlink {
                row_first: 0,
                row_last: 0,
                col_first: 0,
                col_last: 1,
                target: "http://example.com".to_string(),
            }],
            ..Default::default()
        };
        let ir = xls_to_ir(&XlsDocument::from_sheets(vec![sheet]));
        let Element::Table(t) = &ir.sections[0].elements[0] else {
            panic!("expected a table");
        };
        for cell in &t.rows[0].cells {
            let Element::Paragraph(p) = &cell.content[0] else {
                panic!("expected a paragraph");
            };
            let InlineContent::Text(span) = &p.content[0] else {
                panic!("expected a text span");
            };
            assert_eq!(span.hyperlink.as_deref(), Some("http://example.com"));
        }
    }

    /// `Sheet::comments` (resolved from NOTE/TXO/OBJ
    /// records) must reach the sheet's elements as endnotes, the same
    /// convention convert_xlsx.rs uses for its own cell comments.
    #[test]
    fn test_comments_reach_the_sheet_as_endnotes() {
        let sheet = Sheet {
            name: "S".into(),
            rows: vec![vec![CellValue::String("data".to_string())]],
            comments: vec![crate::xls::XlsComment {
                row: 0,
                col: 0,
                author: Some("Gilsinei Hansen".to_string()),
                text: "a real cell comment".to_string(),
            }],
            ..Default::default()
        };
        let ir = xls_to_ir(&XlsDocument::from_sheets(vec![sheet]));
        let note = ir.sections[0]
            .elements
            .iter()
            .find_map(|e| match e {
                Element::Endnote(n) => Some(n),
                _ => None,
            })
            .expect("a comment endnote");
        assert_eq!(note.marker.as_deref(), Some("A1 (Gilsinei Hansen)"));
        let Element::Paragraph(p) = &note.content[0] else {
            panic!("expected a paragraph");
        };
        let InlineContent::Text(span) = &p.content[0] else {
            panic!("expected a text span");
        };
        assert_eq!(span.text, "a real cell comment");
        // The IR surfaces name the cell and the author, as the direct
        // renderers do; the body alone used to be all that reached them.
        assert!(
            ir.plain_text()
                .contains("A1 (Gilsinei Hansen): a real cell comment")
        );
        assert!(
            ir.to_markdown()
                .contains("**A1 (Gilsinei Hansen):** a real cell comment")
        );
        assert!(
            ir.to_html()
                .contains("<strong>A1 (Gilsinei Hansen):</strong>")
        );
    }

    /// Run `xls_to_ir` on another thread and fail if it does not finish in
    /// `secs` — a conversion whose cost follows a range's declared area
    /// rather than the cells present never returns.
    fn ir_within(doc: XlsDocument, secs: u64) -> DocumentIR {
        let (tx, rx) = std::sync::mpsc::channel();
        std::thread::spawn(move || {
            let _ = tx.send(xls_to_ir(&doc));
        });
        rx.recv_timeout(std::time::Duration::from_secs(secs))
            .expect("xls_to_ir must finish in time proportional to the cells present")
    }

    fn two_by_two() -> Vec<Vec<CellValue>> {
        vec![
            vec![CellValue::String("a".into()), CellValue::String("b".into())],
            vec![CellValue::String("c".into()), CellValue::String("d".into())],
        ]
    }

    /// A `MERGEDCELLS` entry may legally name the whole BIFF8 grid
    /// (rows 0..=65535, and the raw u16 column fields go past 255).
    /// Expanding it position by position was billions of set inserts, and
    /// the 1-based span overflowed u16.
    #[test]
    fn test_a_full_grid_merge_range_converts_in_bounded_time() {
        let sheet = Sheet {
            name: "S".into(),
            rows: two_by_two(),
            merged_cells: vec![(0, 0xFFFF, 0, 0xFFFF)],
            ..Default::default()
        };
        let ir = ir_within(XlsDocument::from_sheets(vec![sheet]), 20);
        let Element::Table(t) = &ir.sections[0].elements[0] else {
            panic!("expected a table");
        };
        assert_eq!(t.rows[0].cells.len(), 1, "only the anchor survives: {:?}", t.rows[0]);
        assert_eq!(cell_texts(&ir), ["a"]);
        let anchor = &t.rows[0].cells[0];
        assert_eq!(anchor.row_span, 2, "the span is clipped to the rows the table has");
        assert_eq!(anchor.col_span, 2, "the span is clipped to the columns the table has");
    }

    /// An `HLINK` record's range is four raw u16s; a whole-grid one was
    /// expanded into a map entry per position.
    #[test]
    fn test_a_full_grid_hyperlink_range_converts_in_bounded_time() {
        let sheet = Sheet {
            name: "S".into(),
            rows: two_by_two(),
            hyperlinks: vec![crate::xls::XlsHyperlink {
                row_first: 0xFFFF,
                row_last: 0,
                col_first: 0xFFFF,
                col_last: 0,
                target: "http://example.com".to_string(),
            }],
            ..Default::default()
        };
        let ir = ir_within(XlsDocument::from_sheets(vec![sheet]), 20);
        let Element::Table(t) = &ir.sections[0].elements[0] else {
            panic!("expected a table");
        };
        for cell in t.rows.iter().flat_map(|r| &r.cells) {
            let Element::Paragraph(p) = &cell.content[0] else {
                panic!("expected a paragraph");
            };
            let InlineContent::Text(span) = &p.content[0] else {
                panic!("expected a text span");
            };
            assert_eq!(span.hyperlink.as_deref(), Some("http://example.com"));
        }
    }

    /// Overlapping hyperlink ranges: the later record wins, as the map
    /// insert order made it before.
    #[test]
    fn test_overlapping_hyperlink_ranges_resolve_to_the_last_declared() {
        let link = |c0, c1, t: &str| crate::xls::XlsHyperlink {
            row_first: 0,
            row_last: 1,
            col_first: c0,
            col_last: c1,
            target: t.to_string(),
        };
        let sheet = Sheet {
            name: "S".into(),
            rows: two_by_two(),
            hyperlinks: vec![link(0, 1, "first"), link(1, 1, "second")],
            ..Default::default()
        };
        let ir = xls_to_ir(&XlsDocument::from_sheets(vec![sheet]));
        let Element::Table(t) = &ir.sections[0].elements[0] else {
            panic!("expected a table");
        };
        let link_of = |cell: &TableCell| match &cell.content[0] {
            Element::Paragraph(p) => match &p.content[0] {
                InlineContent::Text(s) => s.hyperlink.clone(),
                _ => None,
            },
            _ => None,
        };
        assert_eq!(link_of(&t.rows[1].cells[0]).as_deref(), Some("first"));
        assert_eq!(link_of(&t.rows[1].cells[1]).as_deref(), Some("second"));
    }

    /// Vast numbers of overlapping merge ranges over the same rows: the
    /// work is bounded, and the omission is stated rather than silent.
    #[test]
    fn test_overlapping_merge_ranges_past_the_work_budget_are_reported() {
        let rows: Vec<Vec<CellValue>> = (0..2_000)
            .map(|_| vec![CellValue::String("v".into())])
            .collect();
        let sheet = Sheet {
            name: "S".into(),
            rows,
            merged_cells: vec![(0, 0xFFFF, 0, 1); 40_000],
            ..Default::default()
        };
        let ir = ir_within(XlsDocument::from_sheets(vec![sheet]), 60);
        let texts = cell_texts(&ir);
        assert!(
            texts.iter().any(|t| t.contains("merged")),
            "a notice must record the ignored ranges: {:?}",
            texts.last()
        );
    }
}
