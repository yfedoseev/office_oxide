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

/// Reduce `Sheet::merged_cells` ((row_first, row_last, col_first,
/// col_last) tuples from the MERGEDCELLS record) to an anchor
/// (row, col) -> (row_span, col_span) map plus the set of positions
/// each range covers — the same split `convert_xlsx.rs` uses, so both
/// formats feed the sparse, span-driven TableRow model
/// ir_render.rs's table_grid expects (issue #235, XLS half).
fn merge_lookup(
    merged_cells: &[(u16, u16, u16, u16)],
) -> (
    std::collections::HashMap<(u16, u16), (u16, u16)>,
    std::collections::HashSet<(u16, u16)>,
) {
    let mut span = std::collections::HashMap::new();
    let mut covered = std::collections::HashSet::new();
    for &(row_first, row_last, col_first, col_last) in merged_cells {
        let (row_lo, row_hi) = (row_first.min(row_last), row_first.max(row_last));
        let (col_lo, col_hi) = (col_first.min(col_last), col_first.max(col_last));
        let row_span = row_hi - row_lo + 1;
        let col_span = col_hi - col_lo + 1;
        if row_span <= 1 && col_span <= 1 {
            continue;
        }
        span.insert((row_lo, col_lo), (row_span, col_span));
        for r in row_lo..=row_hi {
            for c in col_lo..=col_hi {
                if (r, c) != (row_lo, col_lo) {
                    covered.insert((r, c));
                }
            }
        }
    }
    (span, covered)
}

pub(crate) fn xls_to_ir(doc: &crate::xls::XlsDocument) -> DocumentIR {
    let mut sections = Vec::new();

    for sheet in &doc.sheets {
        let (merge_span, merge_covered) = merge_lookup(&sheet.merged_cells);
        let mut rows = Vec::new();
        // Rows past the last one carrying data are padding; measuring the
        // sheet against them would report a truncation that dropped nothing.
        let total_rows = sheet
            .rows
            .iter()
            .enumerate()
            .rposition(|(row_idx, row)| {
                row.iter().enumerate().any(|(col_idx, cell_value)| {
                    // Matches the emptiness test the emit loop applies, but
                    // without rendering every cell of the declared grid.
                    let has_display = sheet
                        .display
                        .get(row_idx)
                        .and_then(|r| r.get(col_idx))
                        .is_some_and(|s| !s.is_empty());
                    has_display || !matches!(cell_value, crate::xls::CellValue::Empty)
                })
            })
            .map_or(0, |i| i + 1);
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
            let mut cells = Vec::new();
            for (col_idx, cell_value) in row.iter().enumerate() {
                // `display` carries the number-format-aware rendering: a
                // date cell is an ISO date rather than its raw serial.
                // Fall back to the raw rendering when the sheet had no
                // format tables.
                let text = sheet
                    .display
                    .get(row_idx)
                    .and_then(|r| r.get(col_idx))
                    .filter(|s| !s.is_empty())
                    .cloned()
                    .unwrap_or_else(|| cell_value.as_text());
                cells.push(TableCell {
                    content: vec![Element::Paragraph(Paragraph {
                        content: if text.is_empty() {
                            Vec::new()
                        } else {
                            vec![InlineContent::Text(TextSpan::plain(text))]
                        },
                        ..Default::default()
                    })],
                    col_span: 1,
                    row_span: 1,
                    ..Default::default()
                });
            }

            // Apply merges: the anchor gets its real span, and every
            // position it covers is dropped from the row entirely —
            // `sheet.rows[row_idx]` is already a dense, fully-padded
            // grid (`row.iter().enumerate()` gives every column 0..N),
            // so `row_idx`/`col_idx` are already the true absolute
            // positions `merged_cells` uses, no gap-adjustment needed.
            if !merge_span.is_empty() {
                for (col_idx, cell) in cells.iter_mut().enumerate() {
                    if let Some(&(row_span, col_span)) =
                        merge_span.get(&(row_idx as u16, col_idx as u16))
                    {
                        cell.row_span = row_span as u32;
                        cell.col_span = col_span as u32;
                    }
                }
            }
            if !merge_covered.is_empty() {
                let mut col_idx = 0u16;
                cells.retain(|_| {
                    let keep = !merge_covered.contains(&(row_idx as u16, col_idx));
                    col_idx += 1;
                    keep
                });
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

        sections.push(Section {
            title: Some(sheet.name.clone()),
            elements,
            // A hidden sheet is kept and flagged, not dropped — the same
            // contract `convert_xlsx` already honours.
            hidden: sheet.hidden,
            ..Default::default()
        });
    }

    // Extracted pictures never reached the IR, so every image in a legacy
    // workbook was silently dropped on conversion. Append them to the last
    // section; BIFF drawings carry no reliable per-sheet anchor here.
    append_legacy_images(&mut sections, doc.images());

    let title = sections.first().and_then(|s| s.title.clone());

    DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Xls,
            title,
            has_macros: doc.has_macros(),
            text_truncated: doc.truncated(),
            ..Default::default()
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
            display: Vec::new(),
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

    /// issue #235 (XLS half) — TableCell::col_span/row_span were
    /// hardcoded to 1 on every cell; the MERGEDCELLS record wasn't even
    /// parsed, so merge information was discarded before it was in
    /// memory, not just dropped at IR conversion.
    #[test]
    fn merged_cells_set_col_span_on_the_anchor_and_exclude_covered_cells() {
        let sheet = Sheet {
            name: "S".into(),
            display: Vec::new(),
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
    fn data_past_the_old_row_limit_survives_in_a_mostly_empty_grid() {
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
    fn a_grid_of_padding_emits_no_rows_and_claims_no_truncation() {
        // Every row empty: there is nothing to show and nothing was dropped,
        // so a truncation notice would be a false report. The real files that
        // motivated this are 65,536 x 256; the shape is what matters here.
        let ir = xls_to_ir(&XlsDocument::from_sheets(vec![Sheet {
            name: "S".into(),
            display: Vec::new(),
            rows: vec![vec![CellValue::Empty; 64]; 2_000],
            ..Default::default()
        }]));
        assert!(ir.sections[0].elements.is_empty());
    }

    #[test]
    fn a_sheet_denser_than_the_budget_is_capped_and_says_so() {
        // The cap still has to exist: an unbounded grid built 16.7M IR cells
        // and ran the process out of memory.
        let rows = MAX_CELLS_PER_SHEET / 100 + 50;
        let grid = vec![vec![CellValue::Number(1.0); 100]; rows];
        let ir = xls_to_ir(&XlsDocument::from_sheets(vec![Sheet {
            name: "S".into(),
            display: Vec::new(),
            rows: grid,
            ..Default::default()
        }]));
        let notice = cell_texts(&ir)
            .into_iter()
            .find(|t| t.contains("not shown"))
            .expect("a truncation notice");
        assert!(notice.contains(&rows.to_string()), "notice: {notice}");
    }
}
