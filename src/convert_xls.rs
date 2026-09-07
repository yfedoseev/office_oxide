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
const MAX_ROWS_PER_SHEET: usize = 10_000;

pub(crate) fn xls_to_ir(doc: &crate::xls::XlsDocument) -> DocumentIR {
    let mut sections = Vec::new();

    for sheet in &doc.sheets {
        let mut rows = Vec::new();
        let total_rows = sheet.rows.len();

        for (row_idx, row) in sheet.rows.iter().take(MAX_ROWS_PER_SHEET).enumerate() {
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

            // Drop trailing empty cells. A BIFF sheet reports the whole
            // declared grid — one file here is 65,536 x 256, essentially all
            // empty — and materialising the padding built 16.7M IR cells and
            // ran the process out of memory. `convert_xlsx` has always
            // trimmed; this path never did.
            while cells.last().is_some_and(|c: &TableCell| cell_is_empty(c)) {
                cells.pop();
            }

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
        if total_rows > MAX_ROWS_PER_SHEET {
            let omitted = total_rows - MAX_ROWS_PER_SHEET;
            elements.push(Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(format!(
                    "[{omitted} of {total_rows} rows not shown — worksheet truncated at \
                     {MAX_ROWS_PER_SHEET} rows]"
                )))],
                ..Default::default()
            }));
        }

        sections.push(Section {
            title: Some(sheet.name.clone()),
            elements,
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
            ..Default::default()
        },
        sections,
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
