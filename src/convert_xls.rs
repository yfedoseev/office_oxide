use crate::format::DocumentFormat;
use crate::ir::*;

pub(crate) fn xls_to_ir(doc: &crate::xls::XlsDocument) -> DocumentIR {
    let mut sections = Vec::new();

    for sheet in &doc.sheets {
        let mut rows = Vec::new();

        for (row_idx, row) in sheet.rows.iter().enumerate() {
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

            rows.push(TableRow {
                cells,
                is_header: row_idx == 0,
                ..Default::default()
            });
        }

        let elements = if rows.is_empty() {
            Vec::new()
        } else {
            vec![Element::Table(Table {
                rows,
                ..Default::default()
            })]
        };

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
