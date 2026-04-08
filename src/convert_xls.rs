use crate::format::DocumentFormat;
use crate::ir::*;

pub(crate) fn xls_to_ir(doc: &crate::xls::XlsDocument) -> DocumentIR {
    let mut sections = Vec::new();

    for sheet in &doc.sheets {
        let mut rows = Vec::new();

        for (row_idx, row) in sheet.rows.iter().enumerate() {
            let mut cells = Vec::new();
            for cell_value in row {
                let text = cell_value.as_text();
                cells.push(TableCell {
                    content: vec![Element::Paragraph(Paragraph {
                        content: if text.is_empty() {
                            Vec::new()
                        } else {
                            vec![InlineContent::Text(TextSpan {
                                text,
                                bold: false,
                                italic: false,
                                strikethrough: false,
                                hyperlink: None,
                            })]
                        },
                    })],
                    col_span: 1,
                    row_span: 1,
                });
            }

            rows.push(TableRow {
                cells,
                is_header: row_idx == 0,
            });
        }

        let elements = if rows.is_empty() {
            Vec::new()
        } else {
            vec![Element::Table(Table { rows })]
        };

        sections.push(Section {
            title: Some(sheet.name.clone()),
            elements,
        });
    }

    let title = sections.first().and_then(|s| s.title.clone());

    DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Xls,
            title,
        },
        sections,
    }
}
