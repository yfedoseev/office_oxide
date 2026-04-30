use crate::format::DocumentFormat;
use crate::ir::*;

pub(crate) fn xlsx_to_ir(doc: &crate::xlsx::XlsxDocument) -> DocumentIR {
    let mut sections = Vec::new();

    for ws in &doc.worksheets {
        let mut rows = Vec::new();

        for (row_idx, row) in ws.rows.iter().enumerate() {
            let mut cells = Vec::new();
            for cell in &row.cells {
                let text = doc.format_cell_value(cell);
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
            vec![Element::Table(Table { rows, ..Default::default() })]
        };

        sections.push(Section {
            title: Some(ws.name.clone()),
            elements,
            ..Default::default()
        });
    }

    let title = sections.first().and_then(|s| s.title.clone());

    DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Xlsx,
            title,
            ..Default::default()
        },
        sections,
    }
}
