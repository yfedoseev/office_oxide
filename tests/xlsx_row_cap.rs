//! An XLSX worksheet with more rows than the IR cap must be truncated (with a
//! visible notice) instead of building an unbounded IR that stalls rendering,
//! while ordinary sheets are unaffected.

use office_oxide::ir::{
    Element, InlineContent, Metadata, Paragraph, Section, Table, TableCell, TableRow, TextSpan,
};
use office_oxide::{Document, DocumentFormat, DocumentIR, create};
use std::io::Cursor;

const CAP: usize = 10_000;

fn cell(text: &str) -> TableCell {
    TableCell {
        content: vec![Element::Paragraph(Paragraph {
            content: vec![InlineContent::Text(TextSpan::plain(text.to_string()))],
            ..Default::default()
        })],
        ..Default::default()
    }
}

/// Build an XLSX (via the IR writer) with `n` two-column rows, then read it
/// back to an IR through the capped `xlsx_to_ir` path.
fn roundtrip_sheet(n: usize) -> DocumentIR {
    let rows: Vec<TableRow> = (0..n)
        .map(|i| TableRow {
            cells: vec![cell(&format!("a{i}")), cell(&format!("b{i}"))],
            ..Default::default()
        })
        .collect();
    let ir = DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Xlsx,
            ..Default::default()
        },
        sections: vec![Section {
            title: Some("Sheet1".to_string()),
            elements: vec![Element::Table(Table {
                rows,
                ..Default::default()
            })],
            ..Default::default()
        }],
    };
    let mut buf = Cursor::new(Vec::new());
    create::create_from_ir_to_writer(&ir, DocumentFormat::Xlsx, &mut buf).unwrap();
    buf.set_position(0);
    Document::from_reader(buf, DocumentFormat::Xlsx)
        .unwrap()
        .to_ir()
}

fn table_row_count(ir: &DocumentIR) -> usize {
    ir.sections
        .iter()
        .flat_map(|s| &s.elements)
        .find_map(|e| match e {
            Element::Table(t) => Some(t.rows.len()),
            _ => None,
        })
        .unwrap_or(0)
}

#[test]
fn oversized_sheet_is_capped_with_notice() {
    // 5 rows over the cap. This also completes quickly, demonstrating the sheet
    // no longer builds an unbounded IR.
    let ir = roundtrip_sheet(CAP + 5);
    assert_eq!(table_row_count(&ir), CAP, "table should be capped at {CAP} rows");
    let text = ir
        .sections
        .iter()
        .flat_map(|s| &s.elements)
        .filter_map(|e| match e {
            Element::Paragraph(p) => Some(
                p.content
                    .iter()
                    .filter_map(|c| match c {
                        InlineContent::Text(t) => Some(t.text.as_str()),
                        _ => None,
                    })
                    .collect::<String>(),
            ),
            _ => None,
        })
        .collect::<Vec<_>>()
        .join(" ");
    assert!(
        text.contains("not shown") && text.contains("truncated"),
        "a truncation notice must be present, got: {text:?}"
    );
}

#[test]
fn ordinary_sheet_is_untouched() {
    let ir = roundtrip_sheet(5);
    assert_eq!(table_row_count(&ir), 5, "small sheet must keep all rows");
    let has_notice = ir
        .sections
        .iter()
        .flat_map(|s| &s.elements)
        .any(|e| matches!(e, Element::Paragraph(p) if p.content.iter().any(|c| matches!(c, InlineContent::Text(t) if t.text.contains("not shown")))));
    assert!(!has_notice, "small sheet must not carry a truncation notice");
}
