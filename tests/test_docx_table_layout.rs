//! Table geometry and styling in written DOCX: a table with no known widths
//! must still carry a non-zero grid (Word sizes columns from `w:gridCol`),
//! and a table built from Markdown must be readable — borders, a distinct
//! header row, and a header row that repeats across pages.

use std::io::{Cursor, Read};

use office_oxide::create::create_from_ir_to_writer;
use office_oxide::ir::*;
use office_oxide::{Document, DocumentFormat};

fn write_docx(ir: &DocumentIR) -> Vec<u8> {
    let mut buf = Cursor::new(Vec::new());
    create_from_ir_to_writer(ir, DocumentFormat::Docx, &mut buf).expect("write");
    buf.into_inner()
}

fn document_xml(bytes: &[u8]) -> String {
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut s = String::new();
    zip.by_name("word/document.xml")
        .unwrap()
        .read_to_string(&mut s)
        .unwrap();
    s
}

fn md_docx(md: &str) -> String {
    document_xml(&write_docx(&DocumentIR::from_markdown(md, DocumentFormat::Docx)))
}

/// Every `w:w` value of the `w:gridCol` elements in `xml`.
fn grid_cols(xml: &str) -> Vec<u32> {
    let mut out = Vec::new();
    let mut rest = xml;
    while let Some(at) = rest.find("<w:gridCol ") {
        rest = &rest[at..];
        let end = rest.find("/>").unwrap();
        let tag = &rest[..end];
        let v = tag
            .split("w:w=\"")
            .nth(1)
            .unwrap()
            .split('"')
            .next()
            .unwrap();
        out.push(v.parse().unwrap());
        rest = &rest[end..];
    }
    out
}

const MD: &str = "| Model | Memory | Power | Interface |\n|---|---|---|---|\n\
                  | H100 SXM | 80 GB HBM3 | 700 W | SXM5 |\n\
                  | H100 PCIe | 80 GB HBM2e | 350 W | PCIe 5 |\n";

#[test]
fn test_unsized_table_has_no_zero_width_grid_column() {
    let xml = md_docx(MD);
    let cols = grid_cols(&xml);
    assert_eq!(cols.len(), 4, "{xml}");
    assert!(cols.iter().all(|&w| w > 0), "zero gridCol: {cols:?}");
    assert!(cols.iter().all(|&w| w == cols[0]), "unsized columns split equally: {cols:?}");
}

#[test]
fn test_unsized_table_width_is_fifty_hundredths_percent_of_the_page() {
    let xml = md_docx(MD);
    assert!(xml.contains("<w:tblW w:w=\"5000\" w:type=\"pct\"/>"), "{xml}");
    assert!(!xml.contains("w:w=\"0\""), "a zero width survives: {xml}");
}

#[test]
fn test_explicit_column_widths_are_written_unchanged() {
    let mut ir = DocumentIR::from_markdown(MD, DocumentFormat::Docx);
    let Element::Table(t) = &mut ir.sections[0].elements[0] else {
        panic!("expected table");
    };
    t.column_widths_twips = vec![1000, 2000, 3000, 4000];
    t.width_twips = Some(10000);
    let xml = document_xml(&write_docx(&ir));
    assert_eq!(grid_cols(&xml), vec![1000, 2000, 3000, 4000]);
    assert!(xml.contains("<w:tblW w:w=\"10000\" w:type=\"dxa\"/>"), "{xml}");
}

#[test]
fn test_explicit_table_width_without_columns_is_split_equally() {
    let mut ir = DocumentIR::from_markdown(MD, DocumentFormat::Docx);
    let Element::Table(t) = &mut ir.sections[0].elements[0] else {
        panic!("expected table");
    };
    t.width_twips = Some(8000);
    let cols = grid_cols(&document_xml(&write_docx(&ir)));
    assert_eq!(cols, vec![2000; 4]);
}

#[test]
fn test_ragged_rows_size_the_grid_from_the_widest_row() {
    let xml = md_docx("| a | b |\n|---|---|\n| 1 | 2 | 3 | 4 | 5 |\n");
    let cols = grid_cols(&xml);
    assert_eq!(cols.len(), 5);
    assert!(cols.iter().all(|&w| w > 0), "{cols:?}");
}

#[test]
fn test_a_hundred_column_table_keeps_every_column_nonzero() {
    let head = format!("|{}\n", " h |".repeat(100));
    let sep = format!("|{}\n", "---|".repeat(100));
    let row = format!("|{}\n", " v |".repeat(100));
    let cols = grid_cols(&md_docx(&format!("{head}{sep}{row}")));
    assert_eq!(cols.len(), 100);
    assert!(cols.iter().all(|&w| w >= 1), "{cols:?}");
}

#[test]
fn test_legacy_builder_table_has_nonzero_grid() {
    let mut doc = office_oxide::docx::write::DocxWriter::new();
    doc.add_table(&[vec!["a", "b"], vec!["c", "d"]]);
    let mut buf = Cursor::new(Vec::new());
    doc.write_to(&mut buf).unwrap();
    let cols = grid_cols(&document_xml(&buf.into_inner()));
    assert_eq!(cols.len(), 2);
    assert!(cols.iter().all(|&w| w > 0), "{cols:?}");
}

// ---- markdown table styling ------------------------------------------------

#[test]
fn test_markdown_table_has_borders_on_every_edge_and_between_cells() {
    let xml = md_docx(MD);
    let b = xml
        .split("<w:tblBorders>")
        .nth(1)
        .expect("tblBorders")
        .split("</w:tblBorders>")
        .next()
        .unwrap();
    for side in ["top", "left", "bottom", "right", "insideH", "insideV"] {
        assert!(b.contains(&format!("<w:{side} w:val=\"single\"")), "{side}: {b}");
    }
}

#[test]
fn test_markdown_header_row_repeats_on_each_page() {
    let xml = md_docx(MD);
    assert_eq!(xml.matches("<w:tblHeader/>").count(), 1, "{xml}");
    let first_row = xml
        .split("<w:tr>")
        .nth(1)
        .unwrap()
        .split("</w:tr>")
        .next()
        .unwrap();
    assert!(first_row.contains("<w:tblHeader/>"), "{first_row}");
}

#[test]
fn test_markdown_header_cells_get_a_thick_bottom_rule_and_keep_their_text_plain() {
    let xml = md_docx(MD);
    let rows: Vec<&str> = xml.split("<w:tr>").skip(1).collect();
    assert_eq!(rows.len(), 3);
    let head = rows[0];
    assert_eq!(head.matches("<w:tc>").count(), 4);
    assert!(!head.contains("<w:b/>"), "header text must stay plain: {head}");
    assert_eq!(
        head.matches("<w:tcBorders><w:bottom w:val=\"single\" w:sz=\"12\"")
            .count(),
        4,
        "{head}"
    );
    for body in &rows[1..] {
        assert!(!body.contains("<w:b/>"), "body row is bold: {body}");
        assert!(!body.contains("w:tcBorders"), "{body}");
    }
}

#[test]
fn test_an_explicit_header_style_in_the_ir_is_not_overridden_by_the_writer() {
    // Only tables from Markdown get the default look; a table the caller
    // built (or read from a document) is written as given.
    let ir = DocumentIR {
        metadata: Metadata::default(),
        sections: vec![Section {
            title: None,
            elements: vec![Element::Table(Table {
                rows: vec![TableRow {
                    cells: vec![TableCell {
                        content: vec![Element::Paragraph(Paragraph {
                            content: vec![InlineContent::Text(TextSpan::plain("h"))],
                            ..Default::default()
                        })],
                        col_span: 1,
                        row_span: 1,
                        ..Default::default()
                    }],
                    is_header: true,
                    ..Default::default()
                }],
                ..Default::default()
            })],
            ..Default::default()
        }],
    };
    let xml = document_xml(&write_docx(&ir));
    assert!(!xml.contains("w:tblBorders"), "{xml}");
    assert!(!xml.contains("<w:tblHeader/>"), "{xml}");
    assert!(!xml.contains("<w:b/>"), "{xml}");
}

#[test]
fn test_styled_markdown_table_reads_back_with_same_text_and_header() {
    let bytes = write_docx(&DocumentIR::from_markdown(MD, DocumentFormat::Docx));
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx).unwrap();
    let ir = doc.to_ir();
    let Element::Table(t) = &ir.sections[0].elements[0] else {
        panic!("expected table");
    };
    assert_eq!(t.rows.len(), 3);
    assert!(t.rows[0].repeat_as_header);
    assert!(!t.rows[1].repeat_as_header);
    assert_eq!(t.width_twips, None, "pct width must not read back as twips");
    let md = doc.to_markdown();
    assert!(md.contains("H100 SXM"), "{md}");
    assert!(md.contains("80 GB HBM3"), "{md}");
}

#[test]
fn test_markdown_table_round_trips_to_the_same_markdown() {
    let md = DocumentIR::from_markdown(MD, DocumentFormat::Docx);
    let bytes = write_docx(&md);
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx).unwrap();
    assert_eq!(
        doc.to_markdown().trim_end(),
        "| Model | Memory | Power | Interface |\n| --- | --- | --- | --- |\n\
         | H100 SXM | 80 GB HBM3 | 700 W | SXM5 |\n| H100 PCIe | 80 GB HBM2e | 350 W | PCIe 5 |"
    );
    let text = doc.plain_text();
    for cell in ["Model", "Memory", "Power", "Interface", "700 W", "PCIe 5"] {
        assert!(text.contains(cell), "{cell} missing from {text}");
    }
}

#[test]
fn test_an_all_zero_explicit_grid_is_treated_as_unsized() {
    // A grid read from a document whose widths could not be parsed is all
    // zeros; writing it back would reproduce the collapsed-column table.
    let mut ir = DocumentIR::from_markdown(MD, DocumentFormat::Docx);
    let Element::Table(t) = &mut ir.sections[0].elements[0] else {
        panic!("expected table");
    };
    t.column_widths_twips = vec![0, 0, 0, 0];
    let xml = document_xml(&write_docx(&ir));
    let cols = grid_cols(&xml);
    assert_eq!(cols.len(), 4);
    assert!(cols.iter().all(|&w| w > 0), "{cols:?}");
    assert!(xml.contains("<w:tblW w:w=\"5000\" w:type=\"pct\"/>"), "{xml}");
}
