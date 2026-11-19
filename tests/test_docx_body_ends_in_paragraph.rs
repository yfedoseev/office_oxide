//! A written DOCX body always ends with a paragraph, as Word writes it. A
//! body whose last block was a `w:tbl` sent LibreOffice's DOCX import into an
//! endless loop repeating the table's last cell, so a headless conversion of
//! such a file never finished.

use std::io::{Cursor, Read};

use office_oxide::docx::write::DocxWriter;
use office_oxide::ir::{
    Element, InlineContent, PageSetup, Paragraph, SectionBreakType, Table, TableCell, TableRow,
    TextSpan,
};

fn document_xml(doc: &DocxWriter) -> String {
    let mut buf = Cursor::new(Vec::new());
    doc.write_to(&mut buf).expect("write");
    let mut zip = zip::ZipArchive::new(Cursor::new(buf.into_inner())).unwrap();
    let mut s = String::new();
    zip.by_name("word/document.xml")
        .unwrap()
        .read_to_string(&mut s)
        .unwrap();
    s
}

/// What the body holds between its last table and its end: the body-level
/// `w:sectPr`, or `</w:body>` when there is none.
fn after_last_table(xml: &str) -> &str {
    let table_end = xml.rfind("</w:tbl>").expect("a table") + "</w:tbl>".len();
    let end = xml[table_end..]
        .find("<w:sectPr")
        .or_else(|| xml[table_end..].find("</w:body>"))
        .expect("a body end");
    &xml[table_end..table_end + end]
}

fn cell(text: &str) -> TableCell {
    TableCell {
        content: vec![Element::Paragraph(Paragraph {
            content: vec![InlineContent::Text(TextSpan {
                text: text.into(),
                ..Default::default()
            })],
            ..Default::default()
        })],
        col_span: 1,
        row_span: 1,
        ..Default::default()
    }
}

#[test]
fn test_body_ending_in_a_simple_table_closes_with_a_paragraph() {
    let mut doc = DocxWriter::new();
    doc.add_paragraph("Before");
    doc.add_table(&[vec!["a", "b"]]);
    let xml = document_xml(&doc);
    assert!(after_last_table(&xml).contains("<w:p"), "{xml}");
}

#[test]
fn test_body_ending_in_an_ir_table_closes_with_a_paragraph() {
    let mut doc = DocxWriter::new();
    doc.add_ir_table(&Table {
        rows: vec![TableRow {
            cells: vec![cell("a")],
            ..Default::default()
        }],
        ..Default::default()
    });
    let xml = document_xml(&doc);
    assert!(after_last_table(&xml).contains("<w:p"), "{xml}");
}

#[test]
fn test_closing_paragraph_precedes_the_body_section_properties() {
    let mut doc = DocxWriter::new();
    doc.add_table(&[vec!["a"]]);
    doc.set_section_props(Some(PageSetup::default()), None, SectionBreakType::default());
    let xml = document_xml(&doc);
    assert!(xml.contains("<w:sectPr"), "{xml}");
    assert!(after_last_table(&xml).contains("<w:p"), "{xml}");
}

#[test]
fn test_body_ending_in_a_paragraph_gains_no_closing_paragraph() {
    let mut doc = DocxWriter::new();
    doc.add_table(&[vec!["a"]]);
    doc.add_paragraph("After");
    let xml = document_xml(&doc);
    let after = after_last_table(&xml);
    assert_eq!(after.matches("<w:p").count(), 1, "{xml}");
    assert!(after.contains("After"), "{xml}");
}
