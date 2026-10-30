//! 03_edit: replace text in a DOCX and set a cell in an XLSX, in memory.
//!
//! Run: `cargo run --example 05_edit`

use std::io::Cursor;

use office_oxide::create::create_from_markdown_to_writer;
use office_oxide::edit::EditableDocument;
use office_oxide::xlsx::edit::CellValue;
use office_oxide::xlsx::write::{CellData, XlsxWriter};
use office_oxide::{Document, DocumentFormat};

const MARKDOWN: &str = "# Quarterly Report\n\nRevenue grew in every region.\n\n\
| Region | Q1  | Q2  |\n|--------|-----|-----|\n| North  | 120 | 135 |\n| South  | 95  | 110 |\n";

fn build_docx_bytes(markdown: &str) -> Vec<u8> {
    let mut buf = Cursor::new(Vec::new());
    create_from_markdown_to_writer(markdown, DocumentFormat::Docx, &mut buf).expect("create DOCX");
    buf.into_inner()
}

/// Assert with a message; a failure exits non-zero.
fn check(condition: bool, message: &str) {
    if !condition {
        eprintln!("FAIL: {message}");
        std::process::exit(1);
    }
}

fn main() {
    println!("step 1: build a DOCX in memory");
    let bytes = build_docx_bytes(MARKDOWN);

    println!("step 2: replace text");
    let mut editable = EditableDocument::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .expect("open editable DOCX");
    let replaced = editable
        .replace_text("Quarterly Report", "Annual Report")
        .expect("replace_text");
    println!("  replacements: {replaced}");
    check(replaced >= 1, "replace_text must report at least one replacement");

    println!("step 3: save to bytes and check the text");
    let mut out = Cursor::new(Vec::new());
    editable.write_to(&mut out).expect("save edited DOCX");
    let doc =
        Document::from_reader(Cursor::new(out.into_inner()), DocumentFormat::Docx).expect("reopen");
    let text = doc.plain_text();
    check(text.contains("Annual Report"), "edited text must contain the new title");
    check(!text.contains("Quarterly Report"), "edited text must not contain the old title");

    println!("step 4: set a cell in an XLSX");
    let mut writer = XlsxWriter::new();
    writer
        .add_sheet("Sales")
        .set_cell(0, 0, CellData::String("Region".into()));
    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).expect("write XLSX");
    let mut sheet =
        EditableDocument::from_reader(Cursor::new(buf.into_inner()), DocumentFormat::Xlsx)
            .expect("open editable XLSX");
    sheet
        .set_cell(0, "B1", CellValue::String("Edited".into()))
        .expect("set_cell");
    let mut out = Cursor::new(Vec::new());
    sheet.write_to(&mut out).expect("save edited XLSX");
    let doc =
        Document::from_reader(Cursor::new(out.into_inner()), DocumentFormat::Xlsx).expect("reopen");
    check(doc.plain_text().contains("Edited"), "edited sheet must contain the new cell");
    println!("OK 03_edit");
}
