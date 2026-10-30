//! 03_edit: replace text in a DOCX and set cells in an XLSX, in memory.
//!
//! Run: `cargo run --example 03_edit`

use office_oxide::create::create_from_markdown_to_writer;
use office_oxide::edit::EditableDocument;
use office_oxide::xlsx::edit::CellValue;
use office_oxide::xlsx::write::{CellData, CellStyle, XlsxWriter};
use office_oxide::{Document, DocumentFormat};
use std::io::Cursor;

const MARKDOWN: &str = "# Quarterly Report\n\nRevenue grew in every region.\n\n\
| Region | Q1  | Q2  |\n|--------|-----|-----|\n| North  | 120 | 135 |\n| South  | 95  | 110 |\n";

fn markdown_to_bytes(markdown: &str, format: DocumentFormat) -> Vec<u8> {
    let mut buf = Cursor::new(Vec::new());
    create_from_markdown_to_writer(markdown, format, &mut buf).expect("create from Markdown");
    buf.into_inner()
}

fn build_docx_bytes(title: &str) -> Vec<u8> {
    markdown_to_bytes(&MARKDOWN.replace("Quarterly Report", title), DocumentFormat::Docx)
}

/// `Document::from_reader` needs a `Read + Seek + Send + 'static` source, so the
/// bytes go in an owned `Cursor`, not a borrowed slice.
fn open_bytes(bytes: Vec<u8>, format: DocumentFormat) -> Document {
    Document::from_reader(Cursor::new(bytes), format).expect("open from bytes")
}

fn build_xlsx_bytes() -> Vec<u8> {
    let mut writer = XlsxWriter::new();
    {
        let mut sheet = writer.add_sheet("Sales");
        sheet.set_column_width(0, 18.0);
        let bold = CellStyle::new().bold();
        for (col, header) in ["Region", "Q1", "Q2", "Growth"].into_iter().enumerate() {
            sheet.set_cell_styled(0, col, CellData::String(header.into()), bold.clone()); // bold header row
        }
        sheet.set_cell(1, 0, CellData::String("North".into()));
        sheet.set_cell(1, 1, CellData::Number(120.0));
        sheet.set_cell(1, 2, CellData::Number(135.0));
        sheet.set_cell(1, 3, CellData::Boolean(true)); // D2
        sheet.set_cell(2, 0, CellData::String("South".into()));
        sheet.set_cell(2, 1, CellData::Number(95.0));
        sheet.set_cell(2, 2, CellData::Number(110.0));
        sheet.set_cell(3, 0, CellData::String("Total".into()));
        sheet.set_cell(3, 1, CellData::Formula("SUM(B2:B3)".into()));
        sheet.set_cell(3, 2, CellData::Formula("SUM(C2:C3)".into()));
        sheet.set_cell(5, 0, CellData::String("Notes".into()));
        sheet.merge_cells(5, 0, 1, 3); // A6:C6
    }
    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).expect("write XLSX");
    buf.into_inner()
}

fn check(condition: bool, message: &str) {
    if !condition {
        eprintln!("FAIL: {message}");
        std::process::exit(1);
    }
}

fn has(text: &str, expected: &str) {
    check(text.contains(expected), &format!("\"{expected}\" is missing from the text"));
}

/// Saves an edited document to bytes and reopens it as a `Document`.
fn reopen(editable: &EditableDocument, format: DocumentFormat) -> Document {
    let mut out = Cursor::new(Vec::new());
    editable.write_to(&mut out).expect("save to bytes");
    open_bytes(out.into_inner(), format)
}

fn main() {
    println!("step 1: build a DOCX in memory");
    let docx = build_docx_bytes("Quarterly Report");

    println!("step 2: replace text");
    let mut editable = EditableDocument::from_reader(Cursor::new(docx), DocumentFormat::Docx)
        .expect("open editable DOCX");
    let replaced = editable
        .replace_text("Quarterly Report", "Annual Report")
        .expect("replace_text");
    println!("  replacements: {replaced}");
    check(replaced == 1, "replace_text did not make exactly one replacement");

    println!("step 3: save to bytes and check the text");
    let text = reopen(&editable, DocumentFormat::Docx).plain_text();
    has(&text, "Annual Report");
    check(!text.contains("Quarterly Report"), "the old title is still present");

    println!("step 4: build an XLSX in memory");
    let xlsx = build_xlsx_bytes();

    println!("step 5: set cells");
    let mut editable = EditableDocument::from_reader(Cursor::new(xlsx), DocumentFormat::Xlsx)
        .expect("open editable XLSX");
    editable
        .set_cell(0, "B2", CellValue::Number(999.0))
        .expect("set B2");
    editable
        .set_cell(0, "A2", CellValue::String("Central".into()))
        .expect("set A2");

    println!("step 6: save to bytes and check the cells");
    let text = reopen(&editable, DocumentFormat::Xlsx).plain_text();
    has(&text, "Central");
    has(&text, "999");
    check(!text.contains("North"), "the old A2 value is still present");

    println!("OK 03_edit");
}
