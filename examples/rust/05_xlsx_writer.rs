//! 05_xlsx_writer: write a workbook with typed cells, a formula, a style, a merge and a column width, then read it back.
//!
//! Run: `cargo run --example 05_xlsx_writer`

use office_oxide::xlsx::write::{CellData, CellStyle, XlsxWriter};
use office_oxide::{Document, DocumentFormat};
use std::io::Cursor;

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

fn main() {
    println!("step 1: build an XLSX in memory");
    let data = build_xlsx_bytes();

    println!("step 2: open it from bytes");
    let doc = open_bytes(data, DocumentFormat::Xlsx);
    let text = doc.plain_text();
    let sheets = doc.to_ir().sections.len();
    println!("  format: {}", doc.format().extension());
    println!("  chars: {}", text.len());
    println!("  sheets: {sheets}");

    println!("step 3: check the cells");
    check(sheets == 1, "expected one sheet");
    for expected in [
        "Sales", "Region", "Q1", "Q2", "Growth", "North", "120", "135", "South", "95", "110",
        "Total", "Notes",
    ] {
        has(&text, expected);
    }
    check(text.contains("TRUE"), "the boolean cell is missing");
    check(
        text.contains("SUM(B2:B3)") && text.contains("SUM(C2:C3)"),
        "a formula is missing",
    );

    println!("OK 05_xlsx_writer");
}
