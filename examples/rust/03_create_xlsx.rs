//! 05_xlsx_writer: write a workbook with typed cells, formulas, a merge and a width.
//!
//! Run: `cargo run --example 03_create_xlsx`

use std::io::Cursor;

use office_oxide::xlsx::write::{CellData, CellStyle, XlsxWriter};
use office_oxide::{Document, DocumentFormat};

fn build_xlsx_bytes() -> Vec<u8> {
    let mut writer = XlsxWriter::new();
    {
        let mut sheet = writer.add_sheet("Sales");
        sheet.set_column_width(0, 18.0);
        let bold = CellStyle::new().bold();
        for (col, header) in ["Region", "Q1", "Q2", "Growth"].into_iter().enumerate() {
            sheet.set_cell_styled(0, col, CellData::String(header.into()), bold.clone());
        }
        sheet.set_cell(1, 0, CellData::String("North".into()));
        sheet.set_cell(1, 1, CellData::Number(120.0));
        sheet.set_cell(1, 2, CellData::Number(135.0));
        sheet.set_cell(1, 3, CellData::Boolean(true));
        sheet.set_cell(2, 0, CellData::String("South".into()));
        sheet.set_cell(2, 1, CellData::Number(95.0));
        sheet.set_cell(2, 2, CellData::Number(110.0));
        // A formula plus its cached value: readers that do not evaluate
        // formulas (this one included) then still see the number.
        sheet.set_cell(3, 0, CellData::String("Total".into()));
        for (col, (formula, total)) in [("SUM(B2:B3)", 215.0), ("SUM(C2:C3)", 245.0)]
            .into_iter()
            .enumerate()
        {
            sheet.set_cell(
                3,
                col + 1,
                CellData::FormulaWithValue {
                    formula: formula.into(),
                    cached: Box::new(CellData::Number(total)),
                },
            );
        }
        sheet.set_cell(5, 0, CellData::String("Notes".into()));
        sheet.merge_cells(5, 0, 1, 3); // A6:C6
    }
    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).expect("write XLSX");
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
    println!("step 1: build an XLSX in memory");
    let bytes = build_xlsx_bytes();

    println!("step 2: open it from bytes");
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Xlsx).expect("open XLSX");
    let text = doc.plain_text();
    let sheets = doc.to_ir().sections.len();
    println!("  sheets: {sheets}");
    println!("  chars: {}", text.len());

    println!("step 3: check the cells");
    check(sheets == 1, "the workbook must have one sheet");
    check(
        text.contains("Region") && text.contains("North") && text.contains("South"),
        "xlsx text must contain the labels",
    );
    check(
        text.contains("120") && text.contains("110"),
        "xlsx text must contain the numbers",
    );
    check(
        text.contains("215") && text.contains("245"),
        "xlsx text must contain the formula results",
    );
    check(text.to_lowercase().contains("true"), "xlsx text must contain the boolean");
    check(text.contains("Notes"), "xlsx text must contain the merged cell text");
    println!("OK 05_xlsx_writer");
}
