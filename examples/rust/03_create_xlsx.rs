//! 03_create_xlsx — XLSX creation with formulas, styles, and column widths.
//!
//! Creates a workbook with a header row (bold), data rows, a SUM formula cell
//! with currency style, and explicit column widths. Saves to a temp file and
//! reads back to verify text content.
//!
//! Run: `cargo run --example 03_create_xlsx`

use office_oxide::Document;
use office_oxide::xlsx::write::{CellData, CellStyle, HAlign, NumberFormat, XlsxWriter};

fn main() {
    let out = std::env::temp_dir().join("oo_example_03.xlsx");

    // ── Build the workbook ──────────────────────────────────────────────────
    let mut writer = XlsxWriter::new();
    {
        let mut sheet = writer.add_sheet("Sales Report");

        // Column widths
        sheet.set_column_width(0, 20.0);
        sheet.set_column_width(1, 12.0);
        sheet.set_column_width(2, 10.0);

        // Header row (bold, centred)
        let header = CellStyle::new().bold().align(HAlign::Center);
        sheet.set_cell_styled(0, 0, CellData::String("Product".into()), header.clone());
        sheet.set_cell_styled(0, 1, CellData::String("Price".into()), header.clone());
        sheet.set_cell_styled(0, 2, CellData::String("Quantity".into()), header.clone());
        sheet.set_cell_styled(0, 3, CellData::String("Revenue".into()), header);

        // Data rows
        let rows: &[(&str, f64, f64)] = &[
            ("Widget A", 10.0, 150.0),
            ("Widget B", 25.0, 60.0),
            ("Widget C", 5.0, 500.0),
        ];
        for (i, (name, price, qty)) in rows.iter().enumerate() {
            let row = i + 1;
            sheet.set_cell(row, 0, CellData::String((*name).into()));
            sheet.set_cell_styled(
                row,
                1,
                CellData::Number(*price),
                CellStyle::new().number_format(NumberFormat::Currency),
            );
            sheet.set_cell(row, 2, CellData::Number(*qty));
            // Revenue = Price * Qty (formula referencing row)
            let formula = format!("B{}*C{}", row + 1, row + 1);
            sheet.set_cell_styled(
                row,
                3,
                CellData::Formula(formula),
                CellStyle::new().number_format(NumberFormat::Currency),
            );
        }

        // Totals row
        let total_row = rows.len() + 1;
        sheet.set_cell_styled(
            total_row,
            0,
            CellData::String("TOTAL".into()),
            CellStyle::new().bold(),
        );
        sheet.set_cell_styled(
            total_row,
            3,
            CellData::Formula(format!("SUM(D2:D{})", total_row)),
            CellStyle::new()
                .bold()
                .number_format(NumberFormat::Currency),
        );
    }

    // ── Save ────────────────────────────────────────────────────────────────
    writer.save(&out).expect("save XLSX");

    // ── Read back and verify ────────────────────────────────────────────────
    let doc = Document::open(&out).expect("open saved XLSX");
    let text = doc.plain_text();

    assert!(text.contains("Product"), "header missing");
    assert!(text.contains("Widget A"), "data row missing");
    assert!(text.contains("Widget B"), "data row missing");
    assert!(text.contains("TOTAL"), "total row missing");

    println!("XLSX created and verified: {}", out.display());
    println!("--- plain text ---");
    println!("{text}");
}
