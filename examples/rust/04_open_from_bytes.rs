//! 04_open_from_bytes: build DOCX, XLSX and PPTX in memory and open each from bytes.
//!
//! `Document::from_reader` needs a `Read + Seek + Send + 'static` source: a
//! `Cursor<Vec<u8>>`, not a borrowed slice, because the parse runs on its own thread.
//!
//! Run: `cargo run --example 04_open_from_bytes`

use std::io::Cursor;

use office_oxide::create::create_from_markdown_to_writer;
use office_oxide::ir::ImageFormat;
use office_oxide::pptx::write::PptxWriter;
use office_oxide::xlsx::write::{CellData, CellStyle, XlsxWriter};
use office_oxide::{Document, DocumentFormat};

const MARKDOWN: &str = "# Quarterly Report\n\nRevenue grew in every region.\n\n\
| Region | Q1  | Q2  |\n|--------|-----|-----|\n| North  | 120 | 135 |\n| South  | 95  | 110 |\n";

fn build_docx_bytes(markdown: &str) -> Vec<u8> {
    let mut buf = Cursor::new(Vec::new());
    create_from_markdown_to_writer(markdown, DocumentFormat::Docx, &mut buf).expect("create DOCX");
    buf.into_inner()
}

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

/// The 1x1 PNG from the shared catalogue (its base64 form begins `iVBORw0KGgo`).
const PNG_1X1: [u8; 70] = [
    137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0,
    0, 0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 120, 218, 99, 100, 96, 248, 95, 15, 0, 2,
    135, 1, 128, 235, 71, 186, 146, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
];

fn build_pptx_bytes() -> Vec<u8> {
    let mut writer = PptxWriter::new();
    let slide = writer.add_slide();
    slide.set_title("Welcome");
    slide.add_text("Quarterly Report");
    let slide = writer.add_slide();
    slide.set_title("Results");
    slide.add_text("Revenue grew in every region.");
    // Position and size are in EMU (914400 per inch).
    slide.add_image(PNG_1X1.to_vec(), ImageFormat::Png, 914_400, 2_286_000, 914_400, 914_400);
    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).expect("write PPTX");
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
    println!("step 1: build a DOCX, an XLSX and a PPTX in memory");
    let cases = [
        (DocumentFormat::Docx, build_docx_bytes(MARKDOWN), "Quarterly Report"),
        (DocumentFormat::Xlsx, build_xlsx_bytes(), "North"),
        (DocumentFormat::Pptx, build_pptx_bytes(), "Welcome"),
    ];
    for (format, bytes, expected) in cases {
        println!("step 2: open the {} from bytes", format.extension().to_uppercase());
        let doc = Document::from_reader(Cursor::new(bytes), format).expect("open from bytes");
        let text = doc.plain_text();
        println!("  format: {}", doc.format().extension());
        println!("  chars: {}", text.len());
        println!("step 3: check the text");
        check(doc.format() == format, "the opened format must match the built format");
        check(text.contains(expected), "the text must contain the expected content");
    }
    println!("OK 04_open_from_bytes");
}
