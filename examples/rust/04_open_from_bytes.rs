//! 04_open_from_bytes: build a DOCX, an XLSX and a PPTX in memory and open each from its bytes.
//!
//! Run: `cargo run --example 04_open_from_bytes`

use office_oxide::create::create_from_markdown_to_writer;
use office_oxide::ir::ImageFormat;
use office_oxide::pptx::write::PptxWriter;
use office_oxide::xlsx::write::{CellData, CellStyle, XlsxWriter};
use office_oxide::{Document, DocumentFormat};
use std::io::Cursor;

const MARKDOWN: &str = "# Quarterly Report\n\nRevenue grew in every region.\n\n\
| Region | Q1  | Q2  |\n|--------|-----|-----|\n| North  | 120 | 135 |\n| South  | 95  | 110 |\n";

/// The 1x1 PNG from the catalogue (its base64 form begins `iVBORw0KGgo`).
const PNG_1X1: [u8; 70] = [
    137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0,
    0, 0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 120, 218, 99, 100, 96, 248, 95, 15, 0, 2,
    135, 1, 128, 235, 71, 186, 146, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
];

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

fn build_pptx_bytes() -> Vec<u8> {
    let mut writer = PptxWriter::new();
    writer.set_presentation_size(12_192_000, 6_858_000); // 16:9, in EMU (914400 per inch)
    let slide = writer.add_slide();
    slide.set_title("Welcome");
    slide.add_text("Quarterly Report");
    let slide = writer.add_slide();
    slide.set_title("Results");
    slide.add_text("Revenue grew in every region.");
    slide.add_image(PNG_1X1.to_vec(), ImageFormat::Png, 914_400, 2_743_200, 1_828_800, 1_828_800);
    let mut buf = Cursor::new(Vec::new());
    writer.write_to(&mut buf).expect("write PPTX");
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
    println!("step 1: build a DOCX, an XLSX and a PPTX in memory");
    let built = [
        (DocumentFormat::Docx, build_docx_bytes("Quarterly Report")),
        (DocumentFormat::Xlsx, build_xlsx_bytes()),
        (DocumentFormat::Pptx, build_pptx_bytes()),
    ];

    println!("step 2: open each from bytes");
    let mut texts = Vec::new();
    for (format, data) in &built {
        let doc = open_bytes(data.clone(), *format);
        check(doc.format() == *format, "the opened format differs from the requested one");
        let text = doc.plain_text();
        println!("  format: {}", format.extension());
        println!("  chars: {}", text.len());
        texts.push(text);
    }

    println!("step 3: check the text");
    has(&texts[0], "Revenue grew in every region.");
    has(&texts[1], "North");
    has(&texts[1], "Total");
    has(&texts[2], "Welcome");
    has(&texts[2], "Results");

    println!("step 4: check that a wrong format is an error");
    let wrong = Document::from_reader(Cursor::new(built[1].1.clone()), DocumentFormat::Pptx).err();
    check(wrong.is_some(), "an XLSX opened as a PPTX was accepted");
    println!("  error: {}", wrong.expect("checked above"));

    println!("OK 04_open_from_bytes");
}
