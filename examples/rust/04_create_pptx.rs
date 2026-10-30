//! 06_pptx_writer: write a two-slide presentation with titles, text and an image.
//!
//! Run: `cargo run --example 04_create_pptx`

use std::io::Cursor;

use office_oxide::ir::ImageFormat;
use office_oxide::pptx::write::PptxWriter;
use office_oxide::{Document, DocumentFormat};

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
    println!("step 1: build a PPTX in memory");
    let bytes = build_pptx_bytes();
    // The picture is stored as its own part in the package.
    check(
        bytes.windows(10).any(|w| w == b"ppt/media/"),
        "the pptx package must contain the image",
    );

    println!("step 2: open it from bytes");
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Pptx).expect("open PPTX");
    let text = doc.plain_text();
    let slides = doc.to_ir().sections.len();
    println!("  slides: {slides}");
    println!("  chars: {}", text.len());

    println!("step 3: check the text");
    check(slides == 2, "the presentation must have two slides");
    check(
        text.contains("Welcome") && text.contains("Quarterly Report"),
        "pptx text must contain slide 1",
    );
    check(
        text.contains("Results") && text.contains("Revenue grew in every region."),
        "pptx text must contain slide 2",
    );
    println!("OK 06_pptx_writer");
}
