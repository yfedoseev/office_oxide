//! 06_pptx_writer: write a presentation with a size, slides, titles, text and an image, then read it back.
//!
//! Run: `cargo run --example 06_pptx_writer`

use office_oxide::ir::ImageFormat;
use office_oxide::ir_render::{ImageEmbed, MarkdownOptions};
use office_oxide::pptx::write::PptxWriter;
use office_oxide::{Document, DocumentFormat};
use std::io::Cursor;

/// The 1x1 PNG from the catalogue (its base64 form begins `iVBORw0KGgo`).
const PNG_1X1: [u8; 70] = [
    137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0,
    0, 0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 120, 218, 99, 100, 96, 248, 95, 15, 0, 2,
    135, 1, 128, 235, 71, 186, 146, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
];

/// `Document::from_reader` needs a `Read + Seek + Send + 'static` source, so the
/// bytes go in an owned `Cursor`, not a borrowed slice.
fn open_bytes(bytes: Vec<u8>, format: DocumentFormat) -> Document {
    Document::from_reader(Cursor::new(bytes), format).expect("open from bytes")
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
    println!("step 1: build a PPTX in memory");
    let data = build_pptx_bytes();

    println!("step 2: open it from bytes");
    let doc = open_bytes(data, DocumentFormat::Pptx);
    let text = doc.plain_text();
    let slides = doc.to_ir().sections.len();
    // Plain Markdown drops images; the Base64 option keeps them in place.
    let with_images = doc.to_markdown_with(MarkdownOptions {
        image_embed: ImageEmbed::Base64,
    });
    println!("  format: {}", doc.format().extension());
    println!("  chars: {}", text.len());
    println!("  slides: {slides}");

    println!("step 3: check the slides");
    check(slides == 2, "expected two slides");
    for expected in [
        "Welcome",
        "Quarterly Report",
        "Results",
        "Revenue grew in every region.",
    ] {
        has(&text, expected);
    }
    check(text.find("Welcome") < text.find("Results"), "the slides are out of order");
    check(with_images.contains("[image-base64:"), "the slide image is missing");

    println!("OK 06_pptx_writer");
}
