//! `markdown` and `markdown --embed-images` (MCP: `markdown` and
//! `markdown-with-images`) come from two pipelines: the per-format direct
//! renderer (`Document::to_markdown`) and the IR renderer
//! (`Document::to_markdown_with`). Nothing tied them together, so the two
//! commands could disagree on non-image content. These tests are that tie.

use std::collections::BTreeMap;

use office_oxide::Document;
use office_oxide::format::DocumentFormat;
use office_oxide::ir::ImageFormat;
use office_oxide::ir_render::{ImageEmbed, MarkdownOptions};

const RICH_MD: &str = "\
# Title

Intro with **bold**, *italic*, `code` and a [link](https://example.com/a?b=1&c=2).

## Second level

- one
  - one-a
- two

1. first
2. second

| Col A | Col B |
|---|---|
| a1 | b1 |

```rust
fn main() {}
```

Tail paragraph.
";

/// A 1×1 PNG.
const PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
    0x00, 0x03, 0x01, 0x01, 0x00, 0x18, 0xDD, 0x8D, 0xB0, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
    0x44, 0xAE, 0x42, 0x60, 0x82,
];

/// Every synthetic document the parity checks run over, with whether it
/// carries an image.
fn corpus(tag: &str) -> Vec<(String, Document, bool)> {
    let dir =
        std::env::temp_dir().join(format!("office_oxide_mdparity_{tag}_{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let mut docs = Vec::new();
    for fmt in [
        DocumentFormat::Docx,
        DocumentFormat::Xlsx,
        DocumentFormat::Pptx,
    ] {
        let path = dir.join(format!("rich.{}", fmt.extension()));
        office_oxide::create::create_from_markdown(RICH_MD, fmt, &path).unwrap();
        docs.push((format!("markdown→{fmt:?}"), Document::open(&path).unwrap(), false));
    }

    let mut d = office_oxide::docx::write::DocxWriter::new();
    d.add_heading("Pictures", 1);
    d.add_paragraph("Before the picture.");
    d.add_ir_image(&office_oxide::ir::Image {
        data: Some(PNG.to_vec()),
        format: Some(ImageFormat::Png),
        display_width_emu: Some(500_000),
        display_height_emu: Some(500_000),
        ..Default::default()
    });
    d.add_paragraph("After the picture.");
    let path = dir.join("image.docx");
    d.save(&path).unwrap();
    docs.push(("docx with image".into(), Document::open(&path).unwrap(), true));

    let mut p = office_oxide::pptx::write::PptxWriter::new();
    {
        let s = p.add_slide();
        s.set_title("Slide with a picture");
        s.add_text("Caption text");
        s.add_image(PNG.to_vec(), ImageFormat::Png, 0, 0, 500_000, 500_000);
    }
    let path = dir.join("image.pptx");
    p.save(&path).unwrap();
    docs.push(("pptx with image".into(), Document::open(&path).unwrap(), true));

    std::fs::remove_dir_all(&dir).ok();
    docs
}

fn ir_markdown(doc: &Document, image_embed: ImageEmbed) -> String {
    doc.to_markdown_with(MarkdownOptions { image_embed })
}

/// Remove every `[image-base64:…]` token, and the blank-line block it sat
/// in when it stood alone.
fn strip_embedded_images(md: &str) -> (String, usize) {
    let mut out = String::with_capacity(md.len());
    let mut rest = md;
    let mut n = 0;
    while let Some(at) = rest.find("[image-base64:") {
        out.push_str(&rest[..at]);
        let end = rest[at..].find(']').map_or(rest.len(), |e| at + e + 1);
        rest = &rest[end..];
        n += 1;
    }
    out.push_str(rest);
    // An image that was a block of its own leaves an empty block behind.
    let mut collapsed = out.replace("\n\n\n\n", "\n\n");
    while collapsed.contains("\n\n\n\n") {
        collapsed = collapsed.replace("\n\n\n\n", "\n\n");
    }
    (collapsed.trim().to_string(), n)
}

/// The visible words of a markdown rendering, as a multiset: link targets
/// and markdown punctuation removed, so layout differences (blank lines,
/// bullets, table pipes) do not count, but a word present in one output and
/// missing from the other does.
fn visible_words(md: &str) -> BTreeMap<String, usize> {
    let mut text = String::with_capacity(md.len());
    let mut rest = md;
    // Drop `](target)` link destinations.
    while let Some(at) = rest.find("](") {
        text.push_str(&rest[..at]);
        let end = rest[at..].find(')').map_or(rest.len(), |e| at + e + 1);
        rest = &rest[end..];
    }
    text.push_str(rest);
    let mut words = BTreeMap::new();
    for w in text
        .split(|c: char| !c.is_alphanumeric())
        .filter(|w| !w.is_empty())
    {
        *words.entry(w.to_string()).or_insert(0) += 1;
    }
    words
}

/// `--embed-images` must add the images and change nothing else: stripped
/// of its `[image-base64:…]` tokens it is the IR markdown without images.
#[test]
fn test_embedding_images_changes_nothing_but_the_images() {
    for (name, doc, has_image) in corpus("embed") {
        let plain = ir_markdown(&doc, ImageEmbed::None);
        let (stripped, images) = strip_embedded_images(&ir_markdown(&doc, ImageEmbed::Base64));
        assert_eq!(images > 0, has_image, "{name}: wrong number of embedded images");
        assert_eq!(stripped, plain.trim(), "{name}: embedding images changed other content");
    }
}

/// `markdown` (direct renderer) and `markdown --embed-images` (IR
/// renderer) must present the same text. Layout may differ between the two
/// renderers; a word one of them drops or invents may not.
#[test]
fn test_direct_and_ir_markdown_carry_the_same_words() {
    for (name, doc, _) in corpus("words") {
        let direct = doc.to_markdown();
        let via_ir = ir_markdown(&doc, ImageEmbed::None);
        assert_eq!(
            visible_words(&direct),
            visible_words(&via_ir),
            "{name}: `markdown` and `markdown --embed-images` disagree on content\n\
             --- direct ---\n{direct}\n--- via IR ---\n{via_ir}"
        );
    }
}
