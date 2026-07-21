use crate::format::DocumentFormat;
use crate::ir::*;

/// Maps a Data-stream BLIP image to an IR `Element::Image` carrying the decoded
/// bytes (the HTML renderer inlines them as a data: URI; see ir_render.rs).
fn image_element(img: &crate::doc::DocImage) -> Element {
    use crate::cfb::blip::BlipFormat;
    let format = match img.format {
        BlipFormat::Png => Some(ImageFormat::Png),
        BlipFormat::Jpeg => Some(ImageFormat::Jpeg),
        BlipFormat::Dib => Some(ImageFormat::Bmp),
        BlipFormat::Tiff => Some(ImageFormat::Tiff),
        BlipFormat::Emf => Some(ImageFormat::Emf),
        BlipFormat::Wmf => Some(ImageFormat::Wmf),
        BlipFormat::Pict | BlipFormat::Unknown(_) => None,
    };
    Element::Image(Image {
        data: Some(img.data.clone()),
        format,
        positioning: ImagePositioning::Inline,
        ..Default::default()
    })
}

fn text_element(text: &str, is_first: bool) -> Element {
    // Heuristic: short lines in ALL CAPS, or the first line if short, are headings.
    let trimmed = text.trim();
    let is_heading = trimmed.len() < 100
        && !trimmed.ends_with('.')
        && !trimmed.ends_with(',')
        && (trimmed.chars().filter(|c| c.is_alphabetic()).all(|c| c.is_uppercase())
            || (is_first && trimmed.len() < 60));
    if is_heading {
        Element::Heading(Heading {
            level: if is_first { 1 } else { 2 },
            content: vec![InlineContent::Text(TextSpan { bold: true, ..TextSpan::plain(trimmed) })],
            ..Default::default()
        })
    } else {
        Element::Paragraph(Paragraph {
            content: vec![InlineContent::Text(TextSpan::plain(text))],
            ..Default::default()
        })
    }
}

pub(crate) fn doc_to_ir(doc: &crate::doc::DocDocument) -> DocumentIR {
    // plain_text_ref() retains U+FFFC picture markers at the exact character
    // positions of embedded images. We split the text on them and emit an
    // Element::Image at each marker (consuming doc.images() in document order),
    // so images land inline where they actually appear rather than lumped at the
    // end. Any images without a marker (headers/footers, or a count mismatch) are
    // appended so nothing is lost.
    let text = doc.plain_text_ref();
    let images = doc.images();
    let mut img_idx = 0usize;

    let mut elements: Vec<Element> = Vec::new();

    for line in text.lines() {
        if line.contains('\u{FFFC}') {
            for (i, seg) in line.split('\u{FFFC}').enumerate() {
                if i > 0 && img_idx < images.len() {
                    elements.push(image_element(&images[img_idx]));
                    img_idx += 1;
                }
                if !seg.trim().is_empty() {
                    elements.push(text_element(seg, elements.is_empty()));
                }
            }
            continue;
        }

        let trimmed = line.trim();
        if trimmed.is_empty() {
            continue;
        }
        elements.push(text_element(line, elements.is_empty()));
    }

    // Any remaining images had no in-text marker; append them so they aren't lost.
    while img_idx < images.len() {
        elements.push(image_element(&images[img_idx]));
        img_idx += 1;
    }

    let title = elements.iter().find_map(|e| match e {
        Element::Heading(h) => h.content.first().and_then(|c| match c {
            InlineContent::Text(t) => Some(t.text.clone()),
            _ => None,
        }),
        _ => None,
    });

    DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Doc,
            title: title.clone(),
            ..Default::default()
        },
        sections: vec![Section {
            title,
            elements,
            ..Default::default()
        }],
    }
}
