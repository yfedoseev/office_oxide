use crate::format::DocumentFormat;
use crate::ir::*;

pub(crate) fn doc_to_ir(doc: &crate::doc::DocDocument) -> DocumentIR {
    let text = doc.plain_text_ref();

    let mut elements: Vec<Element> = Vec::new();

    for line in text.lines() {
        let trimmed = line.trim();
        if trimmed.is_empty() {
            continue;
        }

        // Heuristic: short lines in ALL CAPS or ending with no punctuation
        // at the start of a section are likely headings
        let is_heading = trimmed.len() < 100
            && !trimmed.ends_with('.')
            && !trimmed.ends_with(',')
            && (trimmed
                .chars()
                .filter(|c| c.is_alphabetic())
                .all(|c| c.is_uppercase())
                || (elements.is_empty() && trimmed.len() < 60));

        if is_heading {
            elements.push(Element::Heading(Heading {
                level: if elements.is_empty() { 1 } else { 2 },
                content: vec![InlineContent::Text(TextSpan {
                    bold: true,
                    ..TextSpan::plain(trimmed)
                })],
                ..Default::default()
            }));
        } else {
            elements.push(Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(line))],
                ..Default::default()
            }));
        }
    }

    // Surface embedded images extracted from the Data stream. .doc text
    // extraction discards the picture placeholder (\x01), so the exact inline
    // position is not known here; the images are appended in document (BLIP
    // scan) order so they aren't lost. Emitting Element::Image with the decoded
    // bytes lets the HTML/markdown renderers inline them (see ir_render.rs).
    use crate::cfb::blip::BlipFormat;
    for img in doc.images() {
        let format = match img.format {
            BlipFormat::Png => Some(ImageFormat::Png),
            BlipFormat::Jpeg => Some(ImageFormat::Jpeg),
            BlipFormat::Dib => Some(ImageFormat::Bmp),
            BlipFormat::Tiff => Some(ImageFormat::Tiff),
            BlipFormat::Emf => Some(ImageFormat::Emf),
            BlipFormat::Wmf => Some(ImageFormat::Wmf),
            BlipFormat::Pict | BlipFormat::Unknown(_) => None,
        };
        elements.push(Element::Image(Image {
            data: Some(img.data.clone()),
            format,
            ..Default::default()
        }));
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
