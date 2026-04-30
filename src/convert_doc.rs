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
            }));
        } else {
            elements.push(Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(line))],
                ..Default::default()
            }));
        }
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
        sections: vec![Section { title, elements, ..Default::default() }],
    }
}
