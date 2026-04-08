use crate::format::DocumentFormat;
use crate::ir::*;
use crate::ppt::TextType;

pub(crate) fn ppt_to_ir(doc: &crate::ppt::PptDocument) -> DocumentIR {
    let mut sections = Vec::new();

    for (slide_idx, slide) in doc.slides.iter().enumerate() {
        let mut elements = Vec::new();
        let mut slide_title: Option<String> = None;

        for run in &slide.text_runs {
            let text = run.text.trim();
            if text.is_empty() {
                continue;
            }

            match run.text_type {
                TextType::Title | TextType::CenterTitle => {
                    if slide_title.is_none() {
                        slide_title = Some(text.to_string());
                    }
                    elements.push(Element::Heading(Heading {
                        level: 1,
                        content: vec![InlineContent::Text(TextSpan {
                            text: text.to_string(),
                            bold: true,
                            italic: false,
                            strikethrough: false,
                            hyperlink: None,
                        })],
                    }));
                }
                TextType::Body | TextType::HalfBody | TextType::QuarterBody => {
                    for line in text.lines() {
                        if !line.trim().is_empty() {
                            elements.push(Element::Paragraph(Paragraph {
                                content: vec![InlineContent::Text(TextSpan {
                                    text: line.to_string(),
                                    bold: false,
                                    italic: false,
                                    strikethrough: false,
                                    hyperlink: None,
                                })],
                            }));
                        }
                    }
                }
                TextType::Notes => {
                    // Notes as regular paragraphs
                    for line in text.lines() {
                        if !line.trim().is_empty() {
                            elements.push(Element::Paragraph(Paragraph {
                                content: vec![InlineContent::Text(TextSpan {
                                    text: line.to_string(),
                                    bold: false,
                                    italic: true,
                                    strikethrough: false,
                                    hyperlink: None,
                                })],
                            }));
                        }
                    }
                }
                _ => {
                    elements.push(Element::Paragraph(Paragraph {
                        content: vec![InlineContent::Text(TextSpan {
                            text: text.to_string(),
                            bold: false,
                            italic: false,
                            strikethrough: false,
                            hyperlink: None,
                        })],
                    }));
                }
            }
        }

        let title = slide_title.unwrap_or_else(|| format!("Slide {}", slide_idx + 1));

        sections.push(Section {
            title: Some(title),
            elements,
        });
    }

    let title = sections.first().and_then(|s| s.title.clone());

    DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Ppt,
            title,
        },
        sections,
    }
}
