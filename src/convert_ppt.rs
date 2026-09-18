use crate::format::DocumentFormat;
use crate::ir::*;
use crate::ppt::TextType;

pub(crate) fn ppt_to_ir(doc: &crate::ppt::PptDocument) -> DocumentIR {
    let mut sections = Vec::new();

    for (slide_idx, slide) in doc.slides.iter().enumerate() {
        let mut elements = Vec::new();
        let mut slide_title: Option<String> = None;
        // Presenter-only text, kept out of `elements` (which is "what the
        // audience sees") and routed to the dedicated field instead — the
        // PPTX side of this was already fixed in #203; #238 is the same
        // defect on the legacy binary .ppt path, which had never been
        // ported to route TextType::Notes there at all.
        let mut notes_lines: Vec<&str> = Vec::new();

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
                            bold: true,
                            ..TextSpan::plain(text)
                        })],
                        ..Default::default()
                    }));
                },
                TextType::Body | TextType::HalfBody | TextType::QuarterBody => {
                    for line in text.lines() {
                        if !line.trim().is_empty() {
                            elements.push(Element::Paragraph(Paragraph {
                                content: vec![InlineContent::Text(TextSpan::plain(line))],
                                ..Default::default()
                            }));
                        }
                    }
                },
                TextType::Notes => {
                    notes_lines.push(text);
                },
                _ => {
                    elements.push(Element::Paragraph(Paragraph {
                        content: vec![InlineContent::Text(TextSpan::plain(text))],
                        ..Default::default()
                    }));
                },
            }
        }

        let title = slide_title.unwrap_or_else(|| format!("Slide {}", slide_idx + 1));
        let speaker_notes = if notes_lines.is_empty() {
            None
        } else {
            Some(notes_lines.join("\n").trim().to_string()).filter(|s| !s.is_empty())
        };

        sections.push(Section {
            title: Some(title),
            elements,
            speaker_notes,
            ..Default::default()
        });
    }

    // Extracted pictures never reached the IR, so every image in a legacy
    // deck was silently dropped on conversion.
    crate::convert_xls::append_legacy_images(&mut sections, doc.images());

    // The deck's own declared title (from `\x05SummaryInformation`) beats
    // the first slide's own title — a slide title is not a document
    // title, it's just the only thing that was ever there to fall back
    // to (issue #244).
    let summary = doc.summary_properties();
    let title = summary
        .and_then(|s| s.title.clone())
        .filter(|t| !t.is_empty())
        .or_else(|| sections.first().and_then(|s| s.title.clone()));

    DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Ppt,
            title,
            author: summary.and_then(|s| s.author.clone()).filter(|s| !s.is_empty()),
            subject: summary.and_then(|s| s.subject.clone()).filter(|s| !s.is_empty()),
            keywords: summary
                .and_then(|s| s.keywords.as_deref())
                .map(crate::convert_docx::split_keywords)
                .unwrap_or_default(),
            description: summary.and_then(|s| s.comments.clone()).filter(|s| !s.is_empty()),
            created: summary.and_then(|s| s.created.clone()),
            modified: summary.and_then(|s| s.modified.clone()),
            has_macros: doc.has_macros(),
            ..Default::default()
        },
        sections,
        defined_names: Vec::new(),
    }
}
