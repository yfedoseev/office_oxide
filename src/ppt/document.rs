//! High-level PPT document API.

use std::io::{Read, Seek};

use crate::cfb::CfbReader;

use super::error::Result;
use super::images::{PptImage, extract_images};
use super::text::{SlideText, TextType, extract_slides_text};

/// A parsed legacy PowerPoint document.
#[derive(Debug)]
pub struct PptDocument {
    /// Text content extracted from each slide.
    pub slides: Vec<SlideText>,
    images: Vec<PptImage>,
}

impl PptDocument {
    /// Open a PPT file from a reader.
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let mut cfb = CfbReader::new(reader)?;

        let stream = match cfb
            .open_stream("PowerPoint Document")
            .or_else(|_| cfb.open_stream("PP97_DUALSTORAGE"))
        {
            Ok(s) => s,
            Err(_) => {
                return Ok(Self {
                    slides: Vec::new(),
                    images: Vec::new(),
                });
            },
        };

        let slides = extract_slides_text(&stream);

        // Extract images from Pictures stream (if present).
        let images = match cfb.open_stream("Pictures") {
            Ok(pictures) => extract_images(&pictures),
            Err(_) => Vec::new(),
        };

        Ok(Self { slides, images })
    }

    /// Open a PPT file from a path.
    pub fn open<P: AsRef<std::path::Path>>(path: P) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Get all extracted images.
    pub fn images(&self) -> &[PptImage] {
        &self.images
    }

    /// Extract plain text.
    pub fn plain_text(&self) -> String {
        let mut out = String::new();
        for (i, slide) in self.slides.iter().enumerate() {
            if i > 0 {
                out.push('\n');
            }
            for run in &slide.text_runs {
                if run.text_type != TextType::Notes {
                    out.push_str(&run.text);
                    out.push('\n');
                }
            }
        }
        out
    }

    /// Convert to markdown.
    pub fn to_markdown(&self) -> String {
        let mut out = String::new();
        for (i, slide) in self.slides.iter().enumerate() {
            if i > 0 {
                out.push('\n');
            }
            out.push_str(&format!("## Slide {}\n\n", i + 1));

            for run in &slide.text_runs {
                match run.text_type {
                    TextType::Title | TextType::CenterTitle => {
                        out.push_str("### ");
                        out.push_str(&run.text);
                        out.push_str("\n\n");
                    },
                    TextType::Notes => {
                        // Skip notes in main content.
                    },
                    _ => {
                        out.push_str(&run.text);
                        out.push_str("\n\n");
                    },
                }
            }
        }
        out
    }
}

impl crate::core::OfficeDocument for PptDocument {
    fn plain_text(&self) -> String {
        self.plain_text()
    }

    fn to_markdown(&self) -> String {
        self.to_markdown()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ppt::text::TextRun;

    #[test]
    fn plain_text_basic() {
        let doc = PptDocument {
            images: Vec::new(),
            slides: vec![
                SlideText {
                    text_runs: vec![
                        TextRun {
                            text_type: TextType::Title,
                            text: "Welcome".into(),
                        },
                        TextRun {
                            text_type: TextType::Body,
                            text: "Hello world".into(),
                        },
                    ],
                },
                SlideText {
                    text_runs: vec![TextRun {
                        text_type: TextType::Title,
                        text: "Slide 2".into(),
                    }],
                },
            ],
        };
        let text = doc.plain_text();
        assert!(text.contains("Welcome"));
        assert!(text.contains("Hello world"));
        assert!(text.contains("Slide 2"));
    }

    #[test]
    fn markdown_basic() {
        let doc = PptDocument {
            images: Vec::new(),
            slides: vec![SlideText {
                text_runs: vec![
                    TextRun {
                        text_type: TextType::Title,
                        text: "My Title".into(),
                    },
                    TextRun {
                        text_type: TextType::Body,
                        text: "Content here".into(),
                    },
                ],
            }],
        };
        let md = doc.to_markdown();
        assert!(md.contains("## Slide 1"));
        assert!(md.contains("### My Title"));
        assert!(md.contains("Content here"));
    }

    #[test]
    fn notes_excluded_from_plain_text() {
        let doc = PptDocument {
            images: Vec::new(),
            slides: vec![SlideText {
                text_runs: vec![
                    TextRun {
                        text_type: TextType::Title,
                        text: "Title".into(),
                    },
                    TextRun {
                        text_type: TextType::Notes,
                        text: "Speaker notes".into(),
                    },
                ],
            }],
        };
        let text = doc.plain_text();
        assert!(text.contains("Title"));
        assert!(!text.contains("Speaker notes"));
    }

    fn make_slide(runs: Vec<(TextType, &str)>) -> SlideText {
        SlideText {
            text_runs: runs
                .into_iter()
                .map(|(t, s)| TextRun {
                    text_type: t,
                    text: s.to_string(),
                })
                .collect(),
        }
    }

    #[test]
    fn ir_empty_doc_has_no_sections() {
        let doc = PptDocument {
            images: Vec::new(),
            slides: Vec::new(),
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert!(ir.sections.is_empty());
        assert!(ir.metadata.title.is_none());
    }

    #[test]
    fn ir_title_becomes_heading_and_section_title() {
        use crate::ir::Element;
        let doc = PptDocument {
            images: Vec::new(),
            slides: vec![make_slide(vec![(TextType::Title, "My Slide")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.metadata.title.as_deref(), Some("My Slide"));
        assert!(matches!(ir.sections[0].elements[0], Element::Heading(_)));
    }

    #[test]
    fn ir_center_title_treated_like_title() {
        let doc = PptDocument {
            images: Vec::new(),
            slides: vec![make_slide(vec![(TextType::CenterTitle, "Centered")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.sections[0].title.as_deref(), Some("Centered"));
    }

    #[test]
    fn ir_body_half_quarter_produce_paragraphs() {
        use crate::ir::Element;
        let doc = PptDocument {
            images: Vec::new(),
            slides: vec![make_slide(vec![
                (TextType::Body, "Body text"),
                (TextType::HalfBody, "Half body"),
                (TextType::QuarterBody, "Quarter"),
            ])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.sections[0].elements.len(), 3);
        assert!(matches!(ir.sections[0].elements[0], Element::Paragraph(_)));
    }

    #[test]
    fn ir_notes_produce_italic_paragraphs() {
        use crate::ir::{Element, InlineContent};
        let doc = PptDocument {
            images: Vec::new(),
            slides: vec![make_slide(vec![(TextType::Notes, "Speaker note")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        if let Element::Paragraph(ref p) = ir.sections[0].elements[0] {
            if let InlineContent::Text(ref span) = p.content[0] {
                assert!(span.italic);
            } else {
                panic!("expected text span");
            }
        } else {
            panic!("expected paragraph");
        }
    }

    #[test]
    fn ir_other_text_type_produces_paragraph() {
        use crate::ir::Element;
        let doc = PptDocument {
            images: Vec::new(),
            slides: vec![make_slide(vec![(TextType::Other, "misc text")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert!(matches!(ir.sections[0].elements[0], Element::Paragraph(_)));
    }

    #[test]
    fn ir_slide_without_title_gets_fallback_name() {
        let doc = PptDocument {
            images: Vec::new(),
            slides: vec![make_slide(vec![(TextType::Body, "content")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.sections[0].title.as_deref(), Some("Slide 1"));
    }

    #[test]
    fn ir_format_is_ppt() {
        let doc = PptDocument {
            images: Vec::new(),
            slides: Vec::new(),
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.metadata.format, crate::format::DocumentFormat::Ppt);
    }
}
