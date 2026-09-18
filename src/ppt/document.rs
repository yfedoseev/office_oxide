//! High-level PPT document API.

use std::io::{Read, Seek};

use crate::cfb::{CfbReader, SummaryProperties, parse_summary_information};

use super::error::Result;
use super::images::{PptImage, extract_images};
use super::text::{SlideText, TextType, extract_slides_text};

/// A parsed legacy PowerPoint document.
#[derive(Debug)]
pub struct PptDocument {
    /// Text content extracted from each slide.
    pub slides: Vec<SlideText>,
    images: Vec<PptImage>,
    has_macros: bool,
    /// Title/author/subject/keywords/comments/dates from the
    /// `\x05SummaryInformation` OLE property-set stream every real `.ppt`
    /// carries by default — parsed and then never read anywhere in the
    /// crate before (issue #244).
    summary_properties: Option<SummaryProperties>,
}

impl PptDocument {
    /// Open a PPT file from a reader.
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let mut cfb = CfbReader::new(reader)?;
        let has_macros = cfb.has_root_entry("_VBA_PROJECT");
        let summary_properties = cfb
            .open_stream("\u{5}SummaryInformation")
            .ok()
            .and_then(|data| parse_summary_information(&data));

        let stream = match cfb
            .open_stream("PowerPoint Document")
            .or_else(|_| cfb.open_stream("PP97_DUALSTORAGE"))
        {
            Ok(s) => s,
            Err(_) => {
                return Ok(Self {
                    slides: Vec::new(),
                    images: Vec::new(),
                    has_macros,
                    summary_properties,
                });
            },
        };

        let current_user = cfb.open_stream("Current User").ok();
        let slides = extract_slides_text(&stream, current_user.as_deref());

        // Extract images from Pictures stream (if present).
        let images = match cfb.open_stream("Pictures") {
            Ok(pictures) => extract_images(&pictures),
            Err(_) => Vec::new(),
        };

        Ok(Self { slides, images, has_macros, summary_properties })
    }

    /// `true` when the file carries a `_VBA_PROJECT` storage — a cheap
    /// macro-presence signal, no VBA interpretation (issue #283).
    pub fn has_macros(&self) -> bool {
        self.has_macros
    }

    /// Title/author/subject/keywords/comments/dates from the file's
    /// `\x05SummaryInformation` OLE property set, when present and
    /// well-formed (issue #244).
    pub fn summary_properties(&self) -> Option<&SummaryProperties> {
        self.summary_properties.as_ref()
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
            has_macros: false,
            summary_properties: None,
            slides: vec![
                SlideText {
                    text_runs: vec![
                        TextRun {
                            text_type: TextType::Title,
                            text: "Welcome".into(),
                            hyperlink: None,
                            ..Default::default()
                        },
                        TextRun {
                            text_type: TextType::Body,
                            text: "Hello world".into(),
                            hyperlink: None,
                            ..Default::default()
                        },
                    ],
                },
                SlideText {
                    text_runs: vec![TextRun {
                        text_type: TextType::Title,
                        text: "Slide 2".into(),
                        hyperlink: None,
                        ..Default::default()
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
            has_macros: false,
            summary_properties: None,
            slides: vec![SlideText {
                text_runs: vec![
                    TextRun {
                        text_type: TextType::Title,
                        text: "My Title".into(),
                        hyperlink: None,
                        ..Default::default()
                    },
                    TextRun {
                        text_type: TextType::Body,
                        text: "Content here".into(),
                        hyperlink: None,
                        ..Default::default()
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
            has_macros: false,
            summary_properties: None,
            slides: vec![SlideText {
                text_runs: vec![
                    TextRun {
                        text_type: TextType::Title,
                        text: "Title".into(),
                        hyperlink: None,
                        ..Default::default()
                    },
                    TextRun {
                        text_type: TextType::Notes,
                        text: "Speaker notes".into(),
                        hyperlink: None,
                        ..Default::default()
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
                    hyperlink: None,
                    ..Default::default()
                })
                .collect(),
        }
    }

    #[test]
    fn ir_empty_doc_has_no_sections() {
        let doc = PptDocument {
            images: Vec::new(),
            slides: Vec::new(),
            has_macros: false,
            summary_properties: None,
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
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![(TextType::Title, "My Slide")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.metadata.title.as_deref(), Some("My Slide"));
        assert!(matches!(ir.sections[0].elements[0], Element::Heading(_)));
    }

    /// issue #253 — a title-slide layout's real title (`textType=6`) and
    /// subtitle (`textType=5`) must not be swapped: the document title
    /// must come from the real title text, and the subtitle must not
    /// become a bold `Heading`.
    #[test]
    fn ir_title_slide_title_and_subtitle_are_not_swapped() {
        use crate::ir::Element;
        let doc = PptDocument {
            images: Vec::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![
                (TextType::from_u32(6), "BSE in the US"), // real title
                (TextType::from_u32(5), "Lisa A. Ferguson, DVM"), // real subtitle
            ])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(
            ir.metadata.title.as_deref(),
            Some("BSE in the US"),
            "the document title must come from the real title (textType=6), not the subtitle"
        );
        assert_eq!(ir.sections[0].title.as_deref(), Some("BSE in the US"));
        // The subtitle (CenterBody) must land as an ordinary Paragraph via
        // convert_ppt.rs's catch-all arm, never as a bold Heading.
        assert!(
            ir.sections[0]
                .elements
                .iter()
                .any(|e| matches!(e, Element::Paragraph(p)
                    if p.content.iter().any(|c| matches!(c, crate::ir::InlineContent::Text(t) if t.text == "Lisa A. Ferguson, DVM")))),
            "the subtitle must reach the IR as an ordinary paragraph"
        );
        assert!(
            !ir.sections[0]
                .elements
                .iter()
                .any(|e| matches!(e, Element::Heading(h)
                    if h.content.iter().any(|c| matches!(c, crate::ir::InlineContent::Text(t) if t.text == "Lisa A. Ferguson, DVM")))),
            "the subtitle must never become the document heading"
        );
    }

    /// issue #257 — a `TextRun::hyperlink` resolved from `InteractiveInfo`
    /// must reach `TextSpan::hyperlink` in the IR, for every text type
    /// that hyperlink can attach to (not just plain body paragraphs).
    #[test]
    fn ir_hyperlink_reaches_textspan_hyperlink() {
        use crate::ir::{Element, InlineContent};
        let doc = PptDocument {
            images: Vec::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![SlideText {
                text_runs: vec![TextRun {
                    text_type: TextType::Body,
                    text: "Click here".to_string(),
                    hyperlink: Some("http://testuri.org/".to_string()),
                    ..Default::default()
                }],
            }],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        let Element::Paragraph(p) = &ir.sections[0].elements[0] else {
            panic!("expected a paragraph, got {:?}", ir.sections[0].elements[0]);
        };
        let InlineContent::Text(span) = &p.content[0] else {
            panic!("expected text content");
        };
        assert_eq!(span.hyperlink.as_deref(), Some("http://testuri.org/"));
    }

    #[test]
    fn ir_center_title_treated_like_title() {
        let doc = PptDocument {
            images: Vec::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![(TextType::CenterTitle, "Centered")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.sections[0].title.as_deref(), Some("Centered"));
    }

    /// The PPTX side of this was already fixed in #203; #238 is the same
    /// defect on the legacy binary .ppt path — `TextType::Notes` runs must
    /// land on `Section.speaker_notes`, not leak into `elements` as
    /// ordinary (if italicized) paragraphs indistinguishable from body
    /// text to most IR consumers.
    #[test]
    fn ir_notes_go_to_speaker_notes_not_elements() {
        use crate::ir::{Element, InlineContent};
        let doc = PptDocument {
            images: Vec::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![
                (TextType::Title, "Title"),
                (TextType::Body, "Visible body text"),
                (TextType::Notes, "Presenter-only notes"),
            ])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(
            ir.sections[0].speaker_notes.as_deref(),
            Some("Presenter-only notes"),
            "notes text must reach Section::speaker_notes"
        );
        let elements_text = format!("{:?}", ir.sections[0].elements);
        assert!(
            !elements_text.contains("Presenter-only notes"),
            "notes text must not also leak into elements: {elements_text}"
        );
        // Visible content is unaffected.
        assert!(
            ir.sections[0]
                .elements
                .iter()
                .any(|e| matches!(e, Element::Paragraph(p) if matches!(&p.content[0], InlineContent::Text(t) if t.text == "Visible body text"))),
            "body text must still reach elements"
        );
    }

    /// A slide with no notes at all gets `None`, not an empty string.
    #[test]
    fn ir_no_notes_is_none_not_empty_string() {
        let doc = PptDocument {
            images: Vec::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![(TextType::Body, "Just body text")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert!(ir.sections[0].speaker_notes.is_none());
    }

    /// Multiple `TextType::Notes` runs on one slide are joined, not just
    /// the last one kept.
    #[test]
    fn ir_multiple_notes_runs_are_joined() {
        let doc = PptDocument {
            images: Vec::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![
                (TextType::Notes, "First note"),
                (TextType::Notes, "Second note"),
            ])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        let notes = ir.sections[0].speaker_notes.as_deref().unwrap_or_default();
        assert!(notes.contains("First note"), "notes: {notes:?}");
        assert!(notes.contains("Second note"), "notes: {notes:?}");
    }

    #[test]
    fn ir_body_half_quarter_produce_paragraphs() {
        use crate::ir::Element;
        let doc = PptDocument {
            images: Vec::new(),
            has_macros: false,
            summary_properties: None,
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
    fn ir_notes_produce_no_visible_elements() {
        // Superseded by #238: notes used to become an italic Element::
        // Paragraph in `elements`; they now route to Section::speaker_notes
        // instead (see ir_notes_go_to_speaker_notes_not_elements above) and
        // a notes-only slide has no visible elements at all.
        let doc = PptDocument {
            images: Vec::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![(TextType::Notes, "Speaker note")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert!(ir.sections[0].elements.is_empty());
        assert_eq!(ir.sections[0].speaker_notes.as_deref(), Some("Speaker note"));
    }

    #[test]
    fn ir_other_text_type_produces_paragraph() {
        use crate::ir::Element;
        let doc = PptDocument {
            images: Vec::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![make_slide(vec![(TextType::Other, "misc text")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert!(matches!(ir.sections[0].elements[0], Element::Paragraph(_)));
    }

    #[test]
    fn ir_slide_without_title_gets_fallback_name() {
        let doc = PptDocument {
            images: Vec::new(),
            has_macros: false,
            summary_properties: None,
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
            has_macros: false,
            summary_properties: None,
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.metadata.format, crate::format::DocumentFormat::Ppt);
    }

    /// issue #244 — `SummaryInformation` fields must reach `Metadata`, and
    /// the declared title must beat the first-slide-title fallback.
    #[test]
    fn ir_summary_properties_reach_metadata() {
        let doc = PptDocument {
            images: Vec::new(),
            has_macros: false,
            summary_properties: Some(crate::cfb::SummaryProperties {
                title: Some("Declared Title".to_string()),
                subject: Some("Declared Subject".to_string()),
                author: Some("Declared Author".to_string()),
                keywords: Some("alpha, beta".to_string()),
                comments: Some("Declared Comment".to_string()),
                created: Some("2020-01-02T03:04:05Z".to_string()),
                modified: Some("2021-06-07T08:09:10Z".to_string()),
            }),
            slides: vec![make_slide(vec![(TextType::Title, "Slide Title")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.metadata.title.as_deref(), Some("Declared Title"));
        assert_eq!(ir.metadata.author.as_deref(), Some("Declared Author"));
        assert_eq!(ir.metadata.subject.as_deref(), Some("Declared Subject"));
        assert_eq!(ir.metadata.keywords, vec!["alpha".to_string(), "beta".to_string()]);
        assert_eq!(ir.metadata.description.as_deref(), Some("Declared Comment"));
        assert_eq!(ir.metadata.created.as_deref(), Some("2020-01-02T03:04:05Z"));
        assert_eq!(ir.metadata.modified.as_deref(), Some("2021-06-07T08:09:10Z"));
    }

    /// A missing/empty title in `SummaryInformation` must not shadow the
    /// first-slide-title fallback.
    #[test]
    fn ir_empty_summary_title_falls_back_to_slide_title() {
        let doc = PptDocument {
            images: Vec::new(),
            has_macros: false,
            summary_properties: Some(crate::cfb::SummaryProperties {
                title: Some(String::new()),
                ..Default::default()
            }),
            slides: vec![make_slide(vec![(TextType::Title, "Slide Title")])],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.metadata.title.as_deref(), Some("Slide Title"));
    }

    // ── #254: direct character/paragraph formatting ──

    /// issue #254 — real `bold: Some(false)` from a `StyleTextPropAtom`
    /// must override the old synthetic "titles are always bold" default,
    /// not just be ignored in its favor.
    #[test]
    fn ir_real_char_formatting_overrides_synthetic_title_bold() {
        use crate::ir::{Element, InlineContent};
        use crate::ppt::{CharFormat, CharFormatSpan};

        let doc = PptDocument {
            images: Vec::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![SlideText {
                text_runs: vec![TextRun {
                    text_type: TextType::Title,
                    text: "Not Bold".to_string(),
                    hyperlink: None,
                    char_formats: vec![CharFormatSpan {
                        start: 0,
                        end: 8,
                        format: CharFormat { bold: Some(false), ..Default::default() },
                    }],
                    para_formats: Vec::new(),
                }],
            }],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        let Element::Heading(h) = &ir.sections[0].elements[0] else {
            panic!("expected a heading, got {:?}", ir.sections[0].elements[0]);
        };
        let InlineContent::Text(span) = &h.content[0] else {
            panic!("expected text content");
        };
        assert!(!span.bold, "explicit bold:false from the file must win over the old synthetic default");
    }

    /// issue #254 — two `TextCFRun`s with different formatting over the
    /// same run of text must produce two separately-formatted `TextSpan`s,
    /// not one span with the first (or last) run's formatting applied to
    /// everything.
    #[test]
    fn ir_char_formatting_produces_multiple_spans_within_one_paragraph() {
        use crate::ir::{Element, InlineContent};
        use crate::ppt::{CharFormat, CharFormatSpan};

        let doc = PptDocument {
            images: Vec::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![SlideText {
                text_runs: vec![TextRun {
                    text_type: TextType::Body,
                    text: "ABCDEF".to_string(),
                    hyperlink: None,
                    char_formats: vec![
                        CharFormatSpan { start: 0, end: 3, format: CharFormat { bold: Some(true), ..Default::default() } },
                        CharFormatSpan { start: 3, end: 6, format: CharFormat { italic: Some(true), ..Default::default() } },
                    ],
                    para_formats: Vec::new(),
                }],
            }],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        let Element::Paragraph(p) = &ir.sections[0].elements[0] else {
            panic!("expected a paragraph, got {:?}", ir.sections[0].elements[0]);
        };
        assert_eq!(p.content.len(), 2, "one span per formatting run: {:?}", p.content);
        let InlineContent::Text(first) = &p.content[0] else { panic!() };
        let InlineContent::Text(second) = &p.content[1] else { panic!() };
        assert_eq!(first.text, "ABC");
        assert!(first.bold);
        assert!(!first.italic);
        assert_eq!(second.text, "DEF");
        assert!(second.italic);
        assert!(!second.bold);
    }

    /// issue #254 — a `TextPFRun`'s alignment must reach `Paragraph::alignment`.
    #[test]
    fn ir_paragraph_alignment_reaches_ir() {
        use crate::ir::{Element, ParagraphAlignment};
        use crate::ppt::{ParaFormat, ParaFormatSpan};

        let doc = PptDocument {
            images: Vec::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![SlideText {
                text_runs: vec![TextRun {
                    text_type: TextType::Body,
                    text: "Centered".to_string(),
                    hyperlink: None,
                    char_formats: Vec::new(),
                    para_formats: vec![ParaFormatSpan {
                        start: 0,
                        end: 8,
                        format: ParaFormat { alignment: Some(1) }, // Tx_ALIGNCenter
                    }],
                }],
            }],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        let Element::Paragraph(p) = &ir.sections[0].elements[0] else {
            panic!("expected a paragraph, got {:?}", ir.sections[0].elements[0]);
        };
        assert_eq!(p.alignment, Some(ParagraphAlignment::Center));
    }

    /// issue #334 — a lone `\r` inside one text atom (the standard PPT97
    /// multi-bullet layout, per [MS-PPT]'s own worked example) must split
    /// into separate `Paragraph` elements, not survive as a literal `\r`
    /// embedded in one giant paragraph (`str::lines()` alone doesn't split
    /// on a bare `\r`).
    #[test]
    fn ir_bare_cr_splits_into_multiple_paragraphs() {
        use crate::ir::{Element, InlineContent};

        let doc = PptDocument {
            images: Vec::new(),
            has_macros: false,
            summary_properties: None,
            slides: vec![SlideText {
                text_runs: vec![TextRun {
                    text_type: TextType::Body,
                    text: "a sunny day\rthe blue sky\rsome green grass".to_string(),
                    hyperlink: None,
                    ..Default::default()
                }],
            }],
        };
        let ir = crate::convert_ppt::ppt_to_ir(&doc);
        assert_eq!(ir.sections[0].elements.len(), 3, "{:?}", ir.sections[0].elements);
        let texts: Vec<&str> = ir.sections[0]
            .elements
            .iter()
            .map(|e| {
                let Element::Paragraph(p) = e else { panic!("expected paragraphs") };
                let InlineContent::Text(t) = &p.content[0] else { panic!() };
                t.text.as_str()
            })
            .collect();
        assert_eq!(texts, ["a sunny day", "the blue sky", "some green grass"]);
        assert!(!texts.iter().any(|t| t.contains('\r')), "no leftover literal \\r: {texts:?}");
    }
}
