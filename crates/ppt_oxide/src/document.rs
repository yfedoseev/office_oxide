//! High-level PPT document API.

use std::io::{Read, Seek};

use cfb_oxide::CfbReader;

use crate::error::Result;
use crate::images::{extract_images, PptImage};
use crate::text::{extract_slides_text, SlideText, TextType};

/// A parsed legacy PowerPoint document.
#[derive(Debug)]
pub struct PptDocument {
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
                })
            }
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
                    }
                    TextType::Notes => {
                        // Skip notes in main content.
                    }
                    _ => {
                        out.push_str(&run.text);
                        out.push_str("\n\n");
                    }
                }
            }
        }
        out
    }
}

impl office_core::OfficeDocument for PptDocument {
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
    use crate::text::TextRun;

    #[test]
    fn plain_text_basic() {
        let doc = PptDocument {
            images: Vec::new(),
            slides: vec![
                SlideText {
                    text_runs: vec![
                        TextRun { text_type: TextType::Title, text: "Welcome".into() },
                        TextRun { text_type: TextType::Body, text: "Hello world".into() },
                    ],
                },
                SlideText {
                    text_runs: vec![
                        TextRun { text_type: TextType::Title, text: "Slide 2".into() },
                    ],
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
                    TextRun { text_type: TextType::Title, text: "My Title".into() },
                    TextRun { text_type: TextType::Body, text: "Content here".into() },
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
                    TextRun { text_type: TextType::Title, text: "Title".into() },
                    TextRun { text_type: TextType::Notes, text: "Speaker notes".into() },
                ],
            }],
        };
        let text = doc.plain_text();
        assert!(text.contains("Title"));
        assert!(!text.contains("Speaker notes"));
    }
}
