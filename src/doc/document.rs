//! High-level DOC document API.

use std::io::{Read, Seek};

use crate::cfb::CfbReader;

use super::error::{DocError, Result};
use super::fib::Fib;
use super::images::{DocImage, extract_images};
use super::papx::{DocParagraph, build_paragraphs, parse_papx_paragraphs};
use super::piece_table::{extract_text, parse_clx, sanitize_text};

/// A parsed legacy Word document.
#[derive(Debug)]
pub struct DocDocument {
    /// The raw extracted text (after sanitization).
    text: String,
    /// Extracted images from the Data stream.
    images: Vec<DocImage>,
    /// Structured main-text paragraphs with PAP (paragraph property) flags.
    /// Populated only when the FIB advertises a PlcfBtePapx (PAPX FKP index);
    /// empty for very old or minimal files, in which case `doc_to_ir` falls
    /// back to the line-based heuristic on `text`.
    paragraphs: Vec<DocParagraph>,
}

impl DocDocument {
    /// Open a DOC file from a reader.
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let mut cfb = CfbReader::new(reader)?;

        let word_doc = cfb
            .open_stream("WordDocument")
            .map_err(|_| DocError::MissingStream("WordDocument stream not found".into()))?;

        let fib = match Fib::parse(&word_doc) {
            Ok(f) => f,
            Err(_) => {
                return Ok(Self {
                    text: String::new(),
                    images: Vec::new(),
                    paragraphs: Vec::new(),
                });
            }, // Unsupported Word version
        };

        // Open the appropriate table stream; try preferred first, then fallback.
        let table_stream = if fib.use_table1 {
            cfb.open_stream("1Table")
                .or_else(|_| cfb.open_stream("0Table"))
        } else {
            cfb.open_stream("0Table")
                .or_else(|_| cfb.open_stream("1Table"))
        };
        let table_stream = match table_stream {
            Ok(s) => s,
            Err(_) => {
                return Ok(Self {
                    text: String::new(),
                    images: Vec::new(),
                    paragraphs: Vec::new(),
                });
            }, // Word 6/95 or corrupted
        };

        // Extract CLX from the table stream.
        let clx_start = fib.clx_offset as usize;
        let clx_end = clx_start + fib.clx_size as usize;

        if clx_start >= table_stream.len()
            || clx_size_zero_or_oob(fib.clx_size, clx_start, table_stream.len())
        {
            // CLX not available — return empty document.
            return Ok(Self {
                text: String::new(),
                images: Vec::new(),
                paragraphs: Vec::new(),
            });
        }

        let clx_end = clx_end.min(table_stream.len());
        let clx_data = &table_stream[clx_start..clx_end];
        let pieces = match parse_clx(clx_data) {
            Ok(p) => p,
            Err(_) => {
                return Ok(Self {
                    text: String::new(),
                    images: Vec::new(),
                    paragraphs: Vec::new(),
                });
            },
        };

        // Extract main document text only (not footnotes, headers, etc.).
        let raw_text = extract_text(&word_doc, &pieces, fib.text_len);
        let text = sanitize_text(&raw_text);

        // Build structured paragraphs (with table / list PAP flags) from the
        // PAPX FKP, when the FIB advertises one. Without it we cannot detect
        // tables or lists, so `doc_to_ir` falls back to the line heuristic.
        let paragraphs = if fib.fc_plcf_bte_papx != 0 && fib.lcb_plcf_bte_papx != 0 {
            let fkp = parse_papx_paragraphs(
                &word_doc,
                &table_stream,
                fib.fc_plcf_bte_papx,
                fib.lcb_plcf_bte_papx,
            );
            build_paragraphs(&word_doc, &pieces, &fkp, fib.text_len)
        } else {
            Vec::new()
        };

        // Extract images from the Data stream (if present).
        let images = match cfb.open_stream("Data") {
            Ok(data_stream) => extract_images(&data_stream),
            Err(_) => Vec::new(),
        };

        Ok(Self {
            text,
            images,
            paragraphs,
        })
    }

    /// Open a DOC file from a path.
    pub fn open<P: AsRef<std::path::Path>>(path: P) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Get all extracted images.
    pub fn images(&self) -> &[DocImage] {
        &self.images
    }

    /// Get the extracted plain text.
    pub fn plain_text(&self) -> String {
        self.text.clone()
    }

    /// Get a reference to the extracted plain text.
    pub fn plain_text_ref(&self) -> &str {
        &self.text
    }

    /// Structured main-text paragraphs with PAP flags (table / list).
    ///
    /// Empty when the document has no PAPX FKP, in which case callers fall
    /// back to the line-based heuristic on [`Self::plain_text_ref`].
    pub(crate) fn paragraphs(&self) -> &[DocParagraph] {
        &self.paragraphs
    }

    /// Convert to markdown (basic: paragraphs separated by blank lines).
    pub fn to_markdown(&self) -> String {
        let mut result = String::new();
        let mut prev_empty = false;

        for line in self.text.lines() {
            let trimmed = line.trim();
            if trimmed.is_empty() {
                if !prev_empty {
                    result.push('\n');
                }
                prev_empty = true;
            } else {
                result.push_str(trimmed);
                result.push_str("\n\n");
                prev_empty = false;
            }
        }

        result
    }
}

fn clx_size_zero_or_oob(clx_size: u32, clx_start: usize, stream_len: usize) -> bool {
    clx_size == 0 || clx_start + clx_size as usize > stream_len + 1024 // allow some slack
}

impl crate::core::OfficeDocument for DocDocument {
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

    #[test]
    fn markdown_double_spacing() {
        let doc = DocDocument {
            images: Vec::new(),
            text: "First paragraph\nSecond paragraph\n\nAfter gap".into(),
            paragraphs: Vec::new(),
        };
        let md = doc.to_markdown();
        assert!(md.contains("First paragraph\n\n"));
        assert!(md.contains("Second paragraph\n\n"));
        assert!(md.contains("After gap\n\n"));
    }

    #[test]
    fn plain_text_access() {
        let doc = DocDocument {
            images: Vec::new(),
            text: "Hello World".into(),
            paragraphs: Vec::new(),
        };
        assert_eq!(doc.plain_text(), "Hello World");
    }

    fn make_doc(text: &str) -> DocDocument {
        DocDocument {
            images: Vec::new(),
            text: text.to_string(),
            paragraphs: Vec::new(),
        }
    }

    /// Build a `DocDocument` whose IR comes from structured paragraphs
    /// (the PAPX path) rather than the line heuristic. Used to TDD the
    /// table / list walkers without a binary `.doc` fixture.
    fn make_doc_with_paragraphs(paras: Vec<DocParagraph>) -> DocDocument {
        DocDocument {
            images: Vec::new(),
            text: String::new(),
            paragraphs: paras,
        }
    }

    /// Construct a paragraph with the given PAP flags and terminator.
    fn pap(text: &str, props: crate::doc::sprm::PapProps) -> DocParagraph {
        DocParagraph {
            text: text.to_string(),
            terminator: '\r',
            props,
        }
    }

    fn list_props(level: u8) -> crate::doc::sprm::PapProps {
        crate::doc::sprm::PapProps {
            ilvl: Some(level),
            // A real list item carries a valid `ilfo` (0x0001–0x07FE); without
            // it the paragraph is not in a list per [MS-DOC] §2.4.6.3.
            ilfo: Some(1),
            ..Default::default()
        }
    }

    #[test]
    fn ir_list_emits_nested_list_from_ilvl_paragraphs() {
        use crate::ir::Element;
        let doc = make_doc_with_paragraphs(vec![
            pap("Intro.", Default::default()),
            pap("First", list_props(0)),
            pap("Second", list_props(0)),
            pap("Nested", list_props(1)),
            pap("After.", Default::default()),
        ]);
        let ir = crate::convert_doc::doc_to_ir(&doc);
        let elements = &ir.sections[0].elements;

        // [Paragraph, List, Paragraph]
        assert_eq!(elements.len(), 3, "expected intro, list, outro");
        assert!(matches!(elements[0], Element::Paragraph(_)));
        assert!(matches!(elements[2], Element::Paragraph(_)));

        let list = match &elements[1] {
            Element::List(l) => l,
            _ => panic!("expected a List element"),
        };
        assert_eq!(list.items.len(), 2, "two top-level items");
        // Second item nests the level-1 paragraph.
        assert!(list.items[1].nested.is_some(), "second item must nest");
        let nested = list.items[1].nested.as_ref().unwrap();
        assert_eq!(nested.items.len(), 1);
    }

    #[test]
    fn ir_consecutive_list_runs_split_on_prose() {
        use crate::ir::Element;
        let doc = make_doc_with_paragraphs(vec![
            pap("A1", list_props(0)),
            pap("A2", list_props(0)),
            pap("gap", Default::default()),
            pap("B1", list_props(0)),
        ]);
        let ir = crate::convert_doc::doc_to_ir(&doc);
        let elements = &ir.sections[0].elements;
        // [List(A), Paragraph(gap), List(B)]
        let lists: Vec<_> = elements
            .iter()
            .filter(|e| matches!(e, Element::List(_)))
            .collect();
        assert_eq!(lists.len(), 2, "the prose gap must split the run");
    }

    #[test]
    fn ir_empty_doc_produces_empty_section() {
        let ir = crate::convert_doc::doc_to_ir(&make_doc(""));
        assert!(ir.sections[0].elements.is_empty());
        assert!(ir.metadata.title.is_none());
    }

    #[test]
    fn ir_allcaps_first_line_becomes_h1() {
        use crate::ir::Element;
        let ir = crate::convert_doc::doc_to_ir(&make_doc("INTRODUCTION\nSome text here."));
        assert_eq!(ir.metadata.title.as_deref(), Some("INTRODUCTION"));
        assert!(matches!(ir.sections[0].elements[0], Element::Heading(ref h) if h.level == 1));
    }

    #[test]
    fn ir_first_short_line_no_punct_becomes_h1() {
        use crate::ir::Element;
        let ir = crate::convert_doc::doc_to_ir(&make_doc("My Document Title\nThis is body text."));
        assert!(matches!(ir.sections[0].elements[0], Element::Heading(ref h) if h.level == 1));
    }

    #[test]
    fn ir_allcaps_non_first_line_becomes_h2() {
        use crate::ir::Element;
        let ir = crate::convert_doc::doc_to_ir(&make_doc("Title\nSECTION TWO\nBody text."));
        assert!(matches!(ir.sections[0].elements[1], Element::Heading(ref h) if h.level == 2));
    }

    #[test]
    fn ir_line_ending_with_period_becomes_paragraph() {
        use crate::ir::Element;
        let ir = crate::convert_doc::doc_to_ir(&make_doc("This is a sentence."));
        assert!(matches!(ir.sections[0].elements[0], Element::Paragraph(_)));
    }

    #[test]
    fn ir_blank_lines_are_skipped() {
        let ir = crate::convert_doc::doc_to_ir(&make_doc("Title\n\n\nText"));
        assert_eq!(ir.sections[0].elements.len(), 2);
    }

    #[test]
    fn ir_list_run_with_nonzero_base_level_keeps_every_item() {
        // Regression: `.doc` list levels are not guaranteed to start at 0.
        // Word's `simple-list.doc` fixture writes `ilvl = 1` for a flat list,
        // which used to collapse the run to a single item because
        // `build_nested_list` was called with `base_level = 0`.
        use crate::ir::Element;
        let doc = make_doc_with_paragraphs(vec![
            pap("First", list_props(1)),
            pap("Second", list_props(1)),
            pap("Third", list_props(1)),
        ]);
        let ir = crate::convert_doc::doc_to_ir(&doc);
        let elements = &ir.sections[0].elements;
        assert_eq!(elements.len(), 1, "a single list run");
        let list = match &elements[0] {
            Element::List(l) => l,
            _ => panic!("expected a List element"),
        };
        assert_eq!(list.items.len(), 3, "all three items must survive");
    }

    #[test]
    fn ir_format_is_doc() {
        let ir = crate::convert_doc::doc_to_ir(&make_doc("content"));
        assert_eq!(ir.metadata.format, crate::format::DocumentFormat::Doc);
    }
}
