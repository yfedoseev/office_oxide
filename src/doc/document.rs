//! High-level DOC document API.

use std::io::{Read, Seek};

use crate::cfb::{CfbReader, SummaryProperties, parse_summary_information};

use super::error::{DocError, Result};
use super::fib::Fib;
use super::images::{DocImage, extract_images};
use super::list_format::ListFormatting;
use super::papx::{DocParagraph, build_paragraphs, parse_papx_paragraphs};
use super::piece_table::{covers_declared_length, extract_text, parse_clx, sanitize_text};

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
    /// Text of the subdocuments that follow the main document in the piece
    /// table's character space: footnotes, headers/footers, comments,
    /// endnotes and text boxes. The FIB's `ccp*` lengths that delimit them
    /// were parsed and then never used, so none of this reached any
    /// consumer.
    subdocuments: Vec<SubDocument>,
    /// `true` when the CFB container has a top-level `_VBA_PROJECT`
    /// storage — a cheap macro-presence signal, no VBA interpretation
    /// (issue #283).
    has_macros: bool,
    /// `false` when the piece table has a gap before the FIB's declared
    /// `ccpText` — text in that gap is silently absent from `plain_text()`/
    /// `paragraphs()` with no other signal, so a caller who cares can at
    /// least tell "genuinely short document" apart from "84% missing"
    /// (issue #230).
    text_complete: bool,
    /// Title/author/subject/keywords/comments/dates from the
    /// `\x05SummaryInformation` OLE property-set stream every real
    /// `.doc` carries by default — parsed and then never read anywhere
    /// in the crate before (issue #244).
    summary_properties: Option<SummaryProperties>,
    /// `PlfLst`/`PlfLfo` list definitions, resolving a paragraph's
    /// `(ilfo, ilvl)` to its declared start-at value and number format —
    /// parsed from a FIB pointer that was previously read and then never
    /// used anywhere in the crate (issue #250).
    list_formatting: ListFormatting,
    /// Comment author names, parsed from `GrpXstAtnOwners`. The FIB
    /// fields locating this array (`fcGrpXstAtnOwners`/
    /// `lcbGrpXstAtnOwners`) were never parsed at all before, so every
    /// `.doc` comment's authorship was unrecoverable (issue #298).
    comment_authors: Vec<String>,
}

/// One of the subdocuments stored after the main text in a `.doc`.
#[derive(Debug, Clone)]
pub struct SubDocument {
    /// Which subdocument this is.
    pub kind: SubDocumentKind,
    /// Sanitised text of the subdocument.
    pub text: String,
}

/// The subdocument kinds `.doc` stores after the main text, in the fixed
/// order [MS-DOC] §2.5.1 defines for the `ccp*` lengths.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SubDocumentKind {
    /// Footnote bodies (`ccpFtn`).
    Footnotes,
    /// Header and footer bodies (`ccpHdd`).
    HeadersFooters,
    /// Comment bodies (`ccpAtn`).
    Comments,
    /// Endnote bodies (`ccpEdn`).
    Endnotes,
    /// Text-box bodies (`ccpTxbx`).
    TextBoxes,
    /// Header text-box bodies (`ccpHdrTxbx`).
    HeaderTextBoxes,
}

impl DocDocument {
    /// Open a DOC file from a reader.
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let mut cfb = CfbReader::new(reader)?;

        let word_doc = cfb
            .open_stream("WordDocument")
            .map_err(|_| DocError::MissingStream("WordDocument stream not found".into()))?;

        // Propagate FIB errors. Swallowing them into an empty document with
        // `Ok` is what made an encrypted file and a Word 6.0/95 file both
        // look like documents that simply had no text.
        let fib = Fib::parse(&word_doc)?;

        // Open the appropriate table stream; try preferred first, then fallback.
        let table_stream = if fib.use_table1 {
            cfb.open_stream("1Table")
                .or_else(|_| cfb.open_stream("0Table"))
        } else {
            cfb.open_stream("0Table")
                .or_else(|_| cfb.open_stream("1Table"))
        };
        // Each of the three failures below used to return an empty document
        // with `Ok`, which is indistinguishable from a document that has no
        // text. A file that cannot be read must say so.
        let table_stream = table_stream.map_err(|_| {
            DocError::MissingStream("neither 0Table nor 1Table stream is present".into())
        })?;

        // Extract CLX from the table stream.
        let clx_start = fib.clx_offset as usize;
        let clx_end = clx_start + fib.clx_size as usize;

        if clx_start >= table_stream.len()
            || clx_size_zero_or_oob(fib.clx_size, clx_start, table_stream.len())
        {
            return Err(DocError::InvalidPieceTable(format!(
                "CLX at offset {clx_start} size {} is outside the {}-byte table stream",
                fib.clx_size,
                table_stream.len()
            )));
        }

        let clx_end = clx_end.min(table_stream.len());
        let clx_data = &table_stream[clx_start..clx_end];
        let pieces = parse_clx(clx_data)?;
        let text_complete = covers_declared_length(&pieces, fib.text_len);

        let raw_text = extract_text(&word_doc, &pieces, fib.text_len, fib.lid);
        let text = sanitize_text(&raw_text);

        // The subdocuments follow the main text contiguously in the piece
        // table's character space, each delimited by its own `ccp*` length.
        let mut subdocuments = Vec::new();
        let mut cp = fib.text_len;
        for (kind, len) in [
            (SubDocumentKind::Footnotes, fib.footnote_len),
            (SubDocumentKind::HeadersFooters, fib.header_len),
            (SubDocumentKind::Comments, fib.comment_len),
            (SubDocumentKind::Endnotes, fib.endnote_len),
            (SubDocumentKind::TextBoxes, fib.textbox_len),
            (SubDocumentKind::HeaderTextBoxes, fib.header_textbox_len),
        ] {
            if len == 0 {
                continue;
            }
            let end = cp.saturating_add(len);
            let raw =
                super::piece_table::extract_text_range(&word_doc, &pieces, cp, end, fib.lid);
            let sub = sanitize_text(&raw);
            if !sub.trim().is_empty() {
                subdocuments.push(SubDocument { kind, text: sub });
            }
            cp = end;
        }

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
            build_paragraphs(&word_doc, &pieces, &fkp, fib.text_len, fib.lid)
        } else {
            Vec::new()
        };

        let list_formatting = ListFormatting::parse(
            &table_stream,
            fib.fc_plcf_lst,
            fib.lcb_plcf_lst,
            fib.fc_plf_lfo,
            fib.lcb_plf_lfo,
        );

        // Extract images from the Data stream (if present).
        let images = match cfb.open_stream("Data") {
            Ok(data_stream) => extract_images(&data_stream),
            Err(_) => Vec::new(),
        };
        let has_macros = cfb.has_root_entry("_VBA_PROJECT");
        let summary_properties = cfb
            .open_stream("\u{5}SummaryInformation")
            .ok()
            .and_then(|data| parse_summary_information(&data));

        let comment_authors = parse_grp_xst_atn_owners(
            &table_stream,
            fib.fc_grp_xst_atn_owners,
            fib.lcb_grp_xst_atn_owners,
        );

        Ok(Self {
            text,
            images,
            paragraphs,
            subdocuments,
            has_macros,
            text_complete,
            summary_properties,
            list_formatting,
            comment_authors,
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

    /// Footnotes, headers, comments, endnotes and text boxes, in the order
    /// the file stores them.
    pub fn subdocuments(&self) -> &[SubDocument] {
        &self.subdocuments
    }

    /// Comment author names declared by `GrpXstAtnOwners`, in file order.
    /// Empty when the document has no comments. There is no per-comment
    /// author correlation here (that needs `PlcfAtn`/`ATRD`, tracked
    /// separately) — when this holds exactly one name, every comment in
    /// the document was written by that single author (issue #298).
    pub(crate) fn comment_authors(&self) -> &[String] {
        &self.comment_authors
    }

    /// `true` when the file carries a `_VBA_PROJECT` storage — a cheap
    /// macro-presence signal, no VBA interpretation (issue #283).
    pub fn has_macros(&self) -> bool {
        self.has_macros
    }

    /// `false` when the piece table has a gap before the FIB's declared
    /// text length, meaning `plain_text()`/`paragraphs()` are missing real
    /// content that could not be safely recovered (issue #230).
    pub fn text_complete(&self) -> bool {
        self.text_complete
    }

    /// Title/author/subject/keywords/comments/dates from the file's
    /// `\x05SummaryInformation` OLE property set, when present and
    /// well-formed (issue #244).
    pub fn summary_properties(&self) -> Option<&crate::cfb::SummaryProperties> {
        self.summary_properties.as_ref()
    }

    /// `PlfLst`/`PlfLfo` list definitions, resolving a paragraph's
    /// `(ilfo, ilvl)` to its declared start-at value and number format
    /// (issue #250).
    pub(crate) fn list_formatting(&self) -> &ListFormatting {
        &self.list_formatting
    }

    /// Get the extracted plain text.
    ///
    /// Includes footnote/endnote/comment/textbox bodies — `to_ir()` (via
    /// `doc_to_ir`) already carries this content as its own elements, and
    /// leaving it out here made this renderer disagree with that one, the
    /// same gap already fixed for DOCX in #240 (issue #248).
    pub fn plain_text(&self) -> String {
        let mut out = self.text.clone();
        for sub in &self.subdocuments {
            let text = sub.text.trim();
            if text.is_empty() {
                continue;
            }
            if !out.is_empty() && !out.ends_with('\n') {
                out.push('\n');
            }
            out.push_str(text);
            out.push('\n');
        }
        out
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
    ///
    /// Includes footnote/endnote/comment/textbox bodies — see the
    /// identical note on `plain_text()` (issue #248).
    pub fn to_markdown(&self) -> String {
        let mut result = text_to_markdown_blocks(&self.text);
        for sub in &self.subdocuments {
            let block = text_to_markdown_blocks(&sub.text);
            let block = block.trim();
            if block.is_empty() {
                continue;
            }
            if !result.ends_with("\n\n") && !result.is_empty() {
                result.push('\n');
            }
            result.push_str(block);
            result.push_str("\n\n");
        }
        result
    }
}

/// Split `text` into markdown paragraphs, each separated by a blank line —
/// the shared rendering shape `to_markdown()` uses for both the main text
/// and every subdocument body.
fn text_to_markdown_blocks(text: &str) -> String {
    let mut result = String::new();
    let mut prev_empty = false;

    for line in text.lines() {
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

fn clx_size_zero_or_oob(clx_size: u32, clx_start: usize, stream_len: usize) -> bool {
    clx_size == 0 || clx_start + clx_size as usize > stream_len + 1024 // allow some slack
}

/// Parse `GrpXstAtnOwners`: an array of XSTs (comment author names) packed
/// back-to-back at `fc` for `lcb` bytes in the Table stream. Each entry is
/// a `u16` character count `cch` followed by `cch` UTF-16LE code units —
/// no STTBF-style count/extra-data header, per [MS-DOC] §2.5.5's
/// description of `fcGrpXstAtnOwners`. Issue #298.
fn parse_grp_xst_atn_owners(table_stream: &[u8], fc: u32, lcb: u32) -> Vec<String> {
    if lcb == 0 {
        return Vec::new();
    }
    let start = fc as usize;
    let end = start.saturating_add(lcb as usize).min(table_stream.len());
    if start >= end {
        return Vec::new();
    }

    let mut names = Vec::new();
    let mut pos = start;
    while pos + 2 <= end {
        let cch = u16::from_le_bytes([table_stream[pos], table_stream[pos + 1]]) as usize;
        pos += 2;
        let byte_len = cch * 2;
        if pos + byte_len > end {
            break;
        }
        let units: Vec<u16> =
            table_stream[pos..pos + byte_len].chunks_exact(2).map(|c| u16::from_le_bytes([c[0], c[1]])).collect();
        pos += byte_len;
        names.push(String::from_utf16_lossy(&units));
    }
    names
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
            subdocuments: Vec::new(),
            has_macros: false,
            text_complete: true,
            summary_properties: None,
            list_formatting: crate::doc::list_format::ListFormatting::default(),
            comment_authors: Vec::new(),
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
            subdocuments: Vec::new(),
            has_macros: false,
            text_complete: true,
            summary_properties: None,
            list_formatting: crate::doc::list_format::ListFormatting::default(),
            comment_authors: Vec::new(),
            images: Vec::new(),
            text: "Hello World".into(),
            paragraphs: Vec::new(),
        };
        assert_eq!(doc.plain_text(), "Hello World");
    }

    /// issue #248 — `plain_text()`/`to_markdown()` only ever walked
    /// `self.text` (the main body), never `self.subdocuments`, so a
    /// footnote/endnote/comment/textbox-only document silently vanished
    /// from both renderers even though `to_ir()` (via `doc_to_ir`) already
    /// carried the content correctly.
    #[test]
    fn plain_text_and_markdown_include_subdocument_bodies() {
        let doc = DocDocument {
            subdocuments: vec![
                SubDocument { kind: SubDocumentKind::Footnotes, text: "FOOTNOTE ONE".into() },
                SubDocument { kind: SubDocumentKind::Comments, text: "REVIEW NOTE".into() },
                SubDocument { kind: SubDocumentKind::HeaderTextBoxes, text: "SIDEBAR".into() },
            ],
            has_macros: false,
            text_complete: true,
            summary_properties: None,
            list_formatting: crate::doc::list_format::ListFormatting::default(),
            comment_authors: Vec::new(),
            images: Vec::new(),
            text: "Main body text".into(),
            paragraphs: Vec::new(),
        };
        let text = doc.plain_text();
        for token in ["Main body text", "FOOTNOTE ONE", "REVIEW NOTE", "SIDEBAR"] {
            assert!(text.contains(token), "{token} missing from plain_text(): {text:?}");
        }
        let md = doc.to_markdown();
        for token in ["Main body text", "FOOTNOTE ONE", "REVIEW NOTE", "SIDEBAR"] {
            assert!(md.contains(token), "{token} missing from to_markdown(): {md:?}");
        }
    }

    /// An empty subdocument body must contribute nothing — no stray blank
    /// paragraphs or extra separators.
    #[test]
    fn empty_subdocuments_are_skipped_in_both_renderers() {
        let doc = DocDocument {
            subdocuments: vec![SubDocument { kind: SubDocumentKind::Comments, text: "  \n ".into() }],
            has_macros: false,
            text_complete: true,
            summary_properties: None,
            list_formatting: crate::doc::list_format::ListFormatting::default(),
            comment_authors: Vec::new(),
            images: Vec::new(),
            text: "Body".into(),
            paragraphs: Vec::new(),
        };
        assert_eq!(doc.plain_text(), "Body");
        assert_eq!(doc.to_markdown(), "Body\n\n");
    }

    /// `text_complete()` reaches `to_ir()`'s `Metadata::text_truncated` so
    /// a caller who never inspects `DocDocument` directly can still tell a
    /// piece-table gap apart from a genuinely short document (issue #230).
    #[test]
    fn incomplete_text_reaches_metadata_as_truncated() {
        let doc = DocDocument {
            subdocuments: Vec::new(),
            has_macros: false,
            text_complete: false,
            summary_properties: None,
            list_formatting: crate::doc::list_format::ListFormatting::default(),
            comment_authors: Vec::new(),
            images: Vec::new(),
            text: "only the recovered fragment".into(),
            paragraphs: Vec::new(),
        };
        assert!(!doc.text_complete());
        let ir = crate::convert_doc::doc_to_ir(&doc);
        assert!(
            ir.metadata.text_truncated,
            "a piece-table gap must be visible on Metadata::text_truncated"
        );
    }

    /// The common case: a complete piece table must not be flagged.
    #[test]
    fn complete_text_is_not_flagged_truncated() {
        let doc = make_doc("Hello World");
        assert!(doc.text_complete());
        let ir = crate::convert_doc::doc_to_ir(&doc);
        assert!(!ir.metadata.text_truncated);
    }

    fn make_doc(text: &str) -> DocDocument {
        DocDocument {
            subdocuments: Vec::new(),
            has_macros: false,
            text_complete: true,
            summary_properties: None,
            list_formatting: crate::doc::list_format::ListFormatting::default(),
            comment_authors: Vec::new(),
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
            subdocuments: Vec::new(),
            has_macros: false,
            text_complete: true,
            summary_properties: None,
            list_formatting: crate::doc::list_format::ListFormatting::default(),
            comment_authors: Vec::new(),
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
            hyperlinks: Vec::new(),
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

    /// issue #244 — `SummaryInformation` fields must reach `Metadata`, and
    /// the declared title must beat the heading-guess title.
    #[test]
    fn ir_summary_properties_reach_metadata() {
        let mut doc = make_doc("SOME ALL-CAPS HEADING\nBody text follows.");
        doc.summary_properties = Some(SummaryProperties {
            title: Some("Declared Title".to_string()),
            subject: Some("Declared Subject".to_string()),
            author: Some("Declared Author".to_string()),
            keywords: Some("alpha, beta".to_string()),
            comments: Some("Declared Comment".to_string()),
            created: Some("2020-01-02T03:04:05Z".to_string()),
            modified: Some("2021-06-07T08:09:10Z".to_string()),
        });
        let ir = crate::convert_doc::doc_to_ir(&doc);
        assert_eq!(ir.metadata.title.as_deref(), Some("Declared Title"));
        assert_eq!(ir.metadata.author.as_deref(), Some("Declared Author"));
        assert_eq!(ir.metadata.subject.as_deref(), Some("Declared Subject"));
        assert_eq!(ir.metadata.keywords, vec!["alpha".to_string(), "beta".to_string()]);
        assert_eq!(ir.metadata.description.as_deref(), Some("Declared Comment"));
        assert_eq!(ir.metadata.created.as_deref(), Some("2020-01-02T03:04:05Z"));
        assert_eq!(ir.metadata.modified.as_deref(), Some("2021-06-07T08:09:10Z"));
    }

    /// A missing/empty title in `SummaryInformation` must not shadow the
    /// heading-guess fallback (issue #244 must not regress issue #224).
    #[test]
    fn ir_empty_summary_title_falls_back_to_heading_guess() {
        let mut doc = make_doc("A HEADING LINE\nBody text follows.");
        doc.summary_properties = Some(SummaryProperties {
            title: Some(String::new()),
            ..Default::default()
        });
        let ir = crate::convert_doc::doc_to_ir(&doc);
        assert_eq!(ir.metadata.title.as_deref(), Some("A HEADING LINE"));
    }

    /// issue #298 — `GrpXstAtnOwners` is an array of XSTs packed
    /// back-to-back with no STTBF-style header: each entry is a `cch`
    /// `u16` followed by `cch` UTF-16LE code units.
    #[test]
    fn grp_xst_atn_owners_parses_multiple_packed_entries() {
        let mut data = Vec::new();
        for name in ["Michael McCandless", "Miklos Vajna"] {
            let units: Vec<u16> = name.encode_utf16().collect();
            data.extend_from_slice(&(units.len() as u16).to_le_bytes());
            for u in units {
                data.extend_from_slice(&u.to_le_bytes());
            }
        }
        let names = parse_grp_xst_atn_owners(&data, 0, data.len() as u32);
        assert_eq!(names, vec!["Michael McCandless".to_string(), "Miklos Vajna".to_string()]);
    }

    #[test]
    fn grp_xst_atn_owners_zero_length_is_empty() {
        let data = vec![0u8; 32];
        assert!(parse_grp_xst_atn_owners(&data, 4, 0).is_empty());
    }

    /// issue #298 — a single declared comment author unambiguously
    /// attributes every comment in the document; `doc_to_ir` must carry
    /// it onto the merged Comments `Note` via the new `author` field.
    #[test]
    fn a_single_comment_author_reaches_the_comments_note() {
        use crate::ir::Element;
        let mut doc = make_doc("Body text.");
        doc.subdocuments =
            vec![SubDocument { kind: SubDocumentKind::Comments, text: "Here is a comment".into() }];
        doc.comment_authors = vec!["Michael McCandless".to_string()];

        let ir = crate::convert_doc::doc_to_ir(&doc);
        let note = ir.sections[0]
            .elements
            .iter()
            .find_map(|e| if let Element::Endnote(n) = e { Some(n) } else { None })
            .expect("expected a comments Endnote");
        assert_eq!(note.author.as_deref(), Some("Michael McCandless"));
    }

    /// Two or more declared authors cannot be attributed to this merged,
    /// per-subdocument (not per-comment) `Note` without `PlcfAtn`/`ATRD`
    /// correlation — leave `author` unset rather than guess wrong.
    #[test]
    fn multiple_comment_authors_leave_the_note_author_unset() {
        use crate::ir::Element;
        let mut doc = make_doc("Body text.");
        doc.subdocuments =
            vec![SubDocument { kind: SubDocumentKind::Comments, text: "Inner\nOuter".into() }];
        doc.comment_authors = vec!["vmiklos".to_string(), "Miklos Vajna".to_string()];

        let ir = crate::convert_doc::doc_to_ir(&doc);
        let note = ir.sections[0]
            .elements
            .iter()
            .find_map(|e| if let Element::Endnote(n) = e { Some(n) } else { None })
            .expect("expected a comments Endnote");
        assert_eq!(note.author, None);
    }
}
