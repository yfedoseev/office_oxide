//! # office_oxide::pptx
//!
//! High-performance PowerPoint presentation (.pptx) processing.
//!
//! Read, convert, and extract content from PPTX files
//! (Office Open XML PresentationML, ISO 29500 / ECMA-376).
//!
//! # Quick Start
//!
//! ```rust,no_run
//! use office_oxide::pptx::PptxDocument;
//!
//! let doc = PptxDocument::open("slides.pptx").unwrap();
//! println!("{}", doc.plain_text());
//! println!("{}", doc.to_markdown());
//! ```

/// In-place editing of PPTX documents.
pub mod edit;
/// Error types for PPTX parsing and creation.
pub mod error;
/// `ppt/presentation.xml` data model.
pub mod presentation;
/// Shape data model for PresentationML slides.
pub mod shape;
/// Slide XML parser.
pub mod slide;
/// Text extraction utilities for PPTX.
pub mod text;
/// PPTX creation (write) API.
pub mod write;

pub use error::{PptxError, Result};
pub use presentation::{PresentationInfo, SlideId, SlideSize};
pub use shape::{
    AutoShape, ConnectorShape, GraphicContent, GraphicFrame, GroupShape, HyperlinkInfo,
    HyperlinkTarget, PictureShape, PlaceholderInfo, Shape, ShapePosition, Table, TableCell,
    TableRow, TextBody, TextContent, TextField, TextParagraph, TextRun,
};
pub use slide::Slide;

use std::io::{Read, Seek};
use std::path::Path;

use crate::core::opc::OpcReader;
use crate::core::relationships::{Relationships, rel_types};
use crate::core::theme::Theme;
use log::debug;

/// A parsed PPTX document.
#[derive(Debug, Clone)]
pub struct PptxDocument {
    /// Metadata from `ppt/presentation.xml` (slide list, dimensions).
    pub presentation: PresentationInfo,
    /// Parsed slides, in presentation order.
    pub slides: Vec<Slide>,
    /// Theme data (colors, fonts), if present.
    pub theme: Option<Theme>,
}

impl PptxDocument {
    /// Open a PPTX file from a file path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let reader = OpcReader::open(path)?;
        Self::from_opc(reader)
    }

    /// Open a PPTX file using memory-mapped I/O for better performance on large files.
    #[cfg(feature = "mmap")]
    pub fn open_mmap(path: impl AsRef<Path>) -> Result<Self> {
        let reader = OpcReader::open_mmap(path)?;
        Self::from_opc(reader)
    }

    /// Open a PPTX document from any `Read + Seek` source.
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let opc = OpcReader::new(reader)?;
        Self::from_opc(opc)
    }

    fn from_opc<R: Read + Seek>(mut opc: OpcReader<R>) -> Result<Self> {
        debug!("PptxDocument: parsing started");
        let main_part = opc.main_document_part()?;
        let pres_rels = opc.read_rels_for(&main_part)?;

        // Parse theme
        let theme = if let Some(rel) = pres_rels.first_by_type(rel_types::THEME) {
            let part_name = main_part.resolve_relative(&rel.target)?;
            let data = opc.read_part(&part_name)?;
            Some(Theme::parse(&data)?)
        } else {
            None
        };

        // Parse presentation.xml
        let pres_data = opc.read_part(&main_part)?;
        let presentation = PresentationInfo::parse(&pres_data)?;

        // Phase 1: gather raw data sequentially (requires &mut opc)
        struct SlideBundle {
            slide_data: Vec<u8>,
            slide_rels: Relationships,
            notes_data: Option<Vec<u8>>,
        }
        let mut bundles = Vec::with_capacity(presentation.slides.len());
        for (slide_idx, slide_id) in presentation.slides.iter().enumerate() {
            // Try to resolve by rel_id, fall back to positional lookup
            let part_name = if !slide_id.rel_id.is_empty() {
                match pres_rels.resolve_target(&slide_id.rel_id, &main_part) {
                    Ok(pn) => pn,
                    Err(_) => continue,
                }
            } else {
                // No r:id — try convention: ppt/slides/slideN.xml
                let idx = slide_idx + 1;
                let candidate = format!("/ppt/slides/slide{}.xml", idx);
                match crate::core::opc::PartName::new(&candidate) {
                    Ok(pn) if opc.has_part(&pn) => pn,
                    _ => continue,
                }
            };
            let slide_rels = opc
                .read_rels_for(&part_name)
                .unwrap_or_else(|_| Relationships::empty());
            let slide_data = opc.read_part(&part_name)?;

            let notes_data =
                if let Some(notes_rel) = slide_rels.first_by_type(rel_types::NOTES_SLIDE) {
                    let notes_part = part_name.resolve_relative(&notes_rel.target)?;
                    if opc.has_part(&notes_part) {
                        Some(opc.read_part(&notes_part)?)
                    } else {
                        None
                    }
                } else {
                    None
                };

            bundles.push(SlideBundle {
                slide_data,
                slide_rels,
                notes_data,
            });
        }

        // Phase 2: parse slides (parallel when feature enabled)
        let slides = crate::core::parallel::map_collect(bundles, |b| -> Result<Slide> {
            let name = xml_csl_name(&b.slide_data);
            let mut parsed = Slide::parse(&b.slide_data, name, &b.slide_rels)?;
            if let Some(notes_data) = &b.notes_data {
                parsed.notes = extract_notes_text(notes_data);
            }
            Ok(parsed)
        })?;

        debug!("PptxDocument: {} slides parsed", slides.len());
        Ok(PptxDocument {
            presentation,
            slides,
            theme,
        })
    }
}

/// Extract the `name` attribute from `<p:cSld name="...">`, if present.
fn xml_csl_name(xml_data: &[u8]) -> String {
    use quick_xml::events::Event;
    let mut reader = crate::core::xml::make_fast_reader(xml_data);

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if e.local_name().as_ref() == b"cSld" =>
            {
                return crate::core::xml::optional_attr_str(e, b"name")
                    .ok()
                    .flatten()
                    .map(|v| v.into_owned())
                    .unwrap_or_default();
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {},
        }
    }
    String::new()
}

/// Extract speaker notes plain text from a notes slide XML.
/// Finds the body placeholder (type="body") and extracts its text.
fn extract_notes_text(xml_data: &[u8]) -> Option<String> {
    slide::extract_notes_text(xml_data)
}

impl crate::core::OfficeDocument for PptxDocument {
    fn plain_text(&self) -> String {
        self.plain_text()
    }

    fn to_markdown(&self) -> String {
        self.to_markdown()
    }
}
