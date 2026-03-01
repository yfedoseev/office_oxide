//! # pptx_oxide
//!
//! High-performance PowerPoint presentation (.pptx) processing.
//!
//! Read, convert, and extract content from PPTX files
//! (Office Open XML PresentationML, ISO 29500 / ECMA-376).
//!
//! # Quick Start
//!
//! ```rust,no_run
//! use pptx_oxide::PptxDocument;
//!
//! let doc = PptxDocument::open("slides.pptx").unwrap();
//! println!("{}", doc.plain_text());
//! println!("{}", doc.to_markdown());
//! ```

pub mod error;
pub mod presentation;
pub mod shape;
pub mod slide;
pub mod text;
pub mod write;
pub mod edit;

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

use log::debug;
use office_core::opc::OpcReader;
use office_core::relationships::{rel_types, Relationships};
use office_core::theme::Theme;

/// A parsed PPTX document.
#[derive(Debug, Clone)]
pub struct PptxDocument {
    pub presentation: PresentationInfo,
    pub slides: Vec<Slide>,
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
        for slide_id in &presentation.slides {
            let part_name = pres_rels.resolve_target(&slide_id.rel_id, &main_part)?;
            let slide_rels = opc
                .read_rels_for(&part_name)
                .unwrap_or_else(|_| Relationships::empty());
            let slide_data = opc.read_part(&part_name)?;

            let notes_data = if let Some(notes_rel) =
                slide_rels.first_by_type(rel_types::NOTES_SLIDE)
            {
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
        #[cfg(feature = "parallel")]
        let slides: Result<Vec<Slide>> = {
            use rayon::prelude::*;
            bundles
                .into_par_iter()
                .map(|b| {
                    let name = xml_csl_name(&b.slide_data);
                    let mut parsed = Slide::parse(&b.slide_data, name, &b.slide_rels)?;
                    if let Some(notes_data) = &b.notes_data {
                        parsed.notes = extract_notes_text(notes_data);
                    }
                    Ok(parsed)
                })
                .collect()
        };
        #[cfg(not(feature = "parallel"))]
        let slides: Result<Vec<Slide>> = bundles
            .into_iter()
            .map(|b| {
                let name = xml_csl_name(&b.slide_data);
                let mut parsed = Slide::parse(&b.slide_data, name, &b.slide_rels)?;
                if let Some(notes_data) = &b.notes_data {
                    parsed.notes = extract_notes_text(notes_data);
                }
                Ok(parsed)
            })
            .collect();
        let slides = slides?;

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
    let mut reader = office_core::xml::make_reader(xml_data);
    let pml = office_core::xml::ns::PML;

    loop {
        match reader.read_resolved_event() {
            Ok((ref resolve, Event::Start(ref e))) | Ok((ref resolve, Event::Empty(ref e))) => {
                if office_core::xml::matches_ns(resolve, pml)
                    && e.local_name().as_ref() == b"cSld"
                {
                    return office_core::xml::optional_attr_str(e, b"name")
                        .ok()
                        .flatten()
                        .map(|v| v.into_owned())
                        .unwrap_or_default();
                }
            }
            Ok((_, Event::Eof)) => break,
            Err(_) => break,
            _ => {}
        }
    }
    String::new()
}

/// Extract speaker notes plain text from a notes slide XML.
/// Finds the body placeholder (type="body") and extracts its text.
fn extract_notes_text(xml_data: &[u8]) -> Option<String> {
    slide::extract_notes_text(xml_data)
}
