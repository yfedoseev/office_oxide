//! # office_oxide::docx
//!
//! High-performance Word document (.docx) processing.
//!
//! Read, convert, and extract content from DOCX files
//! (Office Open XML WordprocessingML, ISO 29500 / ECMA-376).
//!
//! # Quick Start
//!
//! ```rust,no_run
//! use office_oxide::docx::DocxDocument;
//!
//! let doc = DocxDocument::open("report.docx").unwrap();
//! println!("{}", doc.plain_text());
//! println!("{}", doc.to_markdown());
//! ```

/// Document body and block-level element types.
pub mod document;
/// In-place editing of existing DOCX files.
pub mod edit;
/// DOCX-specific error type.
pub mod error;
/// Run and paragraph formatting types (`RunProperties`, `ParagraphProperties`, etc.).
pub mod formatting;
/// Section properties, headers, footers, page size/margin types.
pub mod headers;
/// Hyperlink types (`Hyperlink`, `HyperlinkTarget`).
pub mod hyperlink;
/// Drawing/image reference type (`DrawingInfo`).
pub mod image;
/// Numbering definitions and list format types.
pub mod numbering;
/// Paragraph, run, and inline content types.
pub mod paragraph;
/// Style sheet and style definition types.
pub mod styles;
/// Table structure types.
pub mod table;
/// Text extraction and markdown rendering for DOCX.
pub mod text;
/// DOCX creation (write) API.
pub mod write;

pub use document::{BlockElement, Body};
pub use error::{DocxError, Result};
pub use formatting::{
    Justification, ParagraphIndent, ParagraphProperties, ParagraphSpacing, RunProperties,
    UnderlineType, VerticalAlign,
};
pub use headers::{
    HeaderFooter, HeaderFooterType, PageMargins, PageOrientation, PageSize, SectionProperties,
};
pub use hyperlink::{Hyperlink, HyperlinkTarget};
pub use image::{AnchorFrame, AnchorPosition, DrawingInfo, ShapeInfo, ShapeKind};
pub use numbering::{NumberFormat, NumberingDefinitions};
pub use paragraph::{BreakType, Paragraph, ParagraphContent, Run, RunContent};
pub use styles::{Style, StyleSheet, StyleType};
pub use table::{Table, TableCell, TableProperties, TableRow};

use std::io::{Read, Seek};
use std::path::Path;

use log::debug;
use quick_xml::events::Event;

use crate::core::opc::OpcReader;
use crate::core::relationships::{TargetMode, rel_types};
use crate::core::theme::Theme;
use crate::core::units::Emu;
use crate::core::xml;

use self::formatting::{parse_paragraph_properties_fast, parse_run_properties_fast};
use self::headers::HeaderFooterRef;
use self::table::{
    MergeType, Shading, TableCellProperties, TableRowProperties, TableWidth, TableWidthType,
};

// Use crate::core::Result internally for all XML parsing (it has From<quick_xml::Error>).
// DocxError wraps crate::core::Error, so conversion at the public boundary is automatic via `?`.
type CoreResult<T> = crate::core::Result<T>;

/// Create a fast reader that does NOT trim text content.
/// Unlike `xml::make_reader`, this preserves whitespace so `xml:space="preserve"` works.
fn make_content_reader(xml_data: &[u8]) -> quick_xml::Reader<&[u8]> {
    let mut reader = quick_xml::Reader::from_reader(xml_data);
    reader.config_mut().check_end_names = false;
    reader.config_mut().check_comments = false;
    reader
}

/// A parsed DOCX document.
#[derive(Debug, Clone)]
pub struct DocxDocument {
    /// The document body.
    pub body: Body,
    /// Parsed stylesheet from `word/styles.xml`.
    pub styles: Option<StyleSheet>,
    /// Numbering definitions from `word/numbering.xml`.
    pub numbering: Option<NumberingDefinitions>,
    /// Theme from the document.
    pub theme: Option<Theme>,
    /// Section properties (from the last `w:sectPr` in the body).
    pub sections: Vec<SectionProperties>,
    /// Parsed headers and footers.
    pub headers_footers: Vec<HeaderFooter>,
    /// Font programs found under `word/fonts/`. Each entry is
    /// `(font_name, ttf_or_otf_bytes)`. PDF→DOCX→PDF round-trips use these
    /// to preserve typeface fidelity (e.g. CJK / math fonts beyond
    /// pdf_oxide's bundled DejaVu fallback).
    pub embedded_fonts: Vec<(String, Vec<u8>)>,
    /// Image parts referenced from the main document, keyed by the
    /// relationship id used in `<a:blip r:embed="rIdN"/>`. Lets the
    /// IR converter populate `Image::data` so downstream renderers
    /// (the positional PDF reader, plain-text export with alt-text,
    /// etc.) can place actual bitmap content.
    pub images: std::collections::HashMap<String, (Vec<u8>, Option<String>)>,
}

impl DocxDocument {
    /// Open a DOCX file from a file path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let reader = OpcReader::open(path)?;
        Self::from_opc(reader)
    }

    /// Open a DOCX file using memory-mapped I/O for better performance on large files.
    #[cfg(feature = "mmap")]
    pub fn open_mmap(path: impl AsRef<Path>) -> Result<Self> {
        let reader = OpcReader::open_mmap(path)?;
        Self::from_opc(reader)
    }

    /// Open a DOCX document from any `Read + Seek` source.
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let opc = OpcReader::new(reader)?;
        Self::from_opc(opc)
    }

    fn from_opc<R: Read + Seek>(mut opc: OpcReader<R>) -> Result<Self> {
        debug!("DocxDocument: parsing started");
        let main_part = opc.main_document_part()?;
        let doc_rels = opc.read_rels_for(&main_part)?;

        // Parse theme
        let theme = if let Some(rel) = doc_rels.first_by_type(rel_types::THEME) {
            let part_name = main_part.resolve_relative(&rel.target)?;
            let data = opc.read_part(&part_name)?;
            Some(Theme::parse(&data)?)
        } else {
            None
        };

        // Parse styles
        let styles = if let Some(rel) = doc_rels.first_by_type(rel_types::STYLES) {
            let part_name = main_part.resolve_relative(&rel.target)?;
            let data = opc.read_part(&part_name)?;
            Some(StyleSheet::parse(&data)?)
        } else {
            None
        };

        // Parse numbering (optional — some files reference it but don't include it)
        let numbering = if let Some(rel) = doc_rels.first_by_type(rel_types::NUMBERING) {
            let part_name = main_part.resolve_relative(&rel.target)?;
            match opc.read_part(&part_name) {
                Ok(data) => Some(NumberingDefinitions::parse(&data)?),
                Err(_) => None,
            }
        } else {
            None
        };

        // Parse main document
        let doc_data = opc.read_part(&main_part)?;
        let (body, sections) = parse_document(&doc_data, &doc_rels)?;

        // Parse headers and footers. Walk header refs and footer refs
        // separately so each parsed `HeaderFooter` can record its own
        // role; without that distinction, downstream consumers had to
        // back-derive headers-vs-footers from cumulative ref counts,
        // which silently misclassifies entries in multi-section docs.
        let mut headers_footers = Vec::new();
        let mut parse_hf = |hf_ref: &HeaderFooterRef, is_header: bool| -> CoreResult<()> {
            if let Some(rel) = doc_rels.get_by_id(&hf_ref.relationship_id) {
                if rel.target_mode == TargetMode::Internal {
                    let part_name = main_part.resolve_relative(&rel.target)?;
                    if opc.has_part(&part_name) {
                        let data = opc.read_part(&part_name)?;
                        let content = parse_body_elements(&data)?;
                        headers_footers.push(HeaderFooter {
                            hf_type: hf_ref.hf_type,
                            content,
                            is_header,
                        });
                    }
                }
            }
            Ok(())
        };
        for section in &sections {
            for hf_ref in &section.header_refs {
                parse_hf(hf_ref, true)?;
            }
            for hf_ref in &section.footer_refs {
                parse_hf(hf_ref, false)?;
            }
        }

        // Scan `word/fonts/` for embedded font programs. Files there are
        // typically `font_<n>_<name>.ttf` (written by our own `DocxWriter`)
        // but the loop accepts any `.ttf`/`.otf` for forward-compat.
        let mut embedded_fonts: Vec<(String, Vec<u8>)> = Vec::new();
        for name in opc.part_names() {
            let s = name.to_string();
            if !s.starts_with("/word/fonts/") {
                continue;
            }
            let lower = s.to_lowercase();
            if !(lower.ends_with(".ttf") || lower.ends_with(".otf")) {
                continue;
            }
            if let Ok(data) = opc.read_part(&name) {
                // Extract a usable face name from the OPC part. Writers
                // ship fonts as `font_<n>_<face_name>.<ext>` (the
                // `embedded_fonts` writer convention used by all three
                // PDF→office paths) — strip the leading `font_<n>_`
                // prefix and the trailing `.ttf`/`.otf` so the
                // registered name matches what the IR carries on each
                // run's `font_name` (e.g. `TeXGyreTermesX-Regular`).
                // Falls back to the basename for files that don't
                // follow the convention.
                let basename = s.rsplit('/').next().unwrap_or("font");
                let face = strip_embedded_font_filename(basename);
                let font_name = if face.is_empty() {
                    basename.to_string()
                } else {
                    face
                };
                embedded_fonts.push((font_name, data));
            }
        }

        // Pull image parts referenced by the main document
        // relationships. We capture the raw bytes plus the lower-cased
        // file extension so downstream code can decide on the format
        // without re-sniffing magic bytes.
        let mut images: std::collections::HashMap<String, (Vec<u8>, Option<String>)> =
            std::collections::HashMap::new();
        for rel in doc_rels.get_by_type(rel_types::IMAGE) {
            if rel.target_mode != TargetMode::Internal {
                continue;
            }
            let part_name = match main_part.resolve_relative(&rel.target) {
                Ok(p) => p,
                Err(_) => continue,
            };
            if !opc.has_part(&part_name) {
                continue;
            }
            let data = match opc.read_part(&part_name) {
                Ok(d) => d,
                Err(_) => continue,
            };
            let ext = part_name
                .as_str()
                .rsplit('.')
                .next()
                .map(|s| s.to_lowercase());
            images.insert(rel.id.clone(), (data, ext));
        }

        debug!(
            "DocxDocument: {} block elements, {} sections, {} embedded fonts, {} images",
            body.elements.len(),
            sections.len(),
            embedded_fonts.len(),
            images.len()
        );
        Ok(DocxDocument {
            body,
            styles,
            numbering,
            theme,
            sections,
            headers_footers,
            embedded_fonts,
            images,
        })
    }
}

/// Parse body-level elements from XML (used for headers/footers which share the same structure).
fn parse_body_elements(xml_data: &[u8]) -> CoreResult<Vec<BlockElement>> {
    let mut reader = make_content_reader(xml_data);
    let mut elements = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"p" => {
                    elements.push(BlockElement::Paragraph(parse_paragraph(&mut reader)?));
                },
                b"tbl" => {
                    elements.push(BlockElement::Table(parse_table(&mut reader)?));
                },
                _ => {},
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(elements)
}

/// Parse `word/document.xml` and return the Body and SectionProperties.
fn parse_document(
    xml_data: &[u8],
    rels: &crate::core::relationships::Relationships,
) -> CoreResult<(Body, Vec<SectionProperties>)> {
    let mut reader = make_content_reader(xml_data);
    let mut elements = Vec::new();
    let mut sections = Vec::new();
    let mut in_body = false;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"body" => {
                    in_body = true;
                },
                b"p" if in_body => {
                    elements.push(BlockElement::Paragraph(parse_paragraph(&mut reader)?));
                },
                b"tbl" if in_body => {
                    elements.push(BlockElement::Table(parse_table(&mut reader)?));
                },
                b"sectPr" if in_body => {
                    sections.push(parse_section_properties(&mut reader, e)?);
                },
                _ => {},
            },
            Event::End(ref e) if e.local_name().as_ref() == b"body" => {
                in_body = false;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    // Resolve hyperlink targets using relationships
    resolve_hyperlinks(&mut elements, rels);

    // Detect mid-document section breaks: paragraphs whose <w:pPr>
    // carries a <w:sectPr>. Each such paragraph terminates a section,
    // and its sectPr describes the section that ends there. Trailing
    // elements after the last break belong to a final section
    // described by the body-level sectPr (already in `sections`).
    let mut section_breaks: Vec<usize> = Vec::new();
    let mut break_sections: Vec<SectionProperties> = Vec::new();
    for (idx, el) in elements.iter().enumerate() {
        if let BlockElement::Paragraph(p) = el {
            if let Some(props) = &p.properties {
                if let Some(sp) = &props.section_properties {
                    section_breaks.push(idx + 1);
                    break_sections.push(sp.clone());
                }
            }
        }
    }
    // Stitch break-derived section_properties in front of the
    // body-level final sectPr so the section list is in document order.
    let mut all_sections = break_sections;
    all_sections.extend(sections);

    let body = Body {
        elements,
        section_breaks,
    };
    Ok((body, all_sections))
}

/// Walk the element tree and resolve hyperlink rIds to actual URLs.
fn resolve_hyperlinks(
    elements: &mut [BlockElement],
    rels: &crate::core::relationships::Relationships,
) {
    for elem in elements.iter_mut() {
        match elem {
            BlockElement::Paragraph(p) => {
                for content in &mut p.content {
                    if let ParagraphContent::Hyperlink(hl) = content {
                        if let HyperlinkTarget::External(ref r_id) = hl.target {
                            if let Some(rel) = rels.get_by_id(r_id) {
                                if rel.target_mode == TargetMode::External {
                                    hl.target = HyperlinkTarget::External(rel.target.clone());
                                } else {
                                    hl.target = HyperlinkTarget::Internal(rel.target.clone());
                                }
                            }
                        }
                    }
                }
            },
            BlockElement::Table(t) => {
                for row in &mut t.rows {
                    for cell in &mut row.cells {
                        resolve_hyperlinks(&mut cell.content, rels);
                    }
                }
            },
        }
    }
}

// ---------------------------------------------------------------------------
// Paragraph parsing
// ---------------------------------------------------------------------------

fn parse_paragraph(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<Paragraph> {
    let mut paragraph = Paragraph::default();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"pPr" => {
                    paragraph.properties = Some(parse_paragraph_properties_fast(reader)?);
                },
                b"r" => {
                    paragraph
                        .content
                        .push(ParagraphContent::Run(parse_run(reader)?));
                },
                b"hyperlink" => {
                    paragraph
                        .content
                        .push(ParagraphContent::Hyperlink(parse_hyperlink(reader, e)?));
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) if e.local_name().as_ref() == b"p" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(paragraph)
}

fn parse_run(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<Run> {
    let mut run = Run::default();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"rPr" => {
                    run.properties = Some(parse_run_properties_fast(reader)?);
                },
                b"t" => {
                    let text = xml::read_text_content_fast(reader)?;
                    if !text.is_empty() {
                        run.content.push(RunContent::Text(text));
                    }
                },
                b"br" => {
                    let break_type = match xml::optional_attr_str(e, b"w:type")? {
                        Some(ref t) => match t.as_ref() {
                            "page" => BreakType::Page,
                            "column" => BreakType::Column,
                            _ => BreakType::Line,
                        },
                        None => BreakType::Line,
                    };
                    run.content.push(RunContent::Break(break_type));
                    xml::skip_element_fast(reader)?;
                },
                b"drawing" => {
                    if let Some(drawing) = parse_drawing(reader)? {
                        run.content.push(RunContent::Drawing(drawing));
                    }
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::Empty(ref e) => match e.local_name().as_ref() {
                b"br" => {
                    let break_type = match xml::optional_attr_str(e, b"w:type")? {
                        Some(ref t) => match t.as_ref() {
                            "page" => BreakType::Page,
                            "column" => BreakType::Column,
                            _ => BreakType::Line,
                        },
                        None => BreakType::Line,
                    };
                    run.content.push(RunContent::Break(break_type));
                },
                b"tab" => {
                    run.content.push(RunContent::Tab);
                },
                _ => {},
            },
            Event::End(ref e) if e.local_name().as_ref() == b"r" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(run)
}

fn parse_hyperlink(
    reader: &mut quick_xml::Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> CoreResult<Hyperlink> {
    // Determine target: r:id for external, w:anchor for internal
    let r_id = xml::optional_attr_str(start, b"r:id")?.map(|v| v.into_owned());
    let anchor = xml::optional_attr_str(start, b"w:anchor")?.map(|v| v.into_owned());
    let tooltip = xml::optional_attr_str(start, b"w:tooltip")?.map(|v| v.into_owned());

    let target = if let Some(anchor) = anchor {
        HyperlinkTarget::Internal(anchor)
    } else if let Some(r_id) = r_id {
        // Will be resolved to actual URL after parsing via resolve_hyperlinks()
        HyperlinkTarget::External(r_id)
    } else {
        HyperlinkTarget::Internal(String::new())
    };

    let mut runs = Vec::new();
    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.local_name().as_ref() == b"r" {
                    runs.push(parse_run(reader)?);
                } else {
                    xml::skip_element_fast(reader)?;
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"hyperlink" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(Hyperlink {
        target,
        tooltip,
        runs,
    })
}

// ---------------------------------------------------------------------------
// Drawing / image parsing
// ---------------------------------------------------------------------------

/// Parse a `<w:drawing>` element. The opening tag has already been
/// consumed by the caller, so we drive forward until the matching
/// `</w:drawing>` End event.
///
/// A drawing wraps either `<wp:inline>` or `<wp:anchor>` (anchor =
/// floating). Everything we care about lives inside that single
/// wrapper, so we delegate to `parse_inline_or_anchor_body` and treat
/// any other top-level event as ignorable filler.
fn parse_drawing(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<Option<DrawingInfo>> {
    let mut info: Option<DrawingInfo> = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"inline" => {
                    info = parse_inline_or_anchor_body(reader, /*inline=*/ true, b"inline")?;
                },
                b"anchor" => {
                    info = parse_inline_or_anchor_body(reader, /*inline=*/ false, b"anchor")?;
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) if e.local_name().as_ref() == b"drawing" => break,
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(info)
}

/// Parse the body of `<wp:inline>` or `<wp:anchor>` until the matching
/// closing tag (`end_local`). Collects extent, docPr, position, and the
/// graphic payload (image or shape) into a `DrawingInfo`.
fn parse_inline_or_anchor_body(
    reader: &mut quick_xml::Reader<&[u8]>,
    inline: bool,
    end_local: &[u8],
) -> CoreResult<Option<DrawingInfo>> {
    use crate::docx::image::{AnchorFrame, AnchorPosition};

    let mut width = Emu(0);
    let mut height = Emu(0);
    let mut description: Option<String> = None;
    let mut relationship_id: Option<String> = None;
    let mut shape: Option<crate::docx::image::ShapeInfo> = None;

    let mut anchor_x: Option<i64> = None;
    let mut anchor_y: Option<i64> = None;
    let mut h_frame = AnchorFrame::default();
    let mut v_frame = AnchorFrame::default();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"extent" => {
                    parse_extent_attrs(e, &mut width, &mut height);
                    xml::skip_element_fast(reader)?;
                },
                b"docPr" => {
                    if let Some(desc) = xml::optional_attr_str(e, b"descr")? {
                        description = Some(desc.into_owned());
                    }
                    xml::skip_element_fast(reader)?;
                },
                b"positionH" => {
                    if let Some(rf) = xml::optional_attr_str(e, b"relativeFrom")? {
                        h_frame = parse_anchor_frame(&rf);
                    }
                    anchor_x = parse_position_offset(reader, b"positionH")?;
                },
                b"positionV" => {
                    if let Some(rf) = xml::optional_attr_str(e, b"relativeFrom")? {
                        v_frame = parse_anchor_frame(&rf);
                    }
                    anchor_y = parse_position_offset(reader, b"positionV")?;
                },
                b"graphic" => {
                    let g = parse_graphic(reader)?;
                    if let Some(rid) = g.relationship_id {
                        relationship_id = Some(rid);
                    }
                    if let Some(s) = g.shape {
                        shape = Some(s);
                    }
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::Empty(ref e) => match e.local_name().as_ref() {
                b"extent" => parse_extent_attrs(e, &mut width, &mut height),
                b"docPr" => {
                    if let Some(desc) = xml::optional_attr_str(e, b"descr")? {
                        description = Some(desc.into_owned());
                    }
                },
                _ => {},
            },
            Event::End(ref e) if e.local_name().as_ref() == end_local => break,
            Event::Eof => break,
            _ => {},
        }
    }

    let anchor_position = if !inline && (anchor_x.is_some() || anchor_y.is_some()) {
        Some(AnchorPosition {
            x_emu: anchor_x.unwrap_or(0),
            y_emu: anchor_y.unwrap_or(0),
            h_relative_from: h_frame,
            v_relative_from: v_frame,
        })
    } else {
        None
    };

    if relationship_id.is_some() || shape.is_some() {
        Ok(Some(DrawingInfo {
            relationship_id: relationship_id.unwrap_or_default(),
            description,
            width,
            height,
            inline,
            anchor_position,
            shape,
        }))
    } else {
        Ok(None)
    }
}

/// Parse the inside of `<wp:positionH>` or `<wp:positionV>` looking for
/// the nested `<wp:posOffset>` text value. Reads through the matching
/// closing tag (`end_local`).
fn parse_position_offset(
    reader: &mut quick_xml::Reader<&[u8]>,
    end_local: &[u8],
) -> CoreResult<Option<i64>> {
    let mut offset: Option<i64> = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == b"posOffset" => {
                let text = xml::read_text_content_fast(reader)?;
                if let Ok(v) = text.trim().parse::<i64>() {
                    offset = Some(v);
                }
            },
            Event::Start(_) => {
                xml::skip_element_fast(reader)?;
            },
            Event::End(ref e) if e.local_name().as_ref() == end_local => break,
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(offset)
}

/// Result of parsing an `<a:graphic>` element: at most one of an
/// embedded picture (`relationship_id`) or a vector shape (`shape`).
struct GraphicPayload {
    relationship_id: Option<String>,
    shape: Option<crate::docx::image::ShapeInfo>,
}

/// Parse `<a:graphic>` and any contained `<pic:pic>` (image) or
/// `<wps:wsp>` (vector shape). Reads through `</a:graphic>`.
fn parse_graphic(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<GraphicPayload> {
    let mut relationship_id: Option<String> = None;
    let mut shape: Option<crate::docx::image::ShapeInfo> = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"pic" => {
                    if let Some(rid) = parse_pic(reader)? {
                        relationship_id = Some(rid);
                    }
                },
                b"wsp" => {
                    if let Some(s) = parse_wsp(reader)? {
                        shape = Some(s);
                    }
                },
                // <a:graphicData> is just a wrapper; descend into it.
                b"graphicData" => continue,
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) if e.local_name().as_ref() == b"graphic" => break,
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(GraphicPayload {
        relationship_id,
        shape,
    })
}

/// Parse `<pic:pic>` looking for the embedded `<a:blip r:embed="…"/>`.
/// Reads through `</pic:pic>`. The blip lives inside `<pic:blipFill>`,
/// so we descend through whatever wrappers we encounter rather than
/// skipping siblings.
fn parse_pic(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<Option<String>> {
    let mut rid: Option<String> = None;
    // Track depth relative to <pic:pic>: we entered after its Start was
    // consumed by the caller, so we are at depth 1. Exit when we close
    // back out.
    let mut depth: u32 = 1;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.local_name().as_ref() == b"blip" {
                    if let Some(embed) = xml::optional_attr_str(e, b"r:embed")? {
                        rid = Some(embed.into_owned());
                    }
                    // Skip over blip's own children (e.g. <a:extLst>).
                    xml::skip_element_fast(reader)?;
                } else {
                    depth += 1;
                }
            },
            Event::Empty(ref e) if e.local_name().as_ref() == b"blip" => {
                if let Some(embed) = xml::optional_attr_str(e, b"r:embed")? {
                    rid = Some(embed.into_owned());
                }
            },
            Event::End(_) => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(rid)
}

/// Parse `<wps:wsp>` (a DrawingML vector shape). Reads through
/// `</wps:wsp>` and returns the assembled `ShapeInfo`, or `None` if no
/// `<a:prstGeom>` was seen.
fn parse_wsp(
    reader: &mut quick_xml::Reader<&[u8]>,
) -> CoreResult<Option<crate::docx::image::ShapeInfo>> {
    use crate::docx::image::{ShapeInfo, ShapeKind};

    let mut kind: Option<ShapeKind> = None;
    let mut stroke_rgb: Option<(u8, u8, u8)> = None;
    let mut fill_rgb: Option<(u8, u8, u8)> = None;
    let mut stroke_w_emu: Option<i64> = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"spPr" => {
                    parse_sp_pr(
                        reader,
                        &mut kind,
                        &mut stroke_rgb,
                        &mut fill_rgb,
                        &mut stroke_w_emu,
                    )?;
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) if e.local_name().as_ref() == b"wsp" => break,
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(kind.map(|k| ShapeInfo {
        kind: k,
        stroke_rgb,
        fill_rgb,
        stroke_w_emu,
    }))
}

/// Parse `<wps:spPr>`: contains the geometry preset, an optional fill,
/// and an optional `<a:ln>` (line/stroke) sub-element. Reads through
/// `</wps:spPr>`.
fn parse_sp_pr(
    reader: &mut quick_xml::Reader<&[u8]>,
    kind: &mut Option<crate::docx::image::ShapeKind>,
    stroke_rgb: &mut Option<(u8, u8, u8)>,
    fill_rgb: &mut Option<(u8, u8, u8)>,
    stroke_w_emu: &mut Option<i64>,
) -> CoreResult<()> {
    use crate::docx::image::ShapeKind;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"prstGeom" => {
                    if let Some(prst) = xml::optional_attr_str(e, b"prst")? {
                        *kind = match prst.as_ref() {
                            "line" | "straightConnector1" => Some(ShapeKind::Line),
                            "rect" => Some(ShapeKind::Rect),
                            _ => *kind,
                        };
                    }
                    xml::skip_element_fast(reader)?;
                },
                b"ln" => {
                    if let Some(w) = xml::optional_attr_str(e, b"w")? {
                        *stroke_w_emu = w.parse().ok();
                    }
                    *stroke_rgb = parse_line_color(reader)?.or(*stroke_rgb);
                },
                b"solidFill" => {
                    *fill_rgb = parse_solid_fill_color(reader)?.or(*fill_rgb);
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::Empty(ref e) => match e.local_name().as_ref() {
                b"prstGeom" => {
                    if let Some(prst) = xml::optional_attr_str(e, b"prst")? {
                        *kind = match prst.as_ref() {
                            "line" | "straightConnector1" => Some(ShapeKind::Line),
                            "rect" => Some(ShapeKind::Rect),
                            _ => *kind,
                        };
                    }
                },
                b"ln" => {
                    if let Some(w) = xml::optional_attr_str(e, b"w")? {
                        *stroke_w_emu = w.parse().ok();
                    }
                },
                _ => {},
            },
            Event::End(ref e) if e.local_name().as_ref() == b"spPr" => break,
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(())
}

/// Parse `<a:ln>` looking for an inner `<a:solidFill><a:srgbClr/>`.
/// Reads through `</a:ln>`.
fn parse_line_color(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<Option<(u8, u8, u8)>> {
    let mut rgb: Option<(u8, u8, u8)> = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"solidFill" => {
                    if let Some(c) = parse_solid_fill_color(reader)? {
                        rgb = Some(c);
                    }
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) if e.local_name().as_ref() == b"ln" => break,
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(rgb)
}

/// Parse `<a:solidFill>` looking for an inner `<a:srgbClr val="…"/>`.
/// Reads through `</a:solidFill>`.
fn parse_solid_fill_color(
    reader: &mut quick_xml::Reader<&[u8]>,
) -> CoreResult<Option<(u8, u8, u8)>> {
    let mut rgb: Option<(u8, u8, u8)> = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.local_name().as_ref() == b"srgbClr" {
                    if let Some(val) = xml::optional_attr_str(e, b"val")? {
                        if let Some(parsed) = parse_hex_rgb(&val) {
                            rgb = Some(parsed);
                        }
                    }
                }
                xml::skip_element_fast(reader)?;
            },
            Event::Empty(ref e) if e.local_name().as_ref() == b"srgbClr" => {
                if let Some(val) = xml::optional_attr_str(e, b"val")? {
                    if let Some(parsed) = parse_hex_rgb(&val) {
                        rgb = Some(parsed);
                    }
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"solidFill" => break,
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(rgb)
}

fn parse_anchor_frame(s: &str) -> crate::docx::image::AnchorFrame {
    use crate::docx::image::AnchorFrame;
    match s {
        "page" => AnchorFrame::Page,
        "margin" | "leftMargin" | "rightMargin" | "topMargin" | "bottomMargin" | "insideMargin"
        | "outsideMargin" => AnchorFrame::Margin,
        "column" => AnchorFrame::Column,
        "paragraph" => AnchorFrame::Paragraph,
        "line" => AnchorFrame::Line,
        "character" => AnchorFrame::Character,
        _ => AnchorFrame::Page,
    }
}

fn parse_hex_rgb(s: &str) -> Option<(u8, u8, u8)> {
    let bytes = s.trim().as_bytes();
    if bytes.len() != 6 {
        return None;
    }
    fn hex_pair(a: u8, b: u8) -> Option<u8> {
        let h = |c: u8| match c {
            b'0'..=b'9' => Some(c - b'0'),
            b'a'..=b'f' => Some(10 + c - b'a'),
            b'A'..=b'F' => Some(10 + c - b'A'),
            _ => None,
        };
        Some((h(a)? << 4) | h(b)?)
    }
    let r = hex_pair(bytes[0], bytes[1])?;
    let g = hex_pair(bytes[2], bytes[3])?;
    let b = hex_pair(bytes[4], bytes[5])?;
    Some((r, g, b))
}

fn parse_extent_attrs(e: &quick_xml::events::BytesStart, width: &mut Emu, height: &mut Emu) {
    if let Ok(Some(cx)) = xml::optional_attr_str(e, b"cx") {
        *width = Emu(cx.parse().unwrap_or(0));
    }
    if let Ok(Some(cy)) = xml::optional_attr_str(e, b"cy") {
        *height = Emu(cy.parse().unwrap_or(0));
    }
}

// ---------------------------------------------------------------------------
// Table parsing
// ---------------------------------------------------------------------------

fn parse_table(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<Table> {
    let mut properties = None;
    let mut grid = Vec::new();
    let mut rows = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"tblPr" => {
                    properties = Some(parse_table_properties(reader)?);
                },
                b"tblGrid" => {
                    grid = parse_table_grid(reader)?;
                },
                b"tr" => {
                    rows.push(parse_table_row(reader)?);
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) if e.local_name().as_ref() == b"tbl" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(Table {
        properties,
        grid,
        rows,
    })
}

fn parse_table_properties(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<TableProperties> {
    let mut props = TableProperties::default();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"tblW" => {
                    props.width = parse_table_width(e)?;
                    xml::skip_element_fast(reader)?;
                },
                b"jc" => {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                        props.justification =
                            Some(self::formatting::parse_justification_value(&val));
                    }
                    xml::skip_element_fast(reader)?;
                },
                b"tblStyle" => {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                        props.style_id = Some(val.into_owned());
                    }
                    xml::skip_element_fast(reader)?;
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::Empty(ref e) => match e.local_name().as_ref() {
                b"tblW" => {
                    props.width = parse_table_width(e)?;
                },
                b"jc" => {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                        props.justification =
                            Some(self::formatting::parse_justification_value(&val));
                    }
                },
                b"tblStyle" => {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                        props.style_id = Some(val.into_owned());
                    }
                },
                _ => {},
            },
            Event::End(ref e) if e.local_name().as_ref() == b"tblPr" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(props)
}

fn parse_table_grid(
    reader: &mut quick_xml::Reader<&[u8]>,
) -> CoreResult<Vec<crate::core::units::Twip>> {
    let mut cols = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) if e.local_name().as_ref() == b"gridCol" => {
                if let Ok(Some(w)) = xml::optional_attr_str(e, b"w:w") {
                    let val: i32 = w.parse().unwrap_or(0);
                    cols.push(crate::core::units::Twip(val));
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"tblGrid" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(cols)
}

fn parse_table_row(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<TableRow> {
    let mut properties = None;
    let mut cells = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"trPr" => {
                    properties = Some(parse_table_row_properties(reader)?);
                },
                b"tc" => {
                    cells.push(parse_table_cell(reader)?);
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) if e.local_name().as_ref() == b"tr" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(TableRow { properties, cells })
}

fn parse_table_row_properties(
    reader: &mut quick_xml::Reader<&[u8]>,
) -> CoreResult<TableRowProperties> {
    let mut props = TableRowProperties::default();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e)
                if e.local_name().as_ref() == b"tblHeader" =>
            {
                props.is_header = true;
            },
            Event::End(ref e) if e.local_name().as_ref() == b"trPr" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(props)
}

fn parse_table_cell(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<TableCell> {
    let mut properties = None;
    let mut content = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"tcPr" => {
                    properties = Some(parse_table_cell_properties(reader)?);
                },
                b"p" => {
                    content.push(BlockElement::Paragraph(parse_paragraph(reader)?));
                },
                b"tbl" => {
                    content.push(BlockElement::Table(parse_table(reader)?));
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) if e.local_name().as_ref() == b"tc" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(TableCell {
        properties,
        content,
    })
}

fn parse_table_cell_properties(
    reader: &mut quick_xml::Reader<&[u8]>,
) -> CoreResult<TableCellProperties> {
    let mut props = TableCellProperties::default();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"tcW" => {
                    props.width = parse_table_width(e)?;
                    xml::skip_element_fast(reader)?;
                },
                b"vMerge" => {
                    let val = xml::optional_attr_str(e, b"w:val")?;
                    props.vertical_merge = Some(match val.as_deref() {
                        Some("restart") => MergeType::Restart,
                        _ => MergeType::Continue,
                    });
                    xml::skip_element_fast(reader)?;
                },
                b"gridSpan" => {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                        props.grid_span = val.parse().ok();
                    }
                    xml::skip_element_fast(reader)?;
                },
                b"shd" => {
                    props.shading = Some(Shading {
                        fill: xml::optional_attr_str(e, b"w:fill")?.map(|v| v.into_owned()),
                        color: xml::optional_attr_str(e, b"w:color")?.map(|v| v.into_owned()),
                        pattern: xml::optional_attr_str(e, b"w:val")?.map(|v| v.into_owned()),
                    });
                    xml::skip_element_fast(reader)?;
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::Empty(ref e) => match e.local_name().as_ref() {
                b"tcW" => {
                    props.width = parse_table_width(e)?;
                },
                b"vMerge" => {
                    let val = xml::optional_attr_str(e, b"w:val")?;
                    props.vertical_merge = Some(match val.as_deref() {
                        Some("restart") => MergeType::Restart,
                        _ => MergeType::Continue,
                    });
                },
                b"gridSpan" => {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                        props.grid_span = val.parse().ok();
                    }
                },
                b"shd" => {
                    props.shading = Some(Shading {
                        fill: xml::optional_attr_str(e, b"w:fill")?.map(|v| v.into_owned()),
                        color: xml::optional_attr_str(e, b"w:color")?.map(|v| v.into_owned()),
                        pattern: xml::optional_attr_str(e, b"w:val")?.map(|v| v.into_owned()),
                    });
                },
                _ => {},
            },
            Event::End(ref e) if e.local_name().as_ref() == b"tcPr" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(props)
}

fn parse_table_width(e: &quick_xml::events::BytesStart) -> CoreResult<Option<TableWidth>> {
    let w = xml::optional_attr_str(e, b"w:w")?;
    let t = xml::optional_attr_str(e, b"w:type")?;

    if let Some(ref w_val) = w {
        let value: i32 = w_val.parse().unwrap_or(0);
        let width_type = match t.as_deref() {
            Some("pct") => TableWidthType::Pct,
            Some("dxa") => TableWidthType::Dxa,
            Some("auto") => TableWidthType::Auto,
            Some("nil") => TableWidthType::Nil,
            _ => TableWidthType::Dxa,
        };
        Ok(Some(TableWidth { value, width_type }))
    } else {
        Ok(None)
    }
}

// ---------------------------------------------------------------------------
// Section properties parsing
// ---------------------------------------------------------------------------

pub(crate) fn parse_section_properties(
    reader: &mut quick_xml::Reader<&[u8]>,
    _start: &quick_xml::events::BytesStart,
) -> CoreResult<SectionProperties> {
    let mut props = SectionProperties::default();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) => match e.local_name().as_ref() {
                b"pgSz" => {
                    let w: i32 = xml::optional_attr_str(e, b"w:w")?
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(12240);
                    let h: i32 = xml::optional_attr_str(e, b"w:h")?
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(15840);
                    let orient =
                        xml::optional_attr_str(e, b"w:orient")?.map(|v| match v.as_ref() {
                            "landscape" => PageOrientation::Landscape,
                            _ => PageOrientation::Portrait,
                        });
                    props.page_size = Some(PageSize {
                        width: crate::core::units::Twip(w),
                        height: crate::core::units::Twip(h),
                        orient,
                    });
                },
                b"pgMar" => {
                    props.margins = Some(PageMargins {
                        top: crate::core::units::Twip(
                            xml::optional_attr_str(e, b"w:top")?
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(1440),
                        ),
                        bottom: crate::core::units::Twip(
                            xml::optional_attr_str(e, b"w:bottom")?
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(1440),
                        ),
                        left: crate::core::units::Twip(
                            xml::optional_attr_str(e, b"w:left")?
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(1440),
                        ),
                        right: crate::core::units::Twip(
                            xml::optional_attr_str(e, b"w:right")?
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(1440),
                        ),
                        header: xml::optional_attr_str(e, b"w:header")?
                            .and_then(|v| v.parse().ok())
                            .map(crate::core::units::Twip),
                        footer: xml::optional_attr_str(e, b"w:footer")?
                            .and_then(|v| v.parse().ok())
                            .map(crate::core::units::Twip),
                        gutter: xml::optional_attr_str(e, b"w:gutter")?
                            .and_then(|v| v.parse().ok())
                            .map(crate::core::units::Twip),
                    });
                },
                b"headerReference" => {
                    let hf_type = parse_hf_type(e)?;
                    if let Ok(Some(rid)) = xml::optional_attr_str(e, b"r:id") {
                        props.header_refs.push(HeaderFooterRef {
                            hf_type,
                            relationship_id: rid.into_owned(),
                        });
                    }
                },
                b"footerReference" => {
                    let hf_type = parse_hf_type(e)?;
                    if let Ok(Some(rid)) = xml::optional_attr_str(e, b"r:id") {
                        props.footer_refs.push(HeaderFooterRef {
                            hf_type,
                            relationship_id: rid.into_owned(),
                        });
                    }
                },
                b"cols" => {
                    if let Ok(Some(num)) = xml::optional_attr_str(e, b"w:num") {
                        props.columns = num.parse().ok();
                    }
                },
                _ => {},
            },
            Event::End(ref e) if e.local_name().as_ref() == b"sectPr" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(props)
}

/// Recover the original face name from an embedded-font filename
/// produced by `core::embedded_fonts::write_embedded_fonts`. The
/// writer ships fonts as `font_<n>_<face>.<ext>` where `<face>` is
/// the original face name (with `/`, `?`, `*` etc. sanitized to `_`
/// — but NOT alphabetic characters, which earlier callers' naive
/// `trim_end_matches(alphabetic)` was greedily eating).
///
/// Examples:
///   `font_4_TeXGyreTermesX-Regular.ttf` → `TeXGyreTermesX-Regular`
///   `font_1_NewTXBMI.ttf`               → `NewTXBMI`
///   `font.otf`                          → `` (caller falls back to basename)
pub(crate) fn strip_embedded_font_filename(basename: &str) -> String {
    // Drop extension.
    let stem = match basename.rfind('.') {
        Some(i) => &basename[..i],
        None => basename,
    };
    // Strip the `font_<digits>_` prefix when present.
    if let Some(rest) = stem.strip_prefix("font_") {
        if let Some(under_idx) = rest.find('_') {
            // Everything before the underscore must be digits;
            // otherwise treat the whole stem as the face name.
            if rest[..under_idx].chars().all(|c| c.is_ascii_digit()) {
                return rest[under_idx + 1..].to_string();
            }
        }
    }
    stem.to_string()
}

fn parse_hf_type(e: &quick_xml::events::BytesStart) -> CoreResult<HeaderFooterType> {
    Ok(match xml::optional_attr_str(e, b"w:type")? {
        Some(ref val) => match val.as_ref() {
            "first" => HeaderFooterType::First,
            "even" => HeaderFooterType::Even,
            _ => HeaderFooterType::Default,
        },
        None => HeaderFooterType::Default,
    })
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

impl crate::core::OfficeDocument for DocxDocument {
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
    use std::io::Cursor;

    use crate::core::opc::{OpcWriter, PartName};

    fn make_minimal_docx(document_xml: &[u8]) -> Vec<u8> {
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut writer = OpcWriter::new(cursor).unwrap();

        let doc_part = PartName::new("/word/document.xml").unwrap();
        writer
            .add_part(
                &doc_part,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                document_xml,
            )
            .unwrap();
        writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");

        let result = writer.finish().unwrap();
        result.into_inner()
    }

    #[test]
    fn parse_empty_document() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body/>
</w:document>"#;
        let data = make_minimal_docx(xml);
        let doc = DocxDocument::from_reader(Cursor::new(data)).unwrap();
        assert!(doc.body.elements.is_empty());
        assert_eq!(doc.plain_text(), "");
    }

    #[test]
    fn parse_single_paragraph() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello, World!</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;
        let data = make_minimal_docx(xml);
        let doc = DocxDocument::from_reader(Cursor::new(data)).unwrap();
        assert_eq!(doc.body.elements.len(), 1);
        assert_eq!(doc.plain_text(), "Hello, World!");
    }

    #[test]
    fn parse_multiple_paragraphs() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>First paragraph.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Second paragraph.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;
        let data = make_minimal_docx(xml);
        let doc = DocxDocument::from_reader(Cursor::new(data)).unwrap();
        assert_eq!(doc.body.elements.len(), 2);
        assert_eq!(doc.plain_text(), "First paragraph.\nSecond paragraph.");
    }

    #[test]
    fn parse_multiple_runs() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">Hello </w:t></w:r>
      <w:r><w:t>World</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;
        let data = make_minimal_docx(xml);
        let doc = DocxDocument::from_reader(Cursor::new(data)).unwrap();
        assert_eq!(doc.plain_text(), "Hello World");
    }

    #[test]
    fn parse_break_and_tab() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Before</w:t>
        <w:tab/>
        <w:t>After</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;
        let data = make_minimal_docx(xml);
        let doc = DocxDocument::from_reader(Cursor::new(data)).unwrap();
        assert_eq!(doc.plain_text(), "Before\tAfter");
    }

    #[test]
    fn parse_table_basic() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#;
        let data = make_minimal_docx(xml);
        let doc = DocxDocument::from_reader(Cursor::new(data)).unwrap();
        assert_eq!(doc.body.elements.len(), 1);
        if let BlockElement::Table(ref table) = doc.body.elements[0] {
            assert_eq!(table.rows.len(), 2);
            assert_eq!(table.rows[0].cells.len(), 2);
        } else {
            panic!("expected table");
        }
        assert_eq!(doc.plain_text(), "A1\tB1\nA2\tB2");
    }

    #[test]
    fn parse_paragraph_with_formatting() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:jc w:val="center"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:sz w:val="32"/>
        </w:rPr>
        <w:t>Bold Heading</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;
        let data = make_minimal_docx(xml);
        let doc = DocxDocument::from_reader(Cursor::new(data)).unwrap();

        if let BlockElement::Paragraph(ref p) = doc.body.elements[0] {
            let pp = p.properties.as_ref().unwrap();
            assert_eq!(pp.style_id.as_deref(), Some("Heading1"));
            assert_eq!(pp.justification, Some(Justification::Center));

            if let ParagraphContent::Run(ref run) = p.content[0] {
                let rp = run.properties.as_ref().unwrap();
                assert_eq!(rp.bold, Some(true));
                assert_eq!(rp.font_size, Some(crate::core::units::HalfPoint(32)));
            } else {
                panic!("expected run");
            }
        } else {
            panic!("expected paragraph");
        }
    }

    #[test]
    fn markdown_bold_italic() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>bold</w:t>
      </w:r>
      <w:r>
        <w:t xml:space="preserve"> and </w:t>
      </w:r>
      <w:r>
        <w:rPr><w:i/></w:rPr>
        <w:t>italic</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;
        let data = make_minimal_docx(xml);
        let doc = DocxDocument::from_reader(Cursor::new(data)).unwrap();
        assert_eq!(doc.to_markdown(), "**bold** and *italic*");
    }

    #[test]
    fn markdown_table() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Header1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Header2</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Cell1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Cell2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#;
        let data = make_minimal_docx(xml);
        let doc = DocxDocument::from_reader(Cursor::new(data)).unwrap();
        let md = doc.to_markdown();
        assert!(md.contains("| Header1 | Header2 |"));
        assert!(md.contains("| --- | --- |"));
        assert!(md.contains("| Cell1 | Cell2 |"));
    }

    #[test]
    fn parse_drawing_anchor_position() {
        let xml =
            br#"<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <wp:anchor>
                <wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>
                <wp:positionV relativeFrom="page"><wp:posOffset>457200</wp:posOffset></wp:positionV>
                <wp:extent cx="2000000" cy="1500000"/>
                <a:graphic><a:graphicData uri="">
                    <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill>
                    </pic:pic>
                </a:graphicData></a:graphic>
            </wp:anchor>
        </w:drawing>"#;
        let mut reader = make_content_reader(xml);
        // Advance past the outer <w:drawing> Start so parse_drawing
        // sees the inner contents (it expects to be entered with
        // depth=1 already accounting for that wrapper).
        loop {
            match reader.read_event().unwrap() {
                quick_xml::events::Event::Start(ref e) if e.local_name().as_ref() == b"drawing" => {
                    break;
                },
                quick_xml::events::Event::Eof => panic!("no drawing"),
                _ => {},
            }
        }
        let info = parse_drawing(&mut reader).unwrap().expect("drawing");
        assert!(!info.inline);
        let pos = info.anchor_position.expect("anchor position");
        assert_eq!(pos.x_emu, 914400);
        assert_eq!(pos.y_emu, 457200);
        assert_eq!(pos.h_relative_from, crate::docx::AnchorFrame::Page);
        assert_eq!(info.relationship_id, "rId7");
    }

    #[test]
    fn parse_drawing_wsp_line_shape() {
        let xml =
            br#"<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wp:anchor>
                <wp:positionH relativeFrom="page"><wp:posOffset>100000</wp:posOffset></wp:positionH>
                <wp:positionV relativeFrom="page"><wp:posOffset>200000</wp:posOffset></wp:positionV>
                <wp:extent cx="500000" cy="0"/>
                <a:graphic><a:graphicData>
                    <wps:wsp>
                        <wps:spPr>
                            <a:prstGeom prst="line"/>
                            <a:ln w="9525">
                                <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                            </a:ln>
                        </wps:spPr>
                    </wps:wsp>
                </a:graphicData></a:graphic>
            </wp:anchor>
        </w:drawing>"#;
        let mut reader = make_content_reader(xml);
        loop {
            match reader.read_event().unwrap() {
                quick_xml::events::Event::Start(ref e) if e.local_name().as_ref() == b"drawing" => {
                    break;
                },
                quick_xml::events::Event::Eof => panic!("no drawing"),
                _ => {},
            }
        }
        let info = parse_drawing(&mut reader).unwrap().expect("drawing");
        let shape = info.shape.expect("shape");
        assert_eq!(shape.kind, crate::docx::ShapeKind::Line);
        assert_eq!(shape.stroke_rgb, Some((0xFF, 0x00, 0x00)));
        assert_eq!(shape.stroke_w_emu, Some(9525));
    }

    #[test]
    fn section_properties() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Content</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/>
    </w:sectPr>
  </w:body>
</w:document>"#;
        let data = make_minimal_docx(xml);
        let doc = DocxDocument::from_reader(Cursor::new(data)).unwrap();
        assert_eq!(doc.sections.len(), 1);
        let sect = &doc.sections[0];
        let ps = sect.page_size.as_ref().unwrap();
        assert_eq!(ps.width.0, 12240);
        assert_eq!(ps.height.0, 15840);
        let margins = sect.margins.as_ref().unwrap();
        assert_eq!(margins.left.0, 1800);
    }

    // ── strip_embedded_font_filename ────────────────────────────────────

    #[test]
    fn strip_embedded_font_writer_convention() {
        // Writer convention: font_<n>_<face>.<ext>
        assert_eq!(
            strip_embedded_font_filename("font_4_TeXGyreTermesX-Regular.ttf"),
            "TeXGyreTermesX-Regular"
        );
        assert_eq!(strip_embedded_font_filename("font_1_NewTXBMI.ttf"), "NewTXBMI");
        assert_eq!(strip_embedded_font_filename("font_12_DejaVuSans.otf"), "DejaVuSans");
    }

    #[test]
    fn strip_embedded_font_no_prefix_keeps_stem() {
        // No `font_<n>_` prefix → return the stem unchanged.
        assert_eq!(strip_embedded_font_filename("Arial.ttf"), "Arial");
        assert_eq!(strip_embedded_font_filename("MyFont.otf"), "MyFont");
    }

    #[test]
    fn strip_embedded_font_no_extension() {
        // No extension → use the whole input.
        assert_eq!(strip_embedded_font_filename("font_1_Calibri"), "Calibri");
        assert_eq!(strip_embedded_font_filename("Calibri"), "Calibri");
    }

    #[test]
    fn strip_embedded_font_non_digit_prefix_keeps_stem() {
        // `font_xxx_<face>` where xxx isn't digits → don't strip.
        assert_eq!(strip_embedded_font_filename("font_abc_Foo.ttf"), "font_abc_Foo");
    }

    #[test]
    fn strip_embedded_font_alphabetic_face_preserved() {
        // Regression: greedy trim_end_matches(alphabetic) used to eat
        // the face name. Verify a face with trailing alphabetic chars
        // survives intact.
        assert_eq!(
            strip_embedded_font_filename("font_4_TeXGyreTermesX-Bold.ttf"),
            "TeXGyreTermesX-Bold"
        );
    }

    #[test]
    fn strip_embedded_font_empty() {
        assert_eq!(strip_embedded_font_filename(""), "");
    }

    #[test]
    fn strip_embedded_font_no_face_after_prefix() {
        // `font_<n>_` with nothing after the underscore → empty face.
        // Caller of this helper falls back to the full basename.
        assert_eq!(strip_embedded_font_filename("font_5_.ttf"), "");
    }
}
