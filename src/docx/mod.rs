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

pub mod document;
pub mod error;
pub mod formatting;
pub mod headers;
pub mod hyperlink;
pub mod image;
pub mod numbering;
pub mod paragraph;
pub mod styles;
pub mod table;
pub mod text;
pub mod write;
pub mod edit;

pub use document::{BlockElement, Body};
pub use error::{DocxError, Result};
pub use formatting::{
    Justification, ParagraphIndent, ParagraphProperties, ParagraphSpacing, RunProperties,
    UnderlineType, VerticalAlign,
};
pub use headers::{HeaderFooter, HeaderFooterType, PageMargins, PageSize, SectionProperties};
pub use hyperlink::{Hyperlink, HyperlinkTarget};
pub use image::DrawingInfo;
pub use numbering::{NumberFormat, NumberingDefinitions};
pub use paragraph::{BreakType, Paragraph, ParagraphContent, Run, RunContent};
pub use styles::{Style, StyleSheet, StyleType};
pub use table::{Table, TableCell, TableProperties, TableRow};

use std::io::{Read, Seek};
use std::path::Path;

use log::debug;
use quick_xml::events::Event;

use crate::core::opc::OpcReader;
use crate::core::relationships::{rel_types, TargetMode};
use crate::core::theme::Theme;
use crate::core::units::Emu;
use crate::core::xml;

use self::formatting::{parse_paragraph_properties_fast, parse_run_properties_fast};
use self::headers::{HeaderFooterRef, PageOrientation};
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

        // Parse headers and footers
        let mut headers_footers = Vec::new();
        for section in &sections {
            for hf_ref in section.header_refs.iter().chain(section.footer_refs.iter()) {
                if let Some(rel) = doc_rels.get_by_id(&hf_ref.relationship_id) {
                    if rel.target_mode == TargetMode::Internal {
                        let part_name = main_part.resolve_relative(&rel.target)?;
                        if opc.has_part(&part_name) {
                            let data = opc.read_part(&part_name)?;
                            let content = parse_body_elements(&data)?;
                            headers_footers.push(HeaderFooter {
                                hf_type: hf_ref.hf_type,
                                content,
                            });
                        }
                    }
                }
            }
        }

        debug!(
            "DocxDocument: {} block elements, {} sections",
            body.elements.len(),
            sections.len()
        );
        Ok(DocxDocument {
            body,
            styles,
            numbering,
            theme,
            sections,
            headers_footers,
        })
    }
}

/// Parse body-level elements from XML (used for headers/footers which share the same structure).
fn parse_body_elements(xml_data: &[u8]) -> CoreResult<Vec<BlockElement>> {
    let mut reader = make_content_reader(xml_data);
    let mut elements = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                match e.local_name().as_ref() {
                    b"p" => {
                        elements.push(BlockElement::Paragraph(parse_paragraph(&mut reader)?));
                    }
                    b"tbl" => {
                        elements.push(BlockElement::Table(parse_table(&mut reader)?));
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
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
            Event::Start(ref e) => {
                match e.local_name().as_ref() {
                    b"body" => {
                        in_body = true;
                    }
                    b"p" if in_body => {
                        elements.push(BlockElement::Paragraph(parse_paragraph(&mut reader)?));
                    }
                    b"tbl" if in_body => {
                        elements.push(BlockElement::Table(parse_table(&mut reader)?));
                    }
                    b"sectPr" if in_body => {
                        sections.push(parse_section_properties(&mut reader, e)?);
                    }
                    _ => {}
                }
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"body" {
                    in_body = false;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // Resolve hyperlink targets using relationships
    resolve_hyperlinks(&mut elements, rels);

    let body = Body { elements };
    Ok((body, sections))
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
            }
            BlockElement::Table(t) => {
                for row in &mut t.rows {
                    for cell in &mut row.cells {
                        resolve_hyperlinks(&mut cell.content, rels);
                    }
                }
            }
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
            Event::Start(ref e) => {
                match e.local_name().as_ref() {
                    b"pPr" => {
                        paragraph.properties = Some(parse_paragraph_properties_fast(reader)?);
                    }
                    b"r" => {
                        paragraph
                            .content
                            .push(ParagraphContent::Run(parse_run(reader)?));
                    }
                    b"hyperlink" => {
                        paragraph.content.push(ParagraphContent::Hyperlink(
                            parse_hyperlink(reader, e)?,
                        ));
                    }
                    _ => {
                        xml::skip_element_fast(reader)?;
                    }
                }
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"p" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(paragraph)
}

fn parse_run(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<Run> {
    let mut run = Run::default();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                match e.local_name().as_ref() {
                    b"rPr" => {
                        run.properties = Some(parse_run_properties_fast(reader)?);
                    }
                    b"t" => {
                        let text = xml::read_text_content_fast(reader)?;
                        if !text.is_empty() {
                            run.content.push(RunContent::Text(text));
                        }
                    }
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
                    }
                    b"drawing" => {
                        if let Some(drawing) = parse_drawing(reader)? {
                            run.content.push(RunContent::Drawing(drawing));
                        }
                    }
                    _ => {
                        xml::skip_element_fast(reader)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                match e.local_name().as_ref() {
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
                    }
                    b"tab" => {
                        run.content.push(RunContent::Tab);
                    }
                    _ => {}
                }
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"r" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
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
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"hyperlink" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
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

fn parse_drawing(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<Option<DrawingInfo>> {
    let mut inline = true;
    let mut width = Emu(0);
    let mut height = Emu(0);
    let mut description: Option<String> = None;
    let mut relationship_id: Option<String> = None;
    let mut depth = 1u32;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                depth += 1;
                let local = e.local_name();
                let local_bytes = local.as_ref();
                match local_bytes {
                    b"inline" => inline = true,
                    b"anchor" => inline = false,
                    b"extent" => parse_extent_attrs(e, &mut width, &mut height),
                    b"docPr" => {
                        if let Ok(Some(desc)) = xml::optional_attr_str(e, b"descr") {
                            description = Some(desc.into_owned());
                        }
                    }
                    b"blip" => {
                        if let Ok(Some(embed)) = xml::optional_attr_str(e, b"r:embed") {
                            relationship_id = Some(embed.into_owned());
                        }
                    }
                    _ => {}
                }
            }
            Event::Empty(ref e) => {
                let local = e.local_name();
                let local_bytes = local.as_ref();
                match local_bytes {
                    b"extent" => parse_extent_attrs(e, &mut width, &mut height),
                    b"docPr" => {
                        if let Ok(Some(desc)) = xml::optional_attr_str(e, b"descr") {
                            description = Some(desc.into_owned());
                        }
                    }
                    b"blip" => {
                        if let Ok(Some(embed)) = xml::optional_attr_str(e, b"r:embed") {
                            relationship_id = Some(embed.into_owned());
                        }
                    }
                    _ => {}
                }
            }
            Event::End(_) => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    if let Some(rid) = relationship_id {
        Ok(Some(DrawingInfo {
            relationship_id: rid,
            description,
            width,
            height,
            inline,
        }))
    } else {
        Ok(None)
    }
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
            Event::Start(ref e) => {
                match e.local_name().as_ref() {
                    b"tblPr" => {
                        properties = Some(parse_table_properties(reader)?);
                    }
                    b"tblGrid" => {
                        grid = parse_table_grid(reader)?;
                    }
                    b"tr" => {
                        rows.push(parse_table_row(reader)?);
                    }
                    _ => {
                        xml::skip_element_fast(reader)?;
                    }
                }
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"tbl" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
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
            Event::Start(ref e) => {
                match e.local_name().as_ref() {
                    b"tblW" => {
                        props.width = parse_table_width(e)?;
                        xml::skip_element_fast(reader)?;
                    }
                    b"jc" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.justification =
                                Some(self::formatting::parse_justification_value(&val));
                        }
                        xml::skip_element_fast(reader)?;
                    }
                    b"tblStyle" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.style_id = Some(val.into_owned());
                        }
                        xml::skip_element_fast(reader)?;
                    }
                    _ => {
                        xml::skip_element_fast(reader)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                match e.local_name().as_ref() {
                    b"tblW" => {
                        props.width = parse_table_width(e)?;
                    }
                    b"jc" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.justification =
                                Some(self::formatting::parse_justification_value(&val));
                        }
                    }
                    b"tblStyle" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.style_id = Some(val.into_owned());
                        }
                    }
                    _ => {}
                }
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"tblPr" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(props)
}

fn parse_table_grid(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<Vec<crate::core::units::Twip>> {
    let mut cols = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) => {
                if e.local_name().as_ref() == b"gridCol" {
                    if let Ok(Some(w)) = xml::optional_attr_str(e, b"w:w") {
                        let val: i32 = w.parse().unwrap_or(0);
                        cols.push(crate::core::units::Twip(val));
                    }
                }
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"tblGrid" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(cols)
}

fn parse_table_row(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<TableRow> {
    let mut properties = None;
    let mut cells = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                match e.local_name().as_ref() {
                    b"trPr" => {
                        properties = Some(parse_table_row_properties(reader)?);
                    }
                    b"tc" => {
                        cells.push(parse_table_cell(reader)?);
                    }
                    _ => {
                        xml::skip_element_fast(reader)?;
                    }
                }
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"tr" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(TableRow { properties, cells })
}

fn parse_table_row_properties(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<TableRowProperties> {
    let mut props = TableRowProperties::default();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) => {
                if e.local_name().as_ref() == b"tblHeader" {
                    props.is_header = true;
                }
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"trPr" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(props)
}

fn parse_table_cell(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<TableCell> {
    let mut properties = None;
    let mut content = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                match e.local_name().as_ref() {
                    b"tcPr" => {
                        properties = Some(parse_table_cell_properties(reader)?);
                    }
                    b"p" => {
                        content.push(BlockElement::Paragraph(parse_paragraph(reader)?));
                    }
                    b"tbl" => {
                        content.push(BlockElement::Table(parse_table(reader)?));
                    }
                    _ => {
                        xml::skip_element_fast(reader)?;
                    }
                }
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"tc" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
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
            Event::Start(ref e) => {
                match e.local_name().as_ref() {
                    b"tcW" => {
                        props.width = parse_table_width(e)?;
                        xml::skip_element_fast(reader)?;
                    }
                    b"vMerge" => {
                        let val = xml::optional_attr_str(e, b"w:val")?;
                        props.vertical_merge = Some(match val.as_deref() {
                            Some("restart") => MergeType::Restart,
                            _ => MergeType::Continue,
                        });
                        xml::skip_element_fast(reader)?;
                    }
                    b"gridSpan" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.grid_span = val.parse().ok();
                        }
                        xml::skip_element_fast(reader)?;
                    }
                    b"shd" => {
                        props.shading = Some(Shading {
                            fill: xml::optional_attr_str(e, b"w:fill")?
                                .map(|v| v.into_owned()),
                            color: xml::optional_attr_str(e, b"w:color")?
                                .map(|v| v.into_owned()),
                            pattern: xml::optional_attr_str(e, b"w:val")?
                                .map(|v| v.into_owned()),
                        });
                        xml::skip_element_fast(reader)?;
                    }
                    _ => {
                        xml::skip_element_fast(reader)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                match e.local_name().as_ref() {
                    b"tcW" => {
                        props.width = parse_table_width(e)?;
                    }
                    b"vMerge" => {
                        let val = xml::optional_attr_str(e, b"w:val")?;
                        props.vertical_merge = Some(match val.as_deref() {
                            Some("restart") => MergeType::Restart,
                            _ => MergeType::Continue,
                        });
                    }
                    b"gridSpan" => {
                        if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                            props.grid_span = val.parse().ok();
                        }
                    }
                    b"shd" => {
                        props.shading = Some(Shading {
                            fill: xml::optional_attr_str(e, b"w:fill")?
                                .map(|v| v.into_owned()),
                            color: xml::optional_attr_str(e, b"w:color")?
                                .map(|v| v.into_owned()),
                            pattern: xml::optional_attr_str(e, b"w:val")?
                                .map(|v| v.into_owned()),
                        });
                    }
                    _ => {}
                }
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"tcPr" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
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

fn parse_section_properties(
    reader: &mut quick_xml::Reader<&[u8]>,
    _start: &quick_xml::events::BytesStart,
) -> CoreResult<SectionProperties> {
    let mut props = SectionProperties::default();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) => {
                match e.local_name().as_ref() {
                        b"pgSz" => {
                            let w: i32 = xml::optional_attr_str(e, b"w:w")?
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(12240);
                            let h: i32 = xml::optional_attr_str(e, b"w:h")?
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(15840);
                            let orient = xml::optional_attr_str(e, b"w:orient")?.map(|v| {
                                match v.as_ref() {
                                    "landscape" => PageOrientation::Landscape,
                                    _ => PageOrientation::Portrait,
                                }
                            });
                            props.page_size = Some(PageSize {
                                width: crate::core::units::Twip(w),
                                height: crate::core::units::Twip(h),
                                orient,
                            });
                        }
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
                        }
                        b"headerReference" => {
                            let hf_type = parse_hf_type(e)?;
                            if let Ok(Some(rid)) = xml::optional_attr_str(e, b"r:id") {
                                props.header_refs.push(HeaderFooterRef {
                                    hf_type,
                                    relationship_id: rid.into_owned(),
                                });
                            }
                        }
                        b"footerReference" => {
                            let hf_type = parse_hf_type(e)?;
                            if let Ok(Some(rid)) = xml::optional_attr_str(e, b"r:id") {
                                props.footer_refs.push(HeaderFooterRef {
                                    hf_type,
                                    relationship_id: rid.into_owned(),
                                });
                            }
                        }
                        b"cols" => {
                            if let Ok(Some(num)) = xml::optional_attr_str(e, b"w:num") {
                                props.columns = num.parse().ok();
                            }
                        }
                        _ => {}
                    }
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"sectPr" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(props)
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
}
