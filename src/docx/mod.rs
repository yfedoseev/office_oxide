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
    BorderEdge, Justification, LineSpacingRule, ParagraphBorders, ParagraphIndent,
    ParagraphProperties, ParagraphSpacing, RunProperties, TabStopDef, TableBorders, UnderlineType,
    VerticalAlign,
};
pub use headers::{
    ColumnDefs, HeaderFooter, HeaderFooterType, PageMargins, PageOrientation, PageSize,
    SectionBreakKind, SectionProperties,
};
pub use hyperlink::{Hyperlink, HyperlinkTarget};
pub use image::{AnchorFrame, AnchorPosition, DrawingInfo, ShapeInfo, ShapeKind};
pub use numbering::{NumberFormat, NumberingDefinitions};
pub use paragraph::{
    BreakType, FormField, FormFieldKind, Paragraph, ParagraphContent, Run, RunContent,
};
pub use styles::{Style, StyleSheet, StyleType};
pub use table::{
    CellMargins, CellVAlign, RowHeightRule, Table, TableCell, TableProperties, TableRow,
    TableWidth, TableWidthType,
};

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
use self::table::{MergeType, Shading, TableCellProperties, TableRowProperties};

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
    /// Parsed `docProps/core.xml`. `None` when the package carries no
    /// core-properties part.
    pub core_properties: Option<crate::core::properties::CoreProperties>,
    /// Footnote bodies from `word/footnotes.xml`, in document order.
    /// Separator/continuation pseudo-notes are filtered out.
    pub footnotes: Vec<NoteBody>,
    /// Endnote bodies from `word/endnotes.xml`.
    pub endnotes: Vec<NoteBody>,
    /// Comment bodies from `word/comments.xml`.
    pub comments: Vec<NoteBody>,
}

/// A footnote, endnote or comment body.
#[derive(Debug, Clone)]
pub struct NoteBody {
    /// The note's `w:id`.
    pub id: u32,
    /// Author, for comments (`w:author`); `None` for foot/endnotes.
    pub author: Option<String>,
    /// Block content of the note.
    pub content: Vec<BlockElement>,
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
        // Refuse a package whose primary part is not WordprocessingML — an
        // XLSX opened as a DOCX used to parse to an empty document with no
        // diagnostic at all.
        opc.verify_main_content_type(
            &[
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
                "application/vnd.ms-word.document.macroEnabled.main+xml",
                "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
            ],
            "a WordprocessingML document",
        )?;
        let core_properties = crate::core::properties::read_core_properties(&mut opc);
        let main_part = opc.main_document_part()?;
        let doc_rels = opc.read_rels_for(&main_part)?;

        // Parse theme
        // A theme is decoration: it supplies colour-scheme lookups and
        // nothing more. A malformed or missing theme part used to fail the
        // whole open, so one bad ancillary part made an otherwise readable
        // document unreadable — while a *missing* theme was fine, which is
        // the inconsistency that gave it away.
        let theme = doc_rels
            .first_by_type(rel_types::THEME)
            .and_then(|rel| main_part.resolve_relative(&rel.target).ok())
            .filter(|pn| opc.has_part(pn))
            .and_then(|pn| opc.read_part(&pn).ok())
            .and_then(|data| match Theme::parse(&data) {
                Ok(t) => Some(t),
                Err(e) => {
                    debug!("DocxDocument: ignoring unreadable theme part: {e}");
                    None
                },
            });

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
        // A part cut off mid-element parses to whatever was read before the
        // cut and reports success; check the root actually closed.
        xml::check_root_closed(&doc_data, main_part.as_str(), "document")?;
        let (mut body, sections, alt_chunk_rids, alt_chunk_positions) =
            parse_document(&doc_data, &doc_rels)?;

        // Resolve each `<w:altChunk>` and splice its text in at the point
        // the document references it. HTML and plain-text chunks are read;
        // a nested package chunk is left alone rather than guessed at.
        if !alt_chunk_rids.is_empty() {
            let mut inserts: Vec<(usize, Vec<BlockElement>)> = Vec::new();
            for (rid, pos) in alt_chunk_rids.iter().zip(alt_chunk_positions.iter()) {
                let Some(rel) = doc_rels.get_by_id(rid) else {
                    continue;
                };
                if rel.target_mode != TargetMode::Internal {
                    continue;
                }
                let Ok(part) = main_part.resolve_relative(&rel.target) else {
                    continue;
                };
                if !opc.has_part(&part) {
                    continue;
                }
                let Ok(data) = opc.read_part(&part) else {
                    continue;
                };
                let paras = parse_alt_chunk(&data, part.as_str());
                if !paras.is_empty() {
                    inserts.push((*pos, paras));
                }
            }
            // Splice back-to-front so earlier indices stay valid.
            inserts.sort_by_key(|(pos, _)| std::cmp::Reverse(*pos));
            for (pos, paras) in inserts {
                let at = pos.min(body.elements.len());
                body.elements.splice(at..at, paras);
            }
        }

        // Resolve every native DrawingML chart the body references. A
        // `<w:drawing>` holding a chart carries only `<c:chart r:id="…"/>`;
        // the chart's title, axis titles, category labels, series names and
        // cached data values all live in the separate part that id resolves
        // to (`word/charts/chartN.xml`), which was never opened, so none of
        // that text reached any consumer (issue #273).
        let mut chart_text_by_rid: std::collections::HashMap<String, Vec<String>> =
            std::collections::HashMap::new();
        for rel in doc_rels.all() {
            if rel.rel_type != rel_types::CHART || rel.target_mode != TargetMode::Internal {
                continue;
            }
            let Ok(part) = main_part.resolve_relative(&rel.target) else {
                continue;
            };
            if !opc.has_part(&part) {
                continue;
            }
            let Ok(data) = opc.read_part(&part) else {
                continue;
            };
            let lines = crate::core::chart::chart_text_lines(&data);
            if !lines.is_empty() {
                chart_text_by_rid.insert(rel.id.clone(), lines);
            }
        }
        // Resolve every SmartArt diagram the body references, the same
        // way charts are resolved just above (issue #271).
        let mut dgm_text_by_rid: std::collections::HashMap<String, Vec<String>> =
            std::collections::HashMap::new();
        for rel in doc_rels.all() {
            if rel.rel_type != rel_types::DIAGRAM_DATA || rel.target_mode != TargetMode::Internal {
                continue;
            }
            let Ok(part) = main_part.resolve_relative(&rel.target) else {
                continue;
            };
            if !opc.has_part(&part) {
                continue;
            }
            let Ok(data) = opc.read_part(&part) else {
                continue;
            };
            let lines = diagram_text_lines(&data);
            if !lines.is_empty() {
                dgm_text_by_rid.insert(rel.id.clone(), lines);
            }
        }
        if !chart_text_by_rid.is_empty() || !dgm_text_by_rid.is_empty() {
            attach_chart_text(&mut body.elements, &chart_text_by_rid, &dgm_text_by_rid);
        }

        // Resolve every embedded native OOXML package object
        // (`<o:OLEObject Type="Embed">`, issue #304) by opening its bytes
        // with the appropriate format reader and folding in its text.
        resolve_deferred_parts(&mut body.elements, &mut opc, &main_part, &doc_rels);

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

        // Footnotes, endnotes and comments are separate parts referenced
        // from the document relationships. They were never read, which made
        // `ir::Element::Footnote` unreachable and dropped every note body.
        let mut read_notes = |rel_type: &str, end: &[u8]| -> Vec<NoteBody> {
            let Some(rel) = doc_rels.first_by_type(rel_type) else {
                return Vec::new();
            };
            let Ok(part) = main_part.resolve_relative(&rel.target) else {
                return Vec::new();
            };
            if !opc.has_part(&part) {
                return Vec::new();
            }
            let Ok(data) = opc.read_part(&part) else {
                return Vec::new();
            };
            let mut notes = parse_notes_part(&data, end).unwrap_or_default();
            // `r:id` is scoped per OPC part: a hyperlink inside a
            // footnote/endnote/comment resolves against *that* part's
            // `_rels`, not `document.xml.rels`. Resolving against the
            // document's relationships left the raw `rIdN` string as the
            // "URL" for every note hyperlink (issue #293).
            if let Ok(note_rels) = opc.read_rels_for(&part) {
                for n in &mut notes {
                    resolve_hyperlinks(&mut n.content, &note_rels);
                }
            }
            notes
        };
        let footnotes = read_notes(rel_types::FOOTNOTES, b"footnote");
        let endnotes = read_notes(rel_types::ENDNOTES, b"endnote");
        let comments = read_notes(rel_types::COMMENTS, b"comment");

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
            core_properties,
            footnotes,
            endnotes,
            comments,
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

/// Read a `<w:altChunk>` target part into block elements.
///
/// HTML and plain-text chunks are handled: both are trivial to read, and
/// together they cover what mail-merge and report generators emit. A nested
/// `.docx` package chunk is deliberately left unread rather than guessed
/// at — it returns no paragraphs, so the caller inserts nothing.
fn parse_alt_chunk(data: &[u8], part_name: &str) -> Vec<BlockElement> {
    let lower = part_name.to_ascii_lowercase();
    let text = if lower.ends_with(".html") || lower.ends_with(".htm") || lower.ends_with(".xhtml") {
        strip_html_tags(&String::from_utf8_lossy(data))
    } else if lower.ends_with(".txt") {
        String::from_utf8_lossy(data).into_owned()
    } else {
        // A `.docx` chunk is a whole nested package; not supported.
        return Vec::new();
    };

    text.lines()
        .map(str::trim)
        .filter(|l| !l.is_empty())
        .map(|line| {
            BlockElement::Paragraph(Paragraph {
                properties: None,
                content: vec![ParagraphContent::Run(Run {
                    properties: None,
                    content: vec![RunContent::Text(line.to_string())],
                })],
            })
        })
        .collect()
}

/// Flatten an HTML fragment to text: drop tags and `<script>`/`<style>`
/// bodies, resolve the handful of entities that matter, and put a line
/// break where a block element ends.
fn strip_html_tags(html: &str) -> String {
    let mut out = String::with_capacity(html.len());
    let mut chars = html.char_indices().peekable();
    let mut skip_until: Option<&str> = None;

    while let Some((i, c)) = chars.next() {
        if let Some(end) = skip_until {
            if html[i..].to_ascii_lowercase().starts_with(end) {
                for _ in 0..end.len() - 1 {
                    chars.next();
                }
                skip_until = None;
            }
            continue;
        }
        if c != '<' {
            out.push(c);
            continue;
        }
        let rest = html[i..].to_ascii_lowercase();
        if rest.starts_with("<script") {
            skip_until = Some("</script>");
            continue;
        }
        if rest.starts_with("<style") {
            skip_until = Some("</style>");
            continue;
        }
        // Consume through the closing '>'.
        let mut tag = String::new();
        for (_, tc) in chars.by_ref() {
            if tc == '>' {
                break;
            }
            tag.push(tc);
        }
        let name = tag
            .trim_start_matches('/')
            .split(|c: char| c.is_whitespace())
            .next()
            .unwrap_or("")
            .to_ascii_lowercase();
        if matches!(
            name.as_str(),
            "p" | "div" | "br" | "li" | "tr" | "h1" | "h2" | "h3" | "h4" | "h5" | "h6"
        ) {
            out.push('\n');
        }
    }

    // Resolve the entities an HTML chunk actually uses.
    out.replace("&nbsp;", " ")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&#39;", "'")
        .replace("&amp;", "&")
}

/// Parse a `footnotes.xml` / `endnotes.xml` / `comments.xml` part.
///
/// `end` is the per-part item element name (`footnote`, `endnote` or
/// `comment`). Word emits two pseudo-notes at ids 0 and -1 (the separator
/// and continuation marks) with `w:type` set; those are not document
/// content and are filtered out.
fn parse_notes_part(xml_data: &[u8], end: &[u8]) -> CoreResult<Vec<NoteBody>> {
    let mut reader = make_content_reader(xml_data);
    let mut notes = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == end => {
                let note_type = xml::optional_attr_str(e, b"w:type")?;
                let id: i64 = xml::optional_attr_str(e, b"w:id")?
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                let author = xml::optional_attr_str(e, b"w:author")?.map(|v| v.into_owned());
                let content = parse_block_elements_until(&mut reader, end)?;
                let is_pseudo = note_type.as_deref().is_some_and(|t| t != "normal") || id < 0;
                if !is_pseudo {
                    notes.push(NoteBody {
                        id: id.max(0) as u32,
                        author,
                        content,
                    });
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(notes)
}

/// Parse block elements (paragraphs and tables) until the matching
/// `</end_local>`. Used for `<w:txbxContent>` bodies, which hold ordinary
/// document content inside a shape.
fn parse_block_elements_until(
    reader: &mut quick_xml::Reader<&[u8]>,
    end_local: &[u8],
) -> CoreResult<Vec<BlockElement>> {
    let mut elements = Vec::new();
    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"p" => elements.push(BlockElement::Paragraph(parse_paragraph(reader)?)),
                b"tbl" => elements.push(BlockElement::Table(parse_table(reader)?)),
                _ => xml::skip_element_fast(reader)?,
            },
            Event::End(ref e) if e.local_name().as_ref() == end_local => break,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(elements)
}

/// Everything a legacy VML `<w:pict>` / `<w:object>` subtree can carry.
///
/// Only `<w:txbxContent>` used to be read, so a `w:pict` wrapping an
/// image (#268), WordArt (#274) or an embedded package (#304) contributed
/// nothing at all to the document.
#[derive(Default)]
struct VmlContent {
    /// `<w:txbxContent>` bodies (text boxes).
    boxes: Vec<Vec<BlockElement>>,
    /// `<v:imagedata>` relationship ids with the enclosing shape's size.
    images: Vec<(String, Emu, Emu)>,
    /// `<v:textpath string="…">` values (WordArt).
    wordart: Vec<String>,
    /// `<o:OLEObject Type="Embed">` relationship ids.
    ole_rids: Vec<String>,
}

impl VmlContent {
    fn is_empty(&self) -> bool {
        self.boxes.is_empty()
            && self.images.is_empty()
            && self.wordart.is_empty()
            && self.ole_rids.is_empty()
    }

    fn merge(&mut self, other: VmlContent) {
        self.boxes.extend(other.boxes);
        self.images.extend(other.images);
        self.wordart.extend(other.wordart);
        self.ole_rids.extend(other.ole_rids);
    }

    /// Append this payload to a run's content in a stable order: WordArt
    /// text, then images, then text boxes, then deferred packages.
    fn push_into(self, out: &mut Vec<RunContent>) {
        for s in self.wordart {
            out.push(RunContent::Text(s));
        }
        for (rid, w, h) in self.images {
            out.push(RunContent::Drawing(DrawingInfo {
                relationship_id: rid,
                description: None,
                width: w,
                height: h,
                inline: true,
                anchor_position: None,
                shape: None,
                chart_rel_id: None,
                chart_text: Vec::new(),
                dgm_data_rel_id: None,
                dgm_text: Vec::new(),
            }));
        }
        for b in self.boxes {
            out.push(RunContent::TextBox(b));
        }
        for rid in self.ole_rids {
            out.push(RunContent::DeferredPart(rid));
        }
    }
}

/// Read a VML `<v:shape style="width:191pt;height:88pt">` size. VML uses
/// CSS-ish lengths, so the unit has to be honoured.
fn vml_style_size(e: &quick_xml::events::BytesStart) -> Option<(Emu, Emu)> {
    let style = xml::optional_attr_str(e, b"style").ok()??;
    fn dim(style: &str, key: &str) -> Option<i64> {
        // Match `width:` but not `mso-wrap-width:`; a leading `;` or the
        // string start must precede it.
        let mut rest = style;
        loop {
            let at = rest.find(key)?;
            let ok = at == 0
                || rest[..at]
                    .chars()
                    .next_back()
                    .is_some_and(|c| c == ';' || c.is_whitespace());
            if ok {
                let v = rest[at + key.len()..].trim();
                let end = v
                    .find(|c: char| !(c.is_ascii_digit() || c == '.' || c == '-'))
                    .unwrap_or(v.len());
                let num: f64 = v[..end].parse().ok()?;
                let unit = v[end..].trim();
                // 1 pt = 12700 EMU; the rest convert through points.
                let emu_per = match unit.trim_end_matches(|c: char| c == ';' || c.is_whitespace()) {
                    "in" => 914_400.0,
                    "cm" => 360_000.0,
                    "mm" => 36_000.0,
                    "px" => 9525.0,
                    "pc" => 152_400.0,
                    _ => 12_700.0,
                };
                return Some((num * emu_per) as i64);
            }
            rest = &rest[at + key.len()..];
        }
    }
    let w = dim(&style, "width:").unwrap_or(0);
    let h = dim(&style, "height:").unwrap_or(0);
    if w == 0 && h == 0 {
        None
    } else {
        Some((Emu(w), Emu(h)))
    }
}

/// Collect the contents of a VML `<w:pict>` / `<w:object>` subtree,
/// reading through the matching closing tag.
fn parse_vml_content_in(
    reader: &mut quick_xml::Reader<&[u8]>,
    end_local: &[u8],
) -> CoreResult<VmlContent> {
    let mut out = VmlContent::default();
    let mut depth = 1i32;
    // `<v:imagedata>` is a child of `<v:shape>`, which is where the
    // display size lives.
    let mut shape_size = (Emu(0), Emu(0));

    loop {
        let (e, is_start) = match reader.read_event()? {
            Event::Start(e) => (e, true),
            Event::Empty(e) => (e, false),
            Event::End(ref e) => {
                if e.local_name().as_ref() == end_local && depth <= 1 {
                    break;
                }
                depth -= 1;
                if depth <= 0 {
                    break;
                }
                continue;
            },
            Event::Eof => break,
            _ => continue,
        };
        match e.local_name().as_ref() {
            b"txbxContent" if is_start => {
                out.boxes.push(parse_block_elements_until(reader, b"txbxContent")?);
                // The subtree is fully consumed, so depth is unchanged.
                continue;
            },
            b"shape" | b"rect" | b"roundrect" | b"oval" | b"line" | b"polyline" => {
                if let Some(sz) = vml_style_size(&e) {
                    shape_size = sz;
                }
            },
            b"imagedata" => {
                // Word writes `r:id`; some producers write `o:relid`.
                let rid = xml::optional_attr_str(&e, b"r:id")?
                    .or(xml::optional_attr_str(&e, b"o:relid")?)
                    .map(|v| v.into_owned());
                if let Some(rid) = rid.filter(|r| !r.is_empty()) {
                    out.images.push((rid, shape_size.0, shape_size.1));
                }
            },
            // WordArt keeps its text in an *attribute*, so neither the
            // text-box nor the run path ever reached it.
            b"textpath" => {
                if let Some(s) = xml::optional_attr_str(&e, b"string")? {
                    let s = s.into_owned();
                    if !s.trim().is_empty() {
                        out.wordart.push(s);
                    }
                }
            },
            b"OLEObject" => {
                let embedded = xml::optional_attr_str(&e, b"Type")?
                    .is_none_or(|t| t.eq_ignore_ascii_case("Embed"));
                let rid = xml::optional_attr_str(&e, b"r:id")?.map(|v| v.into_owned());
                if let Some(rid) = rid.filter(|_| embedded) {
                    out.ole_rids.push(rid);
                }
            },
            _ => {},
        }
        if is_start {
            depth += 1;
        }
    }
    Ok(out)
}

/// Parse `<mc:AlternateContent>`, taking the `<mc:Choice>` branch and
/// discarding `<mc:Fallback>`.
///
/// The two branches describe the *same* shape for different consumers.
/// Extracting both duplicated every text box's contents; extracting
/// neither dropped them. `<mc:Fallback>` is used only when no `<mc:Choice>`
/// yielded content.
fn parse_alternate_content(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<VmlContent> {
    let mut chosen = VmlContent::default();
    let mut fallback = VmlContent::default();
    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"Choice" => chosen.merge(parse_vml_content_in(reader, b"Choice")?),
                b"Fallback" => fallback.merge(parse_vml_content_in(reader, b"Fallback")?),
                _ => xml::skip_element_fast(reader)?,
            },
            Event::End(ref e) if e.local_name().as_ref() == b"AlternateContent" => break,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(if chosen.is_empty() { fallback } else { chosen })
}

/// Decode `<w:sym w:font="Wingdings" w:char="F0B7"/>` into a character.
///
/// The code point is hex. Values in `F000..F0FF` are the Windows
/// symbol-font private-use encoding of a single byte; map them down to the
/// bullet-ish equivalent rather than emitting an unrenderable PUA code
/// point, and pass everything else through unchanged.
fn parse_sym_char(e: &quick_xml::events::BytesStart) -> Option<char> {
    let raw = xml::optional_attr_str(e, b"w:char").ok().flatten()?;
    let code = u32::from_str_radix(raw.trim(), 16).ok()?;
    let mapped = match code {
        0xF0B7 | 0xF0A7 => 0x2022, // Wingdings/Symbol bullet
        0xF0D8 => 0x25BA,          // Wingdings arrowhead
        c @ 0xF000..=0xF0FF => c - 0xF000,
        c => c,
    };
    char::from_u32(mapped)
}

/// Parse `word/document.xml` and return the Body and SectionProperties.
type ParsedDocument = (Body, Vec<SectionProperties>, Vec<String>, Vec<usize>);

fn parse_document(
    xml_data: &[u8],
    rels: &crate::core::relationships::Relationships,
) -> CoreResult<ParsedDocument> {
    let mut reader = make_content_reader(xml_data);
    let mut elements = Vec::new();
    let mut sections = Vec::new();
    let mut in_body = false;
    // Element dispatch below matches on local name only, so a `<evil:p>`
    // inside the body would be parsed as a WordprocessingML paragraph and
    // its text extracted as document content Word never renders. The guard
    // records which prefixes the root bound to WML and rejects the rest;
    // a root that binds its own prefix to something else is refused
    // outright, because such a document is not the format it claims.
    let mut guard = xml::NsGuard::permissive();
    let mut saw_root = false;
    // `<w:altChunk>` references a part holding content injected into the
    // document — mail merge, report generators and CMS exporters all use
    // it, often for the entire body with `document.xml` holding only a
    // shell. Such a document extracted as almost nothing.
    let mut alt_chunk_rids: Vec<String> = Vec::new();
    let mut alt_chunk_positions: Vec<usize> = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if !saw_root {
                    saw_root = true;
                    guard = xml::NsGuard::from_root(
                        e,
                        &[xml::ns::WML, xml::ns::STRICT_WML],
                        "WordprocessingML",
                    )?;
                }
                if !guard.accepts(e) {
                    xml::skip_element_fast(&mut reader)?;
                    continue;
                }
                match e.local_name().as_ref() {
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
                    b"altChunk" if in_body => {
                        if let Ok(Some(rid)) = xml::optional_attr_str(e, b"r:id") {
                            alt_chunk_rids.push(rid.into_owned());
                            // Record the insertion point so the chunk's
                            // content lands where the document puts it.
                            alt_chunk_positions.push(elements.len());
                        }
                    },
                    _ => {},
                }
            },
            Event::Empty(ref e) if in_body && e.local_name().as_ref() == b"altChunk" => {
                if guard.accepts(e) {
                    if let Ok(Some(rid)) = xml::optional_attr_str(e, b"r:id") {
                        alt_chunk_rids.push(rid.into_owned());
                        alt_chunk_positions.push(elements.len());
                    }
                }
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
    Ok((body, all_sections, alt_chunk_rids, alt_chunk_positions))
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
                    match content {
                        ParagraphContent::Hyperlink(hl) => {
                            if let HyperlinkTarget::External(ref r_id) = hl.target {
                                match rels.get_by_id(r_id) {
                                    Some(rel) => {
                                        // A `w:anchor` alongside `r:id` is a URI
                                        // fragment of the resolved target.
                                        let mut t = rel.target.clone();
                                        if let Some(frag) =
                                            hl.fragment.as_deref().filter(|f| !f.is_empty())
                                        {
                                            t.push('#');
                                            t.push_str(frag);
                                        }
                                        hl.target = if rel.target_mode == TargetMode::External {
                                            HyperlinkTarget::External(t)
                                        } else {
                                            HyperlinkTarget::Internal(t)
                                        };
                                    },
                                    // Unresolvable relationship: a bare
                                    // relationship id is not a URL. Fall back
                                    // to the anchor when the element carried
                                    // one rather than reporting `rIdN`.
                                    None => {
                                        if let Some(frag) =
                                            hl.fragment.as_deref().filter(|f| !f.is_empty())
                                        {
                                            hl.target =
                                                HyperlinkTarget::Internal(frag.to_string());
                                        }
                                    },
                                }
                            }
                            for run in &mut hl.runs {
                                resolve_hyperlinks_in_run(run, rels);
                            }
                        },
                        ParagraphContent::Run(run) => {
                            resolve_hyperlinks_in_run(run, rels);
                        },
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

/// Resolve hyperlinks nested inside a run's text-box bodies. Text-box
/// prose is ordinary content and carries ordinary links.
fn resolve_hyperlinks_in_run(run: &mut Run, rels: &crate::core::relationships::Relationships) {
    for rc in &mut run.content {
        if let RunContent::TextBox(blocks) = rc {
            resolve_hyperlinks(blocks, rels);
        }
    }
}

// ---------------------------------------------------------------------------
// Paragraph parsing
// ---------------------------------------------------------------------------

/// Elements that wrap runs without contributing content of their own.
///
/// Each of these is a *transparent* container in WordprocessingML: its
/// children are ordinary paragraph content. Skipping the whole subtree —
/// which is what the catch-all arm did — silently deleted the text inside.
/// `w:ins` (a tracked insertion) is the most common by far: any document
/// edited with track-changes on and not yet accepted keeps its text there.
///
/// `w:del` is deliberately absent: its `w:delText` children are *deleted*
/// text and are not part of the document.
fn is_transparent_paragraph_wrapper(local: &[u8]) -> bool {
    matches!(
        local,
        b"ins"
            | b"moveTo"
            | b"fldSimple"
            | b"smartTag"
            | b"sdt"
            | b"sdtContent"
            | b"ruby"
            | b"rt"
            | b"rubyBase"
            | b"bdo"
            | b"dir"
            | b"customXml"
    )
}

fn parse_paragraph(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<Paragraph> {
    let mut paragraph = Paragraph::default();
    // Depth of transparent wrappers we have descended into, so their
    // closing tags are consumed without ending the paragraph.
    let mut wrapper_depth = 0usize;
    // Open complex fields (`{ HYPERLINK … }`). A field spans several runs,
    // so it is stitched together here rather than inside `parse_run`.
    let mut fields: Vec<OpenField> = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"pPr" => {
                    paragraph.properties = Some(parse_paragraph_properties_fast(reader)?);
                },
                b"r" => {
                    let mut parts = Vec::new();
                    let run = parse_run(reader, &mut parts)?;
                    apply_field_parts(&parts, &mut paragraph.content, &mut fields);
                    if !run.content.is_empty() || run.properties.is_some() {
                        paragraph.content.push(ParagraphContent::Run(run));
                    }
                },
                b"hyperlink" => {
                    paragraph
                        .content
                        .push(ParagraphContent::Hyperlink(parse_hyperlink(reader, e)?));
                },
                // `w:fldSimple` is the one-element form of the same field
                // mechanism; its `w:instr` attribute holds the URL of a
                // HYPERLINK field (issue #267). Other field types keep the
                // existing transparent-wrapper behaviour.
                b"fldSimple" => {
                    let target = xml::optional_attr_str(e, b"w:instr")?
                        .as_deref()
                        .and_then(hyperlink_target_from_instr);
                    match target {
                        Some(target) => {
                            let runs = collect_runs_until(reader, b"fldSimple")?;
                            paragraph.content.push(ParagraphContent::Hyperlink(Hyperlink {
                                target,
                                fragment: None,
                                tooltip: None,
                                runs,
                            }));
                        },
                        None => wrapper_depth += 1,
                    }
                },
                // OMML equations: no structural model, but every `<m:t>`
                // inside one is real, visible text (issue #270).
                b"oMath" | b"oMathPara" => {
                    let end: &[u8] = if e.local_name().as_ref() == b"oMathPara" {
                        b"oMathPara"
                    } else {
                        b"oMath"
                    };
                    let text = collect_omml_text(reader, end)?;
                    if !text.is_empty() {
                        paragraph.content.push(ParagraphContent::Run(Run {
                            properties: None,
                            content: vec![RunContent::Text(text)],
                        }));
                    }
                },
                b"del" | b"moveFrom" => {
                    // Tracked deletions are not document content.
                    xml::skip_element_fast(reader)?;
                },
                local if is_transparent_paragraph_wrapper(local) => {
                    wrapper_depth += 1;
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) => {
                let local = e.local_name();
                if local.as_ref() == b"p" && wrapper_depth == 0 {
                    break;
                }
                if is_transparent_paragraph_wrapper(local.as_ref()) {
                    wrapper_depth = wrapper_depth.saturating_sub(1);
                } else if local.as_ref() == b"p" {
                    // A malformed file closed the paragraph while a wrapper
                    // was still open; stop rather than swallow the rest.
                    break;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(paragraph)
}

/// One part of a complex field (`{ HYPERLINK "…" }`), reported out of
/// `parse_run` so the enclosing paragraph can stitch the sequence back
/// together: the instruction and the display runs live in *different*
/// runs, so neither alone can resolve the field.
#[derive(Debug, Clone)]
enum FieldPart {
    /// `<w:fldChar w:fldCharType="begin"/>`.
    Begin,
    /// `<w:fldChar w:fldCharType="separate"/>` — the field result starts.
    Separate,
    /// `<w:fldChar w:fldCharType="end"/>`.
    End,
    /// `<w:instrText>` content (a field's instruction is often split
    /// across several runs, so these accumulate).
    Instr(String),
}

/// An in-progress complex field, tracked across the several runs that make
/// up `begin … instrText … separate … <display runs> … end`.
struct OpenField {
    /// Accumulated `<w:instrText>` content across every run seen so far.
    instr: String,
    /// Index into the paragraph's `content` vec where the field's display
    /// runs start — set when `separate` is seen, since everything from
    /// that point until `end` is the field's rendered result.
    content_start: usize,
    /// Whether `separate` has been seen yet. A field with no `separate`
    /// (some writers omit it, e.g. a display-only field) has no distinct
    /// result span to splice out.
    separated: bool,
}

/// Fold one run's field-code events into the paragraph's open-field stack,
/// resolving a completed field into a `ParagraphContent::Hyperlink` when it
/// turns out to be a `HYPERLINK` field (issue #267).
///
/// Fields nest in principle (a field's instruction can itself contain
/// another field), so `fields` is a stack — but only the outermost
/// completed `HYPERLINK` field is ever turned into a hyperlink; an inner
/// field's own `instr`/`content_start` are simply discarded when it ends,
/// since nested field results already sit in `content` as ordinary runs.
fn apply_field_parts(
    parts: &[FieldPart],
    content: &mut Vec<ParagraphContent>,
    fields: &mut Vec<OpenField>,
) {
    for part in parts {
        match part {
            FieldPart::Begin => {
                fields.push(OpenField {
                    instr: String::new(),
                    content_start: content.len(),
                    separated: false,
                });
            },
            FieldPart::Instr(s) => {
                if let Some(f) = fields.last_mut() {
                    f.instr.push_str(s);
                }
            },
            FieldPart::Separate => {
                if let Some(f) = fields.last_mut() {
                    f.separated = true;
                    f.content_start = content.len();
                }
            },
            FieldPart::End => {
                let Some(f) = fields.pop() else { continue };
                if !f.separated {
                    continue;
                }
                let Some(target) = hyperlink_target_from_instr(&f.instr) else {
                    continue;
                };
                let start = f.content_start.min(content.len());
                let runs: Vec<Run> = content
                    .drain(start..)
                    .filter_map(|c| match c {
                        ParagraphContent::Run(r) => Some(r),
                        ParagraphContent::Hyperlink(h) => Some(Run {
                            properties: None,
                            content: h.runs.into_iter().flat_map(|r| r.content).collect(),
                        }),
                    })
                    .collect();
                content.push(ParagraphContent::Hyperlink(Hyperlink {
                    target,
                    fragment: None,
                    tooltip: None,
                    runs,
                }));
            },
        }
    }
}

/// Collect every `<m:t>` text run inside an OMML equation (`<m:oMath>` or
/// `<m:oMathPara>`), concatenated with no separators. This is not a
/// structural math model — just enough to stop 100% content loss on a
/// document whose only content is a formula (issue #270).
fn collect_omml_text(reader: &mut quick_xml::Reader<&[u8]>, end_local: &[u8]) -> CoreResult<String> {
    let mut text = String::new();
    let mut depth = 1i32;
    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.local_name().as_ref() == b"t" {
                    text.push_str(&xml::read_text_content_fast(reader)?);
                } else {
                    depth += 1;
                }
            },
            Event::End(ref e) => {
                if e.local_name().as_ref() == end_local && depth <= 1 {
                    break;
                }
                depth -= 1;
                if depth <= 0 {
                    break;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(text)
}

fn parse_run(
    reader: &mut quick_xml::Reader<&[u8]>,
    fields: &mut Vec<FieldPart>,
) -> CoreResult<Run> {
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
                    // A `<w:drawing>` may wrap a picture *or* a shape whose
                    // `<wps:txbx>` holds real prose. Collect both.
                    let (drawing, boxes) = parse_drawing_and_text_boxes(reader)?;
                    if let Some(drawing) = drawing {
                        run.content.push(RunContent::Drawing(drawing));
                    }
                    for b in boxes {
                        run.content.push(RunContent::TextBox(b));
                    }
                },
                // VML shapes (`<w:pict>`) and the compatibility wrapper
                // (`<mc:AlternateContent>`) are the other two places a text
                // box hides. `parse_vml_content_in` also resolves
                // AlternateContent to its `<mc:Choice>` branch so shape text
                // is not extracted twice (once per branch), and picks up the
                // *other* payloads a VML shape can carry: legacy images
                // (issue #268), WordArt (#274) and embedded packages (#304).
                b"pict" | b"object" => {
                    let end: &[u8] = if e.local_name().as_ref() == b"object" {
                        b"object"
                    } else {
                        b"pict"
                    };
                    parse_vml_content_in(reader, end)?.push_into(&mut run.content);
                },
                b"AlternateContent" => {
                    parse_alternate_content(reader)?.push_into(&mut run.content);
                },
                // A note's reference mark. The mark *is* content: it is
                // where the note is cited (issue #241).
                b"footnoteReference" | b"endnoteReference" | b"commentReference" => {
                    push_note_reference(e, &mut run.content)?;
                    xml::skip_element_fast(reader)?;
                },
                // Complex field codes. The `begin` char also carries
                // `<w:ffData>` for legacy form fields, whose state exists
                // nowhere else in the document (issue #276).
                b"fldChar" => {
                    fields.push(fld_char_part(e)?);
                    if let Some(mut ff) = parse_fld_char_body(reader)? {
                        ff.display_text = ff.value_text();
                        run.content.push(RunContent::FormField(ff));
                    }
                },
                // The field instruction — for a HYPERLINK field this holds
                // the URL, which was previously unreachable (issue #267).
                b"instrText" => {
                    fields.push(FieldPart::Instr(xml::read_text_content_fast(reader)?));
                },
                // `<w:sym>` carries its character in the `w:char` attribute
                // as a hex code point, usually in the Wingdings private-use
                // range. Dropping it silently deleted bullet glyphs and
                // maths symbols from the text.
                b"sym" => {
                    if let Some(c) = parse_sym_char(e) {
                        run.content.push(RunContent::Text(c.to_string()));
                    }
                    xml::skip_element_fast(reader)?;
                },
                b"cr" => {
                    run.content.push(RunContent::Break(BreakType::Line));
                    xml::skip_element_fast(reader)?;
                },
                b"noBreakHyphen" => {
                    run.content.push(RunContent::Text("\u{2011}".to_string()));
                    xml::skip_element_fast(reader)?;
                },
                // `<w:softHyphen/>` is a *discretionary* line-break hint, not
                // content: Word draws it only when the line happens to break
                // there. Emitting U+00AD splits the word for every consumer
                // doing word-level work — search, RAG, the uses this library
                // exists for — turning `Fähigkeit` into `Fähig` + `keit`.
                // One real corpus file carries 68 of them. Dropped, which is
                // what every mainstream extractor does. `w:noBreakHyphen`
                // above is the opposite case: a hyphen the document actually
                // draws, so it is kept.
                b"softHyphen" => {
                    xml::skip_element_fast(reader)?;
                },
                b"delText" => {
                    // Deleted revision text is not document content.
                    xml::skip_element_fast(reader)?;
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
                b"cr" => {
                    run.content.push(RunContent::Break(BreakType::Line));
                },
                b"noBreakHyphen" => {
                    run.content.push(RunContent::Text("\u{2011}".to_string()));
                },
                // See the Start arm: a discretionary hyphen is not content.
                b"softHyphen" => {},
                b"sym" => {
                    if let Some(c) = parse_sym_char(e) {
                        run.content.push(RunContent::Text(c.to_string()));
                    }
                },
                b"footnoteReference" | b"endnoteReference" | b"commentReference" => {
                    push_note_reference(e, &mut run.content)?;
                },
                b"fldChar" => {
                    fields.push(fld_char_part(e)?);
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

/// Push the `RunContent` for a `w:footnoteReference` /
/// `w:endnoteReference` / `w:commentReference` mark.
fn push_note_reference(
    e: &quick_xml::events::BytesStart,
    out: &mut Vec<RunContent>,
) -> CoreResult<()> {
    let id: u32 = xml::optional_attr_str(e, b"w:id")?
        .and_then(|v| v.trim().parse::<i64>().ok())
        .map(|v| v.max(0) as u32)
        .unwrap_or(0);
    out.push(match e.local_name().as_ref() {
        b"footnoteReference" => RunContent::FootnoteRef(id),
        b"endnoteReference" => RunContent::EndnoteRef(id),
        _ => RunContent::CommentRef(id),
    });
    Ok(())
}

/// Map a `<w:fldChar w:fldCharType="…">` onto its field part. An absent or
/// unrecognised type is treated as `begin`, which is what Word writes when
/// the attribute is omitted.
fn fld_char_part(e: &quick_xml::events::BytesStart) -> CoreResult<FieldPart> {
    Ok(match xml::optional_attr_str(e, b"w:fldCharType")?.as_deref() {
        Some("separate") => FieldPart::Separate,
        Some("end") => FieldPart::End,
        _ => FieldPart::Begin,
    })
}

/// Read the children of a non-empty `<w:fldChar>`, returning the
/// `<w:ffData>` form-field state when it carries one.
fn parse_fld_char_body(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<Option<FormField>> {
    let mut form = None;
    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.local_name().as_ref() == b"ffData" {
                    form = Some(parse_ff_data(reader)?);
                } else {
                    xml::skip_element_fast(reader)?;
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"fldChar" => break,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(form)
}

/// Parse `<w:ffData>`: the checkbox state, the dropdown option list and
/// selection, or the text field's default value. Reads through
/// `</w:ffData>`.
fn parse_ff_data(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<FormField> {
    // Which `w:ffData` child we are inside: `w:default` means a different
    // thing in each (a checkbox's initial state, a dropdown's index, a text
    // field's value).
    #[derive(PartialEq)]
    enum Kind {
        None,
        CheckBox,
        DdList,
        TextInput,
    }

    let mut name = None;
    let mut kind = Kind::None;
    let mut checked: Option<bool> = None;
    let mut cb_default = false;
    let mut entries: Vec<String> = Vec::new();
    let mut selected = 0usize;
    let mut text_default: Option<String> = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let val = xml::optional_attr_str(e, b"w:val")?.map(|v| v.into_owned());
                match e.local_name().as_ref() {
                    b"name" => name = val,
                    b"checkBox" => kind = Kind::CheckBox,
                    b"ddList" => kind = Kind::DdList,
                    b"textInput" => kind = Kind::TextInput,
                    b"checked" => checked = Some(xml::parse_toggle(e, b"w:val")),
                    b"listEntry" => entries.push(val.unwrap_or_default()),
                    b"result" => {
                        if let Some(v) = val.as_deref().and_then(|v| v.trim().parse::<usize>().ok())
                        {
                            selected = v;
                        }
                    },
                    b"default" => match kind {
                        Kind::CheckBox => cb_default = xml::parse_toggle(e, b"w:val"),
                        Kind::DdList => {
                            if let Some(v) =
                                val.as_deref().and_then(|v| v.trim().parse::<usize>().ok())
                            {
                                selected = v;
                            }
                        },
                        Kind::TextInput => text_default = val,
                        Kind::None => {},
                    },
                    _ => {},
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"ffData" => break,
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(FormField {
        name,
        kind: match kind {
            Kind::CheckBox => FormFieldKind::CheckBox {
                checked: checked.unwrap_or(cb_default),
            },
            Kind::DdList => FormFieldKind::DropDown { entries, selected },
            Kind::TextInput => FormFieldKind::TextInput {
                default: text_default,
            },
            Kind::None => FormFieldKind::Unknown,
        },
        display_text: None,
    })
}

/// Split a field instruction into quote-aware tokens:
/// `HYPERLINK "http://x" \l "frag"` → `[HYPERLINK, http://x, \l, frag]`.
fn field_instr_tokens(instr: &str) -> Vec<String> {
    let mut out = Vec::new();
    let mut cur = String::new();
    let mut quoted = false;
    for c in instr.chars() {
        match c {
            '"' => {
                quoted = !quoted;
                if !quoted {
                    out.push(std::mem::take(&mut cur));
                }
            },
            c if c.is_whitespace() && !quoted => {
                if !cur.is_empty() {
                    out.push(std::mem::take(&mut cur));
                }
            },
            c => cur.push(c),
        }
    }
    if !cur.is_empty() {
        out.push(cur);
    }
    out
}

/// Resolve a `HYPERLINK` field instruction into a hyperlink target.
///
/// `HYPERLINK "https://x" \l "frag"` → external `https://x#frag`;
/// `HYPERLINK \l "bookmark"` → internal `bookmark`. Returns `None` for
/// every other field type (PAGE, TOC, REF, …), whose display text already
/// survives as an ordinary run.
fn hyperlink_target_from_instr(instr: &str) -> Option<HyperlinkTarget> {
    let tokens = field_instr_tokens(instr);
    let (first, rest) = tokens.split_first()?;
    if !first.eq_ignore_ascii_case("HYPERLINK") {
        return None;
    }
    let mut url: Option<String> = None;
    let mut anchor: Option<String> = None;
    let mut i = 0;
    while i < rest.len() {
        let tok = &rest[i];
        if let Some(switch) = tok.strip_prefix('\\') {
            // `\l` takes the sub-address; `\o`/`\t` take an argument we
            // do not model; `\n`/`\h`/`\m` take none.
            let takes_arg = matches!(switch, "l" | "o" | "t" | "L" | "O" | "T");
            if takes_arg {
                if let Some(arg) = rest.get(i + 1) {
                    if switch.eq_ignore_ascii_case("l") {
                        anchor = Some(arg.clone());
                    }
                    i += 1;
                }
            }
        } else if url.is_none() {
            url = Some(tok.clone());
        }
        i += 1;
    }
    match (url, anchor) {
        (Some(mut u), anchor) => {
            if let Some(a) = anchor.filter(|a| !a.is_empty()) {
                u.push('#');
                u.push_str(&a);
            }
            Some(HyperlinkTarget::External(u))
        },
        (None, Some(a)) => Some(HyperlinkTarget::Internal(a)),
        (None, None) => None,
    }
}

fn parse_hyperlink(
    reader: &mut quick_xml::Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> CoreResult<Hyperlink> {
    // Determine target: r:id for external, w:anchor for internal
    let r_id = xml::optional_attr_str(start, b"r:id")?.map(|v| v.into_owned());
    let anchor = xml::optional_attr_str(start, b"w:anchor")?.map(|v| v.into_owned());
    let tooltip = xml::optional_attr_str(start, b"w:tooltip")?.map(|v| v.into_owned());

    // `r:id` wins when both attributes are present: ECMA-376 makes
    // `w:anchor` a *fragment* of the relationship's target in that case
    // (`externalURL#fragment`). Taking the anchor branch unconditionally
    // discarded the real URL and left a dead same-document reference
    // behind — the single most prevalent hyperlink defect in the corpus
    // (issues #242, #292). The anchor is kept in `fragment` and appended
    // by `resolve_hyperlinks` once the relationship is known.
    let (target, fragment) = match (r_id, anchor) {
        // Will be resolved to the actual URL after parsing via
        // `resolve_hyperlinks()`.
        (Some(r_id), anchor) => (HyperlinkTarget::External(r_id), anchor),
        (None, Some(anchor)) => (HyperlinkTarget::Internal(anchor), None),
        (None, None) => (HyperlinkTarget::Internal(String::new()), None),
    };

    let runs = collect_runs_until(reader, b"hyperlink")?;

    Ok(Hyperlink {
        target,
        fragment,
        tooltip,
        runs,
    })
}

/// Collect the `<w:r>` children of an inline container, reading through
/// the matching `</end_local>`. Shared by `w:hyperlink` and the
/// `w:fldSimple` HYPERLINK path.
fn collect_runs_until(
    reader: &mut quick_xml::Reader<&[u8]>,
    end_local: &[u8],
) -> CoreResult<Vec<Run>> {
    let mut runs = Vec::new();
    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.local_name().as_ref() == b"r" {
                    // A field cannot legally nest inside w:fldSimple's own
                    // display runs, so field-part tracking is a fresh,
                    // throwaway vec here.
                    runs.push(parse_run(reader, &mut Vec::new())?);
                } else {
                    xml::skip_element_fast(reader)?;
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == end_local => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(runs)
}

// ---------------------------------------------------------------------------
// Embedded chart text
// ---------------------------------------------------------------------------

/// Walk a parsed body and copy each chart part's extracted text onto the
/// drawing that references it, keyed by relationship id.
///
/// Charts nest wherever drawings do — inside table cells and inside text
/// boxes — so the walk recurses through both rather than only scanning
/// top-level paragraphs.
fn attach_chart_text(
    elements: &mut [BlockElement],
    chart_text_by_rid: &std::collections::HashMap<String, Vec<String>>,
    dgm_text_by_rid: &std::collections::HashMap<String, Vec<String>>,
) {
    for elem in elements {
        match elem {
            BlockElement::Paragraph(p) => {
                for pc in &mut p.content {
                    let runs: &mut [Run] = match pc {
                        ParagraphContent::Run(r) => std::slice::from_mut(r),
                        ParagraphContent::Hyperlink(hl) => &mut hl.runs,
                    };
                    for run in runs {
                        for rc in &mut run.content {
                            match rc {
                                RunContent::Drawing(d) => {
                                    if let Some(rid) = d.chart_rel_id.as_deref() {
                                        if let Some(lines) = chart_text_by_rid.get(rid) {
                                            d.chart_text = lines.clone();
                                        }
                                    }
                                    if let Some(rid) = d.dgm_data_rel_id.as_deref() {
                                        if let Some(lines) = dgm_text_by_rid.get(rid) {
                                            d.dgm_text = lines.clone();
                                        }
                                    }
                                },
                                RunContent::TextBox(blocks) => {
                                    attach_chart_text(blocks, chart_text_by_rid, dgm_text_by_rid);
                                },
                                _ => {},
                            }
                        }
                    }
                }
            },
            BlockElement::Table(t) => {
                for row in &mut t.rows {
                    for cell in &mut row.cells {
                        attach_chart_text(&mut cell.content, chart_text_by_rid, dgm_text_by_rid);
                    }
                }
            },
        }
    }
}

/// Replace every `RunContent::DeferredPart(rid)` in the tree with the
/// extracted plain text of the OOXML package it refers to
/// (`<o:OLEObject Type="Embed">`, issue #304), wrapped as a `TextBox` so
/// it flows as ordinary block content. A reference that can't be resolved
/// (unknown relationship, unreadable part, unrecognised extension, or the
/// nested document fails to open) is left as `DeferredPart` and every
/// renderer's catch-all already drops those silently — matching the
/// pre-existing "unresolvable is dropped" behaviour.
fn resolve_deferred_parts<R: Read + Seek>(
    elements: &mut [BlockElement],
    opc: &mut OpcReader<R>,
    main_part: &crate::core::opc::PartName,
    doc_rels: &crate::core::relationships::Relationships,
) {
    for elem in elements {
        match elem {
            BlockElement::Paragraph(p) => {
                for pc in &mut p.content {
                    let runs: &mut [Run] = match pc {
                        ParagraphContent::Run(r) => std::slice::from_mut(r),
                        ParagraphContent::Hyperlink(hl) => &mut hl.runs,
                    };
                    for run in runs {
                        for rc in &mut run.content {
                            match rc {
                                RunContent::TextBox(blocks) => {
                                    resolve_deferred_parts(blocks, opc, main_part, doc_rels);
                                },
                                RunContent::DeferredPart(rid) => {
                                    if let Some(blocks) =
                                        resolve_one_deferred_part(rid, opc, main_part, doc_rels)
                                    {
                                        *rc = RunContent::TextBox(blocks);
                                    }
                                },
                                _ => {},
                            }
                        }
                    }
                }
            },
            BlockElement::Table(t) => {
                for row in &mut t.rows {
                    for cell in &mut row.cells {
                        resolve_deferred_parts(&mut cell.content, opc, main_part, doc_rels);
                    }
                }
            },
        }
    }
}

fn resolve_one_deferred_part<R: Read + Seek>(
    rid: &str,
    opc: &mut OpcReader<R>,
    main_part: &crate::core::opc::PartName,
    doc_rels: &crate::core::relationships::Relationships,
) -> Option<Vec<BlockElement>> {
    let rel = doc_rels.get_by_id(rid)?;
    if rel.target_mode != TargetMode::Internal {
        return None;
    }
    let part = main_part.resolve_relative(&rel.target).ok()?;
    if !opc.has_part(&part) {
        return None;
    }
    let ext = part.as_str().rsplit('.').next()?;
    let format = crate::format::DocumentFormat::from_extension(ext)?;
    let data = opc.read_part(&part).ok()?;
    let text = crate::Document::from_reader(std::io::Cursor::new(data), format)
        .ok()?
        .plain_text();
    let paragraphs: Vec<BlockElement> = text
        .lines()
        .filter(|l| !l.trim().is_empty())
        .map(|line| {
            BlockElement::Paragraph(Paragraph {
                properties: None,
                content: vec![ParagraphContent::Run(Run {
                    properties: None,
                    content: vec![RunContent::Text(line.to_string())],
                })],
            })
        })
        .collect();
    if paragraphs.is_empty() {
        None
    } else {
        Some(paragraphs)
    }
}

/// Collect every `<a:t>` text node inside a SmartArt data part
/// (`word/diagrams/dataN.xml`), in document order, one per line
/// (issue #271).
fn diagram_text_lines(xml: &[u8]) -> Vec<String> {
    let mut reader = xml::make_fast_reader(xml);
    let mut lines = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"t" => {
                if let Ok(text) = xml::read_text_content_fast(&mut reader) {
                    let text = text.trim();
                    if !text.is_empty() {
                        lines.push(text.to_string());
                    }
                }
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {},
        }
    }
    lines
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
/// Parse `<w:drawing>`, returning both the picture/shape descriptor and the
/// bodies of any text boxes (`<wps:txbx><w:txbxContent>`) it contains. A
/// DrawingML shape can carry real prose, so a parser that only looked for a
/// blip threw that text away.
fn parse_drawing_and_text_boxes(
    reader: &mut quick_xml::Reader<&[u8]>,
) -> CoreResult<(Option<DrawingInfo>, Vec<Vec<BlockElement>>)> {
    let mut info: Option<DrawingInfo> = None;
    let mut boxes: Vec<Vec<BlockElement>> = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"inline" => {
                    info = parse_inline_or_anchor_body(
                        reader, /*inline=*/ true, b"inline", &mut boxes,
                    )?;
                },
                b"anchor" => {
                    info = parse_inline_or_anchor_body(
                        reader, /*inline=*/ false, b"anchor", &mut boxes,
                    )?;
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

    Ok((info, boxes))
}

/// Parse the body of `<wp:inline>` or `<wp:anchor>` until the matching
/// closing tag (`end_local`). Collects extent, docPr, position, and the
/// graphic payload (image or shape) into a `DrawingInfo`.
fn parse_inline_or_anchor_body(
    reader: &mut quick_xml::Reader<&[u8]>,
    inline: bool,
    end_local: &[u8],
    text_boxes: &mut Vec<Vec<BlockElement>>,
) -> CoreResult<Option<DrawingInfo>> {
    use crate::docx::image::{AnchorFrame, AnchorPosition};

    let mut width = Emu(0);
    let mut height = Emu(0);
    let mut description: Option<String> = None;
    let mut relationship_id: Option<String> = None;
    let mut shape: Option<crate::docx::image::ShapeInfo> = None;
    let mut chart_rel_id: Option<String> = None;
    let mut dgm_data_rel_id: Option<String> = None;

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
                    let g = parse_graphic(reader, text_boxes)?;
                    if let Some(rid) = g.relationship_id {
                        relationship_id = Some(rid);
                    }
                    if let Some(s) = g.shape {
                        shape = Some(s);
                    }
                    if let Some(rid) = g.chart_rel_id {
                        chart_rel_id = Some(rid);
                    }
                    if let Some(rid) = g.dgm_data_rel_id {
                        dgm_data_rel_id = Some(rid);
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

    // A chart-only or diagram-only graphic has neither a blip nor a
    // `prstGeom`, so without these two conditions the whole drawing was
    // discarded and the chart/diagram part was never reachable (#273, #271).
    if relationship_id.is_some()
        || shape.is_some()
        || chart_rel_id.is_some()
        || dgm_data_rel_id.is_some()
    {
        Ok(Some(DrawingInfo {
            relationship_id: relationship_id.unwrap_or_default(),
            description,
            width,
            height,
            inline,
            anchor_position,
            shape,
            chart_rel_id,
            chart_text: Vec::new(),
            dgm_data_rel_id,
            dgm_text: Vec::new(),
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
    /// `r:id` of a `<c:chart>` reference, when the graphic is a chart.
    chart_rel_id: Option<String>,
    /// `r:dm` of a `<dgm:relIds>` reference, when the graphic is a
    /// SmartArt diagram (issue #271).
    dgm_data_rel_id: Option<String>,
}

/// Parse `<a:graphic>` and any contained `<pic:pic>` (image) or
/// `<wps:wsp>` (vector shape). Reads through `</a:graphic>`.
fn parse_graphic(
    reader: &mut quick_xml::Reader<&[u8]>,
    text_boxes: &mut Vec<Vec<BlockElement>>,
) -> CoreResult<GraphicPayload> {
    let mut relationship_id: Option<String> = None;
    let mut shape: Option<crate::docx::image::ShapeInfo> = None;
    let mut chart_rel_id: Option<String> = None;
    let mut dgm_data_rel_id: Option<String> = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"pic" => {
                    if let Some(rid) = parse_pic(reader)? {
                        relationship_id = Some(rid);
                    }
                },
                b"wsp" => {
                    if let Some(s) = parse_wsp(reader, text_boxes)? {
                        shape = Some(s);
                    }
                },
                // A native chart: the graphic carries only a relationship
                // pointing at `word/charts/chartN.xml`, where all of its
                // text actually lives. Usually the empty form, but the
                // element is allowed children (`<c:extLst>`).
                b"chart" => {
                    if let Some(rid) = xml::optional_prefixed_attr_str(e, b"id")? {
                        chart_rel_id = Some(rid.into_owned());
                    }
                    xml::skip_element_fast(reader)?;
                },
                // A SmartArt diagram: `<dgm:relIds r:dm="…" r:lo="…" .../>`
                // — `r:dm` points at the data part (`word/diagrams/dataN.xml`)
                // where the diagram's actual text lives (issue #271).
                b"relIds" => {
                    if let Some(rid) = xml::optional_prefixed_attr_str(e, b"dm")? {
                        dgm_data_rel_id = Some(rid.into_owned());
                    }
                    xml::skip_element_fast(reader)?;
                },
                // A group shape nests further `<wps:wsp>` children; descend
                // so text boxes inside groups are not lost.
                b"grpSp" | b"wgp" => continue,
                // <a:graphicData> is just a wrapper; descend into it.
                b"graphicData" => continue,
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::Empty(ref e) if e.local_name().as_ref() == b"chart" => {
                if let Some(rid) = xml::optional_prefixed_attr_str(e, b"id")? {
                    chart_rel_id = Some(rid.into_owned());
                }
            },
            Event::Empty(ref e) if e.local_name().as_ref() == b"relIds" => {
                if let Some(rid) = xml::optional_prefixed_attr_str(e, b"dm")? {
                    dgm_data_rel_id = Some(rid.into_owned());
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"graphic" => break,
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(GraphicPayload {
        relationship_id,
        shape,
        chart_rel_id,
        dgm_data_rel_id,
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
    text_boxes: &mut Vec<Vec<BlockElement>>,
) -> CoreResult<Option<crate::docx::image::ShapeInfo>> {
    use crate::docx::image::{ShapeInfo, ShapeKind};

    let mut kind: Option<ShapeKind> = None;
    let mut stroke_rgb: Option<(u8, u8, u8)> = None;
    let mut fill_rgb: Option<(u8, u8, u8)> = None;
    let mut stroke_w_emu: Option<i64> = None;
    // A `<wps:wsp>` with both `<a:prstGeom>` (shape geometry, schema-
    // required) AND `<wps:txbx><w:txbxContent>` (real text) is the
    // standard way Word encodes a text box — not a shape that happens to
    // also carry a text box. Emitting both `Element::Shape` and
    // `Element::TextBox` for the one drawing doubled it into a
    // content-free phantom shape plus the real text box (issue #263).
    let mut has_text_box = false;

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
                // `<wps:txbx>` is a wrapper; `<w:txbxContent>` inside it is
                // ordinary block content.
                b"txbx" => continue,
                b"txbxContent" => {
                    text_boxes.push(parse_block_elements_until(reader, b"txbxContent")?);
                    has_text_box = true;
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

    if has_text_box {
        return Ok(None);
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
    // Tables nest (a cell may hold another table), so this is the recursion
    // an adversarial file drives. Past the depth limit the subtree is
    // skipped rather than recursed into: a stack overflow aborts the whole
    // process and no caller in any binding can catch it.
    let Some(_depth) = xml::DepthGuard::enter() else {
        // Past the cap the subtree is skipped rather than recursed into.
        // Say so in the content: silent truncation is the defect class this
        // release exists to remove, and a reader cannot otherwise tell a
        // truncated document from a shallow one. Mirrors the visible notice
        // the XLSX row cap emits.
        xml::skip_element_fast(reader)?;
        return Ok(Table {
            properties: None,
            grid: Vec::new(),
            rows: vec![TableRow {
                properties: None,
                cells: vec![TableCell {
                    properties: None,
                    content: vec![BlockElement::Paragraph(Paragraph {
                        properties: None,
                        content: vec![ParagraphContent::Run(Run {
                            properties: None,
                            content: vec![RunContent::Text(format!(
                                "[nested tables deeper than {} levels not shown \
                                 — document truncated]",
                                xml::MAX_NESTING_DEPTH
                            ))],
                        })],
                    })],
                }],
            }],
        });
    };
    let mut properties = None;
    let mut grid = Vec::new();
    let mut rows = Vec::new();
    // A whole row can be wrapped the same way a cell can
    // (`<w:tbl><w:sdt><w:sdtContent><w:tr>`) — same transparent-wrapper
    // treatment as `parse_table_row` gives `<w:tc>` (issue #277).
    let mut wrapper_depth = 0usize;

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
                b"sdt" | b"sdtContent" => {
                    wrapper_depth += 1;
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) => {
                let local = e.local_name();
                if local.as_ref() == b"tbl" && wrapper_depth == 0 {
                    break;
                }
                if matches!(local.as_ref(), b"sdt" | b"sdtContent") {
                    wrapper_depth = wrapper_depth.saturating_sub(1);
                } else if local.as_ref() == b"tbl" {
                    break;
                }
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
                b"tblBorders" => {
                    props.borders =
                        Some(self::formatting::parse_table_borders_fast(reader, b"tblBorders")?);
                },
                b"tblCellMar" => {
                    props.cell_margins = Some(parse_cell_margins(reader, b"tblCellMar")?);
                },
                b"tblInd" => {
                    props.indent = parse_measure_w(e);
                    xml::skip_element_fast(reader)?;
                },
                b"tblCaption" => {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                        props.caption = Some(val.into_owned());
                    }
                    xml::skip_element_fast(reader)?;
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::Empty(ref e) => match e.local_name().as_ref() {
                b"tblInd" => {
                    props.indent = parse_measure_w(e);
                },
                b"tblCaption" => {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                        props.caption = Some(val.into_owned());
                    }
                },
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
    // A cell can be wrapped in a content control (`<w:tr><w:sdt><w:sdtContent>
    // <w:tc>`). Treating `sdt`/`sdtContent` as opaque dropped the whole cell,
    // shifting every subsequent cell in the row into the wrong column
    // (issue #277) — descend into them instead, matching how
    // `parse_paragraph` already treats them as transparent wrappers.
    let mut wrapper_depth = 0usize;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"trPr" => {
                    properties = Some(parse_table_row_properties(reader)?);
                },
                b"tc" => {
                    cells.push(parse_table_cell(reader)?);
                },
                b"sdt" | b"sdtContent" => {
                    wrapper_depth += 1;
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) => {
                let local = e.local_name();
                if local.as_ref() == b"tr" && wrapper_depth == 0 {
                    break;
                }
                if matches!(local.as_ref(), b"sdt" | b"sdtContent") {
                    wrapper_depth = wrapper_depth.saturating_sub(1);
                } else if local.as_ref() == b"tr" {
                    break;
                }
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
            Event::Start(ref e) | Event::Empty(ref e) => match e.local_name().as_ref() {
                b"tblHeader" => {
                    props.is_header = xml::parse_toggle(e, b"w:val");
                },
                b"cantSplit" => {
                    props.cant_split = xml::parse_toggle(e, b"w:val");
                },
                b"trHeight" => {
                    props.height = xml::optional_attr_str(e, b"w:val")
                        .ok()
                        .flatten()
                        .and_then(|v| v.parse().ok());
                    props.height_rule =
                        xml::optional_attr_str(e, b"w:hRule")
                            .ok()
                            .flatten()
                            .map(|v| match v.as_ref() {
                                "exact" => RowHeightRule::Exact,
                                "auto" => RowHeightRule::Auto,
                                _ => RowHeightRule::AtLeast,
                            });
                },
                _ => {},
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
                b"tcBorders" => {
                    props.borders =
                        Some(self::formatting::parse_table_borders_fast(reader, b"tcBorders")?);
                },
                b"tcMar" => {
                    props.margins = Some(parse_cell_margins(reader, b"tcMar")?);
                },
                b"vAlign" => {
                    props.v_align = parse_cell_v_align(e);
                    xml::skip_element_fast(reader)?;
                },
                b"textDirection" => {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                        props.text_direction = Some(val.into_owned());
                    }
                    xml::skip_element_fast(reader)?;
                },
                // `<w:cellDel>` marks the whole cell deleted via tracked
                // changes, pending acceptance (issue #266).
                b"cellDel" => {
                    props.deleted = true;
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
                b"cellDel" => {
                    props.deleted = true;
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
                b"vAlign" => {
                    props.v_align = parse_cell_v_align(e);
                },
                b"textDirection" => {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                        props.text_direction = Some(val.into_owned());
                    }
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

/// Read a `w:w` twip measure off an element (`w:tblInd`, `w:top` in
/// `w:tblCellMar`, …).
fn parse_measure_w(e: &quick_xml::events::BytesStart) -> Option<crate::core::units::Twip> {
    xml::optional_attr_str(e, b"w:w")
        .ok()
        .flatten()
        .and_then(|v| v.parse().ok())
        .map(crate::core::units::Twip)
}

fn parse_cell_v_align(e: &quick_xml::events::BytesStart) -> Option<CellVAlign> {
    xml::optional_attr_str(e, b"w:val")
        .ok()
        .flatten()
        .map(|v| match v.as_ref() {
            "center" => CellVAlign::Center,
            "bottom" => CellVAlign::Bottom,
            _ => CellVAlign::Top,
        })
}

/// Parse the children of `w:tblCellMar` / `w:tcMar`. The caller has consumed
/// the start tag; `end` names the closing element to stop at.
fn parse_cell_margins(
    reader: &mut quick_xml::Reader<&[u8]>,
    end: &[u8],
) -> CoreResult<CellMargins> {
    let mut m = CellMargins::default();
    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let v = parse_measure_w(e).map(|t| t.0);
                match e.local_name().as_ref() {
                    b"top" => m.top = v,
                    b"bottom" => m.bottom = v,
                    b"left" | b"start" => m.left = v,
                    b"right" | b"end" => m.right = v,
                    _ => {},
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == end => break,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(m)
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
                    props.column_layout = Some(ColumnDefs {
                        space: xml::optional_attr_str(e, b"w:space")
                            .ok()
                            .flatten()
                            .and_then(|v| v.parse().ok()),
                        // Absent `w:sep` means no separator; present-with-no-val
                        // means true, which is what `parse_toggle` gives us.
                        separator: xml::optional_attr_str(e, b"w:sep")
                            .ok()
                            .flatten()
                            .is_some_and(|v| matches!(v.as_ref(), "1" | "true" | "on")),
                        widths: Vec::new(),
                    });
                },
                // `<w:col>` children of a non-self-closing `<w:cols>` arrive
                // through this same loop; `</w:cols>` is ignored and only
                // `</w:sectPr>` ends it.
                b"col" => {
                    if let Some(layout) = props.column_layout.as_mut() {
                        if let Ok(Some(w)) = xml::optional_attr_str(e, b"w:w") {
                            if let Ok(v) = w.parse::<u32>() {
                                layout.widths.push(v);
                            }
                        }
                    }
                },
                b"type" => {
                    props.break_type =
                        xml::optional_attr_str(e, b"w:val")?.map(|v| match v.as_ref() {
                            "continuous" => SectionBreakKind::Continuous,
                            "evenPage" => SectionBreakKind::EvenPage,
                            "oddPage" => SectionBreakKind::OddPage,
                            "nextColumn" => SectionBreakKind::NextColumn,
                            _ => SectionBreakKind::NextPage,
                        });
                },
                b"titlePg" => {
                    props.title_page = xml::parse_toggle(e, b"w:val");
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

    /// Like `make_minimal_docx`, but also writes a SmartArt diagram data
    /// part and the `document.xml -> diagrams/data1.xml` relationship
    /// (`rId7`, matching what real Word/pandoc output uses) so a
    /// `<dgm:relIds r:dm="rId7">` reference in `document_xml` resolves.
    fn make_docx_with_diagram(document_xml: &[u8], diagram_data_xml: &[u8]) -> Vec<u8> {
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

        let dgm_part = PartName::new("/word/diagrams/data1.xml").unwrap();
        writer
            .add_part(
                &dgm_part,
                "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
                diagram_data_xml,
            )
            .unwrap();
        let rid = writer.add_part_rel(&doc_part, rel_types::DIAGRAM_DATA, "diagrams/data1.xml");
        assert_eq!(rid, "rId1", "test fixture assumes the first relationship id");

        let result = writer.finish().unwrap();
        result.into_inner()
    }

    #[test]
    fn test_embedded_package_object_text_is_extracted() {
        // issue #304 — an embedded native OOXML package
        // (<o:OLEObject Type="Embed">) was never opened at all.
        let mut xlsx_writer = crate::xlsx::write::XlsxWriter::new();
        {
            let mut sheet = xlsx_writer.add_sheet("Sheet1");
            sheet.set_cell(0, 0, crate::xlsx::write::CellData::String("EmbeddedCellText".to_string()));
        }
        let mut embedded_xlsx = Vec::new();
        xlsx_writer
            .write_to(Cursor::new(&mut embedded_xlsx))
            .unwrap();

        let document_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:o="urn:schemas-microsoft-com:office:office"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:pict>
      <o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" r:id="rId1"/>
    </w:pict></w:r></w:p>
  </w:body>
</w:document>"#;

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
        let embed_part = PartName::new("/word/embeddings/Microsoft_Excel_Worksheet1.xlsx").unwrap();
        writer
            .add_part(
                &embed_part,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                &embedded_xlsx,
            )
            .unwrap();
        let rid = writer.add_part_rel(
            &doc_part,
            rel_types::PACKAGE,
            "embeddings/Microsoft_Excel_Worksheet1.xlsx",
        );
        assert_eq!(rid, "rId1", "test fixture assumes the first relationship id");
        let data = writer.finish().unwrap().into_inner();

        let doc = DocxDocument::from_reader(Cursor::new(data)).unwrap();
        let text = doc.plain_text();
        assert!(
            text.contains("EmbeddedCellText"),
            "expected embedded workbook text in {text:?}"
        );
    }

    #[test]
    fn test_numbered_list_resumed_after_interruption_continues_counting() {
        // issue #243 — a numbered list interrupted by a plain paragraph
        // then resumed later with the same numId (no explicit override)
        // restarted at 1 instead of continuing where it left off.
        let numbering_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>"#;
        let document_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>one</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>two</w:t></w:r></w:p>
    <w:p><w:r><w:t>an interrupting paragraph, not a list item</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>three</w:t></w:r></w:p>
  </w:body>
</w:document>"#;

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
        let numbering_part = PartName::new("/word/numbering.xml").unwrap();
        writer
            .add_part(
                &numbering_part,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
                numbering_xml,
            )
            .unwrap();
        writer.add_part_rel(&doc_part, rel_types::NUMBERING, "numbering.xml");
        let data = writer.finish().unwrap().into_inner();

        let doc = DocxDocument::from_reader(Cursor::new(data)).unwrap();
        let ir = crate::convert_docx::docx_to_ir(&doc);
        let lists: Vec<&crate::ir::List> = ir.sections[0]
            .elements
            .iter()
            .filter_map(|e| match e {
                crate::ir::Element::List(l) => Some(l),
                _ => None,
            })
            .collect();
        assert_eq!(lists.len(), 2, "expected two separate list groups (interrupted once)");
        assert_eq!(
            lists[0].start_number, None,
            "first group: no explicit start, defaults to 1"
        );
        assert_eq!(
            lists[1].start_number,
            Some(3),
            "second group (numId=1 resumed, no override) must continue from item 3, not restart at 1"
        );
    }

    #[test]
    fn test_smartart_diagram_text_is_extracted() {
        // issue #271 — word/diagrams/dataN.xml was never opened, so
        // SmartArt text was completely invisible.
        let document_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
      <dgm:relIds r:dm="rId1" r:lo="rId1" r:qs="rId1" r:cs="rId1"/>
    </a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
  </w:body>
</w:document>"#;
        let diagram_data_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <dgm:ptLst>
    <dgm:pt><dgm:t><a:p><a:r><a:t>top</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt><dgm:t><a:p><a:r><a:t>middle</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt><dgm:t><a:p><a:r><a:t>bottom</a:t></a:r></a:p></dgm:t></dgm:pt>
  </dgm:ptLst>
</dgm:dataModel>"#;
        let data = make_docx_with_diagram(document_xml, diagram_data_xml);
        let doc = DocxDocument::from_reader(Cursor::new(data)).unwrap();
        let text = doc.plain_text();
        assert!(text.contains("top"), "expected 'top' in {text:?}");
        assert!(text.contains("middle"), "expected 'middle' in {text:?}");
        assert!(text.contains("bottom"), "expected 'bottom' in {text:?}");
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
        let info = parse_drawing_and_text_boxes(&mut reader)
            .unwrap()
            .0
            .expect("drawing");
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
        let info = parse_drawing_and_text_boxes(&mut reader)
            .unwrap()
            .0
            .expect("drawing");
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

    /// Parse a `<w:p>…</w:p>` fragment directly through `parse_paragraph`,
    /// the way the body-element dispatcher does: consume the `<w:p>` start
    /// event first, then hand the reader (now positioned just inside) to
    /// the function under test.
    fn parse_paragraph_fragment(xml: &[u8]) -> Paragraph {
        let mut reader = quick_xml::Reader::from_reader(xml);
        reader.config_mut().trim_text(false);
        match reader.read_event().unwrap() {
            Event::Start(_) => {},
            other => panic!("expected <w:p> start, got {other:?}"),
        }
        parse_paragraph(&mut reader).unwrap()
    }

    fn run_texts(p: &Paragraph) -> Vec<String> {
        p.content
            .iter()
            .filter_map(|c| match c {
                ParagraphContent::Run(r) => Some(
                    r.content
                        .iter()
                        .filter_map(|rc| match rc {
                            RunContent::Text(t) => Some(t.clone()),
                            _ => None,
                        })
                        .collect::<String>(),
                ),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn test_ffdata_checkbox_state_is_captured() {
        // issue #276 — a FORMCHECKBOX's checked state exists nowhere else
        // in the document; skipping w:ffData lost it unrecoverably.
        let xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:r><w:fldChar w:fldCharType="begin"><w:ffData>
  <w:name w:val="Check1"/>
  <w:checkBox><w:default w:val="0"/><w:checked w:val="1"/></w:checkBox>
</w:ffData></w:fldChar></w:r>
</w:p>"#;
        let p = parse_paragraph_fragment(xml);
        let ff = p.content.iter().find_map(|c| match c {
            ParagraphContent::Run(r) => r.content.iter().find_map(|rc| match rc {
                RunContent::FormField(ff) => Some(ff.clone()),
                _ => None,
            }),
            _ => None,
        });
        let ff = ff.expect("FormField was not captured");
        assert_eq!(ff.name.as_deref(), Some("Check1"));
        assert_eq!(ff.kind, FormFieldKind::CheckBox { checked: true });
        assert_eq!(ff.value_text().as_deref(), Some("\u{2612}"));
    }

    #[test]
    fn test_ffdata_dropdown_full_option_list_is_captured() {
        // issue #276 — the full option list, not just the selected value,
        // must survive even when the field also has a cached display run.
        let xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:r><w:fldChar w:fldCharType="begin"><w:ffData>
  <w:ddList>
    <w:listEntry w:val="Red"/>
    <w:listEntry w:val="Green"/>
    <w:listEntry w:val="Blue"/>
    <w:result w:val="1"/>
  </w:ddList>
</w:ffData></w:fldChar></w:r>
</w:p>"#;
        let p = parse_paragraph_fragment(xml);
        let ff = p.content.iter().find_map(|c| match c {
            ParagraphContent::Run(r) => r.content.iter().find_map(|rc| match rc {
                RunContent::FormField(ff) => Some(ff.clone()),
                _ => None,
            }),
            _ => None,
        });
        let ff = ff.expect("FormField was not captured");
        assert_eq!(ff.value_text().as_deref(), Some("Green"));
        match ff.kind {
            FormFieldKind::DropDown { entries, selected } => {
                assert_eq!(entries, vec!["Red", "Green", "Blue"]);
                assert_eq!(selected, 1);
            },
            other => panic!("expected DropDown, got {other:?}"),
        }
    }

    #[test]
    fn test_omml_math_text_is_not_dropped() {
        // issue #270 — every <m:t> inside an m:oMath is real, visible text.
        let xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:d><m:e><m:r><m:t>x</m:t></m:r></m:e><m:e><m:r><m:t>y</m:t></m:r></m:e></m:d>
</m:oMath>
</w:p>"#;
        let p = parse_paragraph_fragment(xml);
        let texts = run_texts(&p).join("");
        assert!(texts.contains('x'), "expected 'x' in {texts:?}");
        assert!(texts.contains('y'), "expected 'y' in {texts:?}");
    }

    #[test]
    fn test_vml_imagedata_is_extracted_as_an_image() {
        // issue #268 — v:imagedata's r:id was only ever read for text
        // boxes; a v:shape wrapping an image, not a text box, vanished.
        let xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <v:shape style="width:100pt;height:50pt">
    <v:imagedata r:id="rId9"/>
  </v:shape>
</w:pict></w:r>
</w:p>"#;
        let p = parse_paragraph_fragment(xml);
        let rid = p.content.iter().find_map(|c| match c {
            ParagraphContent::Run(r) => r.content.iter().find_map(|rc| match rc {
                RunContent::Drawing(d) => Some(d.relationship_id.clone()),
                _ => None,
            }),
            _ => None,
        });
        assert_eq!(rid.as_deref(), Some("rId9"));
    }

    #[test]
    fn test_wordart_textpath_string_is_extracted_as_text() {
        // issue #274 — WordArt's visible text lives in an XML attribute,
        // not element content, so it was never read at all.
        let xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml">
  <v:shape><v:textpath string="WORD-ART"/></v:shape>
</w:pict></w:r>
</w:p>"#;
        let p = parse_paragraph_fragment(xml);
        let texts = run_texts(&p).join("");
        assert!(texts.contains("WORD-ART"), "expected WordArt text in {texts:?}");
    }

    #[test]
    fn test_hyperlink_with_both_rid_and_anchor_keeps_the_external_target() {
        // issues #242, #292 — the anchor branch was checked first and
        // unconditionally taken, discarding a real external r:id whenever
        // a w:anchor fragment was also present.
        let xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:hyperlink r:id="rId7" w:anchor="section1"><w:r><w:t>link</w:t></w:r></w:hyperlink>
</w:p>"#;
        let p = parse_paragraph_fragment(xml);
        let hl = p.content.iter().find_map(|c| match c {
            ParagraphContent::Hyperlink(h) => Some(h.clone()),
            _ => None,
        });
        let hl = hl.expect("hyperlink was not captured");
        match &hl.target {
            HyperlinkTarget::External(rid) => assert_eq!(rid, "rId7"),
            other => panic!("expected an External target carrying r:id, got {other:?}"),
        }
        assert_eq!(hl.fragment.as_deref(), Some("section1"));
    }

    #[test]
    fn test_hyperlink_field_code_instrtext_url_is_captured() {
        // issue #267 — a HYPERLINK expressed via fldChar/instrText (common
        // pandoc/older-tool output) had its URL completely unreachable.
        let xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.com/page" </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>Click here</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>"#;
        let p = parse_paragraph_fragment(xml);
        let hl = p.content.iter().find_map(|c| match c {
            ParagraphContent::Hyperlink(h) => Some(h.clone()),
            _ => None,
        });
        let hl = hl.expect("field-code hyperlink was not resolved");
        match &hl.target {
            HyperlinkTarget::External(url) => assert_eq!(url, "https://example.com/page"),
            other => panic!("expected an External URL target, got {other:?}"),
        }
        let text: String = hl
            .runs
            .iter()
            .flat_map(|r| &r.content)
            .filter_map(|rc| match rc {
                RunContent::Text(t) => Some(t.as_str()),
                _ => None,
            })
            .collect();
        assert_eq!(text, "Click here");
    }

    #[test]
    fn test_sdt_wrapped_table_cell_is_not_dropped() {
        // issue #277 — a cell wrapped in a content control
        // (<w:tr><w:sdt><w:sdtContent><w:tc>) was entirely dropped by the
        // catch-all, shifting every subsequent cell in the row into the
        // wrong column.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:sdt><w:sdtContent><w:tc><w:p><w:r><w:t>SdtCell</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt>
        <w:tc><w:p><w:r><w:t>SecondCell</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#;
        let data = make_minimal_docx(xml);
        let doc = DocxDocument::from_reader(Cursor::new(data)).unwrap();
        let md = doc.to_markdown();
        assert!(
            md.contains("SdtCell") && md.contains("SecondCell"),
            "both cells must survive: {md:?}"
        );
        // Not just present anywhere — in the right columns, not merged/shifted.
        assert!(
            md.contains("| SdtCell | SecondCell |"),
            "cells must stay in their original columns: {md:?}"
        );
    }

    #[test]
    fn test_tracked_change_deleted_cell_is_excluded_from_accepted_view() {
        // issue #266 — a cell marked <w:cellDel> (deleted via tracked
        // changes, pending acceptance) still appeared in the accepted
        // output, same bug w:del was already fixed for at the run level.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:tcPr><w:cellDel w:id="1" w:author="A" w:date="2020-01-01T00:00:00Z"/></w:tcPr>
          <w:p><w:r><w:t>REMOVED CELL</w:t></w:r></w:p>
        </w:tc>
        <w:tc><w:p><w:r><w:t>Kept Cell</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#;
        let data = make_minimal_docx(xml);
        let doc = DocxDocument::from_reader(Cursor::new(data)).unwrap();
        let text = doc.plain_text();
        assert!(
            !text.contains("REMOVED CELL"),
            "deleted cell content must not appear in the accepted view: {text:?}"
        );
        assert!(
            text.contains("Kept Cell"),
            "the surviving cell must still be present: {text:?}"
        );
    }

    #[test]
    fn test_text_box_with_prstgeom_is_not_double_extracted_as_a_phantom_shape() {
        // issue #263 — a <wps:wsp> carrying both <a:prstGeom> (required by
        // schema) and real <wps:txbx> text content produced a content-free
        // phantom Shape(Rect) in addition to the real TextBox.
        let xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<w:r><w:drawing><wp:inline><a:graphic><a:graphicData>
  <wps:wsp>
    <wps:spPr><a:prstGeom prst="rect"/></wps:spPr>
    <wps:txbx><w:txbxContent><w:p><w:r><w:t>7</w:t></w:r></w:p></w:txbxContent></wps:txbx>
  </wps:wsp>
</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
</w:p>"#;
        let p = parse_paragraph_fragment(xml);
        let shape_count = p
            .content
            .iter()
            .filter(|c| {
                matches!(
                    c,
                    ParagraphContent::Run(r)
                        if r.content.iter().any(|rc| matches!(rc, RunContent::Drawing(d) if d.shape.is_some()))
                )
            })
            .count();
        assert_eq!(shape_count, 0, "no phantom Shape should be produced");
        let texts = run_texts(&p);
        let text_box_found = p.content.iter().any(|c| match c {
            ParagraphContent::Run(r) => r
                .content
                .iter()
                .any(|rc| matches!(rc, RunContent::TextBox(_))),
            _ => false,
        });
        assert!(text_box_found, "the real text box must still be present: {texts:?}");
    }

    #[test]
    fn test_footnote_reference_mark_reaches_run_content() {
        // issue #241 — to_ir() carried the note body but lost where in
        // the text it was actually cited.
        let xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:r><w:t>see</w:t></w:r>
<w:r><w:footnoteReference w:id="3"/></w:r>
</w:p>"#;
        let p = parse_paragraph_fragment(xml);
        let found = p.content.iter().any(|c| match c {
            ParagraphContent::Run(r) => r
                .content
                .iter()
                .any(|rc| matches!(rc, RunContent::FootnoteRef(3))),
            _ => false,
        });
        assert!(found, "expected a FootnoteRef(3) in the paragraph content");
    }
}
