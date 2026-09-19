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
/// Slide-master `<p:txStyles>` default formatting — a placeholder's
/// fallback when its own direct formatting leaves a property unset.
pub(crate) mod master;
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
    AutoShape, BulletStyle, ConnectorShape, GraphicContent, GraphicFrame, GroupShape,
    HyperlinkInfo, HyperlinkTarget, PictureShape, PlaceholderInfo, Shape, ShapePosition, Table,
    TableCell, TableRow, TextBody, TextContent, TextField, TextParagraph, TextRun,
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
    /// Font programs found under `ppt/fonts/`. Each entry is
    /// `(font_name, ttf_or_otf_bytes)`. PDF→PPTX→PDF round-trips use
    /// these to preserve the source typeface (mirrors the DOCX side).
    pub embedded_fonts: Vec<(String, Vec<u8>)>,
    /// Parsed `docProps/core.xml`. `None` when the package carries no
    /// core-properties part.
    pub core_properties: Option<crate::core::properties::CoreProperties>,
    /// Parsed `docProps/app.xml` (company, producing application, template,
    /// slide/notes/hidden-slide counts). `None` when the package carries no
    /// extended-properties part.
    pub app_properties: Option<crate::core::properties::AppProperties>,
    /// `true` when the presentation part's own relationships include a
    /// `vbaProject` entry — a cheap macro-presence signal, no VBA
    /// interpretation.
    pub has_macros: bool,
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
    pub fn from_reader<R: Read + Seek>(mut reader: R) -> Result<Self> {
        // A password-protected PPTX is a CFB container, not a zip at all.
        // See the identical check in docx::DocxDocument::from_reader.
        if crate::cfb::is_cfb_container(&mut reader).map_err(crate::core::Error::from)? {
            return Err(crate::core::Error::Unsupported(
                "the file is a password-protected (encrypted) OOXML package; \
                 decryption is not supported"
                    .into(),
            )
            .into());
        }
        let opc = OpcReader::new(reader)?;
        Self::from_opc(opc)
    }

    fn from_opc<R: Read + Seek>(mut opc: OpcReader<R>) -> Result<Self> {
        debug!("PptxDocument: parsing started");
        opc.verify_main_content_type(
            &[
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml",
                "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
                "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
                "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml",
                // `.potm` (macro-enabled template) was missing, so every
                // real .potm failed with FormatMismatch.
                "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
            ],
            "a PresentationML presentation",
        )?;
        let core_properties = crate::core::properties::read_core_properties(&mut opc);
        let app_properties = crate::core::properties::read_app_properties(&mut opc);
        let main_part = opc.main_document_part()?;
        let pres_rels = opc.read_rels_for(&main_part)?;
        let has_macros = pres_rels.first_by_type(rel_types::VBA_PROJECT).is_some();

        // Parse theme
        // See the DOCX reader: a malformed theme is not a reason to refuse
        // the whole presentation.
        let theme = pres_rels
            .first_by_type(rel_types::THEME)
            .and_then(|rel| main_part.resolve_relative(&rel.target).ok())
            .filter(|pn| opc.has_part(pn))
            .and_then(|pn| opc.read_part(&pn).ok())
            .and_then(|data| match Theme::parse(&data) {
                Ok(t) => Some(t),
                Err(e) => {
                    debug!("PptxDocument: ignoring unreadable theme part: {e}");
                    None
                },
            });

        // Parse presentation.xml
        let pres_data = opc.read_part(&main_part)?;
        let presentation = PresentationInfo::parse(&pres_data)?;

        // Phase 1: gather raw data sequentially (requires &mut opc)
        struct SlideBundle {
            slide_data: Vec<u8>,
            slide_rels: Relationships,
            notes_data: Option<Vec<u8>>,
            comments_data: Vec<Vec<u8>>,
            /// rId → (raw bytes, format-extension lowercase like "png" / "jpeg").
            /// Pre-resolved here in Phase 1 so the parallel slide parser
            /// (Phase 2) doesn't need access to the OPC reader.
            media: std::collections::HashMap<String, (Vec<u8>, String)>,
            /// rId → extracted chart text lines. A `<c:chart r:id="…"/>`
            /// in the slide XML holds no text of its own — the title,
            /// axis labels, category names and cached data values live in
            /// the separate part that id resolves to
            /// (`ppt/charts/chartN.xml`), which nothing opened at all
            /// before. Pre-resolved here for the same reason
            /// `media` is.
            charts: std::collections::HashMap<String, Vec<String>>,
            /// This slide's resolved master title/body level-0 defaults,
            /// via its layout's own `SLIDE_MASTER` relationship — `None`
            /// when the layout/master chain can't be resolved at all.
            master_styles: Option<master::MasterTextStyles>,
        }
        // Many slides share one layout/master — resolve and parse each
        // unique master part at most once.
        let mut master_styles_cache: std::collections::HashMap<String, master::MasterTextStyles> =
            std::collections::HashMap::new();
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

            // Slide -> layout -> master, resolved through their own
            // relationships exactly like every other part this reader
            // already follows (images, notes, charts) — never opened at
            // all before this.
            let master_styles = slide_rels
                .first_by_type(rel_types::SLIDE_LAYOUT)
                .and_then(|rel| part_name.resolve_relative(&rel.target).ok())
                .filter(|pn| opc.has_part(pn))
                .and_then(|layout_part| {
                    let layout_rels = opc.read_rels_for(&layout_part).ok()?;
                    let master_rel = layout_rels.first_by_type(rel_types::SLIDE_MASTER)?;
                    let master_part = layout_part.resolve_relative(&master_rel.target).ok()?;
                    if !opc.has_part(&master_part) {
                        return None;
                    }
                    let key = master_part.as_str().to_string();
                    if let Some(cached) = master_styles_cache.get(&key) {
                        return Some(cached.clone());
                    }
                    let data = opc.read_part(&master_part).ok()?;
                    let styles = master::parse_master_text_styles(&data);
                    master_styles_cache.insert(key, styles.clone());
                    Some(styles)
                })
                .filter(|s| !s.is_empty());

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

            // Pre-load all IMAGE-relationship parts the slide references.
            // PPTX picture frames carry `<a:blip r:embed="rIdN"/>`; the
            // relationship resolves to a part like `/ppt/media/image3.png`.
            // Parsing happens in parallel below and can't use the OPC
            // reader, so we materialise the bytes here keyed by rId.
            let mut media = std::collections::HashMap::new();
            for rel in slide_rels.all() {
                if rel.rel_type != rel_types::IMAGE {
                    continue;
                }
                let target = match part_name.resolve_relative(&rel.target) {
                    Ok(t) => t,
                    Err(_) => continue,
                };
                if !opc.has_part(&target) {
                    continue;
                }
                let bytes = match opc.read_part(&target) {
                    Ok(b) => b,
                    Err(_) => continue,
                };
                let ext = std::path::Path::new(&rel.target)
                    .extension()
                    .and_then(|s| s.to_str())
                    .map(|s| s.to_ascii_lowercase())
                    .unwrap_or_else(|| guess_format_from_bytes(&bytes).to_string());
                media.insert(rel.id.clone(), (bytes, ext));
            }

            // Pre-load and extract every embedded chart part the slide
            // references.
            let mut charts = std::collections::HashMap::new();
            for rel in slide_rels.all() {
                if rel.rel_type != rel_types::CHART {
                    continue;
                }
                let target = match part_name.resolve_relative(&rel.target) {
                    Ok(t) => t,
                    Err(_) => continue,
                };
                if !opc.has_part(&target) {
                    continue;
                }
                let Ok(data) = opc.read_part(&target) else {
                    continue;
                };
                let lines = crate::core::chart::chart_text_lines(&data);
                if !lines.is_empty() {
                    charts.insert(rel.id.clone(), lines);
                }
            }

            // Comments hang off the slide's own relationships, both in the
            // legacy `comments` form and the newer `authors`+`modernComment`
            // pair. Neither part was ever read, so review notes on a deck
            // reached no consumer at all.
            let comment_parts: Vec<_> = slide_rels
                .all()
                .iter()
                .filter(|rel| rel.rel_type.ends_with("/comments"))
                .filter_map(|rel| part_name.resolve_relative(&rel.target).ok())
                .filter(|pn| opc.has_part(pn))
                .collect();
            let mut comments_data: Vec<Vec<u8>> = Vec::new();
            for pn in comment_parts {
                if let Ok(data) = opc.read_part(&pn) {
                    comments_data.push(data);
                }
            }

            bundles.push(SlideBundle {
                slide_data,
                slide_rels,
                notes_data,
                comments_data,
                media,
                charts,
                master_styles,
            });
        }

        // Phase 2: parse slides (parallel when feature enabled)
        let slides = crate::core::parallel::map_collect(bundles, |b| -> Result<Slide> {
            let name = xml_csl_name(&b.slide_data);
            let mut parsed =
                Slide::parse(&b.slide_data, name, &b.slide_rels, &b.media, &b.charts)?;
            if let Some(notes_data) = &b.notes_data {
                parsed.notes = extract_notes_body(notes_data);
            }
            for data in &b.comments_data {
                parsed.comments.extend(slide::parse_comments(data));
            }
            if let Some(ref styles) = b.master_styles {
                apply_master_inheritance(&mut parsed.shapes, styles);
            }
            Ok(parsed)
        })?;

        // Scan `ppt/fonts/` for embedded font programs. Mirrors the DOCX
        // reader (`word/fonts/`).
        let mut embedded_fonts: Vec<(String, Vec<u8>)> = Vec::new();
        for name in opc.part_names() {
            let s = name.to_string();
            if !s.starts_with("/ppt/fonts/") {
                continue;
            }
            let lower = s.to_lowercase();
            if !(lower.ends_with(".ttf") || lower.ends_with(".otf")) {
                continue;
            }
            if let Ok(data) = opc.read_part(&name) {
                let basename = s.rsplit('/').next().unwrap_or("font");
                let face = crate::docx::strip_embedded_font_filename(basename);
                let font_name = if face.is_empty() {
                    basename.to_string()
                } else {
                    face
                };
                embedded_fonts.push((font_name, data));
            }
        }

        debug!(
            "PptxDocument: {} slides parsed, {} embedded fonts",
            slides.len(),
            embedded_fonts.len()
        );
        Ok(PptxDocument {
            presentation,
            slides,
            theme,
            embedded_fonts,
            core_properties,
            app_properties,
            has_macros,
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

/// Extract the speaker notes body from a notes slide XML. Finds the
/// body placeholder (type="body") and returns its structured `TextBody`.
fn extract_notes_body(xml_data: &[u8]) -> Option<TextBody> {
    slide::extract_notes_body(xml_data)
}

/// Fill any unset (`None`) character/paragraph-formatting field on
/// every Title/Body placeholder shape's runs from the resolved slide
/// master's level-0 defaults — never overriding a field the run/
/// paragraph already specified directly (the PPTX analogue
/// of the legacy `.ppt` master-inheritance fix).
fn apply_master_inheritance(shapes: &mut [Shape], styles: &master::MasterTextStyles) {
    for shape in shapes {
        match shape {
            Shape::Group(grp) => apply_master_inheritance(&mut grp.children, styles),
            Shape::AutoShape(auto) => {
                let ph_type = auto.placeholder.as_ref().and_then(|p| p.ph_type.as_deref());
                let defaults = match ph_type {
                    Some("title" | "ctrTitle") => styles.title.as_ref(),
                    Some("body" | "subTitle") | None if auto.placeholder.is_some() => {
                        styles.body.as_ref()
                    },
                    _ => None,
                };
                let Some(defaults) = defaults else { continue };
                let Some(ref mut tb) = auto.text_body else { continue };
                for para in &mut tb.paragraphs {
                    if para.alignment.is_none() {
                        para.alignment = defaults.alignment.clone();
                    }
                    for content in &mut para.content {
                        if let shape::TextContent::Run(run) = content {
                            if run.bold.is_none() {
                                run.bold = defaults.bold;
                            }
                            if run.italic.is_none() {
                                run.italic = defaults.italic;
                            }
                            if run.underline.is_none() {
                                run.underline = defaults.underline.clone();
                            }
                            if run.font_size_hundredths_pt.is_none() {
                                run.font_size_hundredths_pt = defaults.font_size_hundredths_pt;
                            }
                            if run.color_rgb.is_none() {
                                run.color_rgb = defaults.color_rgb;
                            }
                        }
                    }
                }
            },
            _ => {},
        }
    }
}

/// Best-effort image-format detection from the raw bytes.
///
/// Used as a fallback when the relationship target has no recognisable
/// extension (rare — DrawingML images almost always carry one). Returns
/// a lowercase extension string suitable for round-tripping back into
/// `office_oxide::ir::ImageFormat::extension()`.
fn guess_format_from_bytes(bytes: &[u8]) -> &'static str {
    if bytes.starts_with(&[0x89, b'P', b'N', b'G']) {
        "png"
    } else if bytes.starts_with(&[0xFF, 0xD8, 0xFF]) {
        "jpeg"
    } else if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        "gif"
    } else if bytes.starts_with(b"BM") {
        "bmp"
    } else if bytes.len() >= 4 && bytes.starts_with(&[0xD7, 0xCD, 0xC6, 0x9A]) {
        "wmf"
    } else if bytes.len() >= 4 && bytes.starts_with(&[0x01, 0x00, 0x00, 0x00]) {
        "emf"
    } else if bytes.len() >= 4 && (bytes.starts_with(b"II*\0") || bytes.starts_with(b"MM\0*")) {
        "tiff"
    } else {
        "png"
    }
}

impl crate::core::OfficeDocument for PptxDocument {
    fn plain_text(&self) -> String {
        self.plain_text()
    }

    fn to_markdown(&self) -> String {
        self.to_markdown()
    }
}

#[cfg(test)]
mod content_type_tests {
    use std::io::{Cursor, Write};

    use crate::core::relationships::rel_types;

    /// Build a minimal PresentationML package whose main part carries
    /// `content_type`.
    fn minimal_package(content_type: &str) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts: zip::write::SimpleFileOptions = zip::write::SimpleFileOptions::default();

        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(
            format!(
                r#"<?xml version="1.0"?><Types
                     xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                   <Override PartName="/ppt/presentation.xml" ContentType="{content_type}"/>
                 </Types>"#
            )
            .as_bytes(),
        )
        .unwrap();

        zip.start_file("_rels/.rels", opts).unwrap();
        zip.write_all(
            format!(
                r#"<?xml version="1.0"?><Relationships
                     xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                   <Relationship Id="rId1" Type="{}" Target="ppt/presentation.xml"/>
                 </Relationships>"#,
                rel_types::OFFICE_DOCUMENT
            )
            .as_bytes(),
        )
        .unwrap();

        zip.start_file("ppt/presentation.xml", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0"?><p:presentation
                  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                <p:sldIdLst/></p:presentation>"#,
        )
        .unwrap();

        zip.finish().unwrap().into_inner()
    }

    /// `.potm`'s real content type was missing from the whitelist, so every
    /// macro-enabled PowerPoint template failed with a format mismatch.
    #[test]
    fn test_potm_content_type_accepted() {
        for ct in [
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
            "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
        ] {
            let bytes = minimal_package(ct);
            super::PptxDocument::from_reader(Cursor::new(bytes))
                .unwrap_or_else(|e| panic!("content type {ct} should be accepted, got {e}"));
        }
        // A non-PresentationML main part is still refused.
        let bytes = minimal_package(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        );
        assert!(super::PptxDocument::from_reader(Cursor::new(bytes)).is_err());
    }

    /// Same gap as DOCX/XLSX, confirmed independently for PPTX.
    #[test]
    fn test_encrypted_pptx_gives_a_friendly_error_via_the_format_specific_reader() {
        let mut cfb = vec![0u8; 512];
        cfb[0..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
        let err = super::PptxDocument::from_reader(Cursor::new(cfb)).unwrap_err();
        let msg = err.to_string();
        assert!(
            msg.contains("password-protected"),
            "expected a friendly password-protected message, got: {msg}"
        );
    }
}
