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
/// Slide layouts and masters, and the static text they put on slides.
pub mod layout;
/// Inherited text formatting: shape, layout and master list styles, the
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
    HyperlinkInfo, HyperlinkTarget, MediaKind, MediaReference, OleObject, PictureShape,
    PlaceholderInfo, Shape, ShapePosition, TabStop, Table, TableCell, TableRow, TextBody,
    TextContent, TextField, TextParagraph, TextRun, TextSpacing,
};
pub use slide::Slide;

use std::io::{Read, Seek};
use std::path::Path;

use crate::core::opc::OpcReader;
use crate::core::relationships::{Relationships, rel_types};
use crate::core::theme::Theme;
use log::debug;

/// Relationship from the presentation part to the legacy comment-authors
/// part, `ppt/commentAuthors.xml` (ECMA-376 Part 1 §13.3.1).
const REL_COMMENT_AUTHORS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors";
/// Relationship from the presentation part to the modern-comments authors
/// part, `ppt/authors.xml` ([MS-PPTX] authors part).
const REL_MODERN_AUTHORS: &str =
    "http://schemas.microsoft.com/office/2018/10/relationships/authors";

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
    /// Package-level properties beyond core/app: custom properties
    /// (`docProps/custom.xml`), digital-signature presence and thumbnail.
    pub package_properties: crate::core::properties::PackageProperties,
    /// `true` when the presentation part's own relationships include a
    /// `vbaProject` entry — a cheap macro-presence signal, no VBA
    /// interpretation.
    pub has_macros: bool,
    /// Parts that could not be used, as `(part, reason)`: an unreadable or
    /// unresolvable slide (which keeps its place in [`Self::slides`] with
    /// [`Slide::parse_error`] set) or an auxiliary part such as a notes
    /// slide or comments part. Each was skipped so the rest of the deck
    /// could be extracted; a non-empty list means content is missing.
    pub unreadable_parts: Vec<(String, String)>,
    /// The slide layouts the slides use, in first-use order
    /// ([`Slide::layout_index`] points here). Their static text is not part
    /// of any slide's text; see [`Self::static_text_for_slide`].
    pub layouts: Vec<layout::SlideLayout>,
    /// The slide masters those layouts belong to.
    pub masters: Vec<layout::SlideMaster>,
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
        let package_properties = crate::core::properties::read_package_properties(&mut opc);
        let main_part = opc.main_document_part()?;
        let pres_rels = opc.read_rels_for(&main_part)?;
        let has_macros = pres_rels.has_vba_project();

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

        // Degradation policy. A part this reader cannot use — a slide or
        // an auxiliary part hanging off one (notes, comments) — is skipped,
        // logged and recorded in `unreadable_parts`; the rest of the deck is
        // still extracted. A slide keeps its place with a notice. The deck
        // is an error only when no slide at all could be read. (An
        // unreadable notes part used to fail the whole deck while an
        // unresolvable slide relationship vanished silently.)
        let mut unreadable_parts: Vec<(String, String)> = Vec::new();
        let mut record = |part: &str, err: &dyn std::fmt::Display| {
            log::warn!("pptx: skipping unreadable part {part}: {err}");
            unreadable_parts.push((part.to_string(), err.to_string()));
        };

        // Comment author names live in presentation-level parts, keyed by
        // the id each comment carries. Both the legacy and the modern part
        // may be present; their ids (integers vs GUIDs) do not collide.
        let mut comment_authors = std::collections::HashMap::new();
        for rel in pres_rels.all() {
            if rel.rel_type != REL_COMMENT_AUTHORS && rel.rel_type != REL_MODERN_AUTHORS {
                continue;
            }
            let Ok(part) = main_part.resolve_relative(&rel.target) else {
                record(&rel.target, &"unresolvable comment-authors target");
                continue;
            };
            if !opc.has_part(&part) {
                continue;
            }
            match opc.read_part(&part) {
                Ok(data) => comment_authors.extend(slide::parse_comment_authors(&data)),
                Err(e) => record(part.as_str(), &e),
            }
        }

        // Phase 1: gather raw data sequentially (requires &mut opc)
        struct SlideBundle {
            /// The slide's part name, for error reporting.
            part_name: String,
            /// `Err` when the slide part itself could not be read.
            slide_data: std::result::Result<Vec<u8>, String>,
            slide_rels: Relationships,
            notes: Option<NotesBundle>,
            comments_data: Vec<Vec<u8>>,
            /// rId → (raw bytes, format-extension lowercase like "png" / "jpeg").
            /// Pre-resolved here in Phase 1 so the parallel slide parser
            /// (Phase 2) doesn't need access to the OPC reader.
            media: std::collections::HashMap<String, (Vec<u8>, String)>,
            /// rId → text lines of a chart part (`ppt/charts/chartN.xml`)
            /// or SmartArt data part (`ppt/diagrams/dataN.xml`). The slide
            /// XML holds only the reference; the text lives in the part.
            part_text: std::collections::HashMap<String, Vec<String>>,
            /// Index into `layouts` of the slide's layout, when its
            /// layout relationship resolves.
            layout_index: Option<usize>,
        }
        struct NotesBundle {
            part_name: String,
            data: Vec<u8>,
            rels: Relationships,
        }
        impl SlideBundle {
            fn unreadable(part_name: String, err: &str) -> Self {
                SlideBundle {
                    part_name,
                    slide_data: Err(err.to_string()),
                    slide_rels: Relationships::empty(),
                    notes: None,
                    comments_data: Vec::new(),
                    media: std::collections::HashMap::new(),
                    part_text: std::collections::HashMap::new(),
                    layout_index: None,
                }
            }
        }

        // Layouts, masters, images and charts are shared by many slides;
        // each is read and parsed at most once.
        let mut parts = PartCache {
            presentation_default: std::sync::Arc::new(master::parse_default_text_style(&pres_data)),
            ..Default::default()
        };
        let mut bundles = Vec::with_capacity(presentation.slides.len());
        for (slide_idx, slide_id) in presentation.slides.iter().enumerate() {
            // Resolve by rel_id, or by the `slideN.xml` convention when the
            // entry carries no r:id.
            let resolved = if !slide_id.rel_id.is_empty() {
                pres_rels
                    .resolve_target(&slide_id.rel_id, &main_part)
                    .map_err(|e| (format!("slide relationship {}", slide_id.rel_id), e.to_string()))
            } else {
                let candidate = format!("/ppt/slides/slide{}.xml", slide_idx + 1);
                crate::core::opc::PartName::new(&candidate)
                    .map_err(|e| (candidate.clone(), e.to_string()))
            };
            let part_name = match resolved {
                Ok(pn) if opc.has_part(&pn) => pn,
                Ok(pn) => {
                    let name = pn.as_str().to_string();
                    record(&name, &"the part is missing from the package");
                    bundles.push(SlideBundle::unreadable(name, "the part is missing"));
                    continue;
                },
                Err((what, err)) => {
                    record(&what, &err);
                    bundles.push(SlideBundle::unreadable(what, &err));
                    continue;
                },
            };
            let slide_rels = opc
                .read_rels_for(&part_name)
                .unwrap_or_else(|_| Relationships::empty());
            let slide_data = match opc.read_part(&part_name) {
                Ok(d) => Ok(d),
                Err(e) => {
                    record(part_name.as_str(), &e);
                    Err(e.to_string())
                },
            };

            // Slide -> layout -> master, resolved through their own
            // relationships.
            let layout_index = slide_rels
                .first_by_type(rel_types::SLIDE_LAYOUT)
                .and_then(|rel| part_name.resolve_relative(&rel.target).ok())
                .and_then(|layout_part| parts.layout(&mut opc, &layout_part));

            let notes = match slide_rels.first_by_type(rel_types::NOTES_SLIDE) {
                None => None,
                Some(notes_rel) => match part_name.resolve_relative(&notes_rel.target) {
                    Err(e) => {
                        record(&notes_rel.target, &e);
                        None
                    },
                    // A notes relationship to an absent part: nothing to read.
                    Ok(notes_part) if !opc.has_part(&notes_part) => None,
                    Ok(notes_part) => match opc.read_part(&notes_part) {
                        Ok(data) => Some(NotesBundle {
                            part_name: notes_part.as_str().to_string(),
                            data,
                            rels: opc
                                .read_rels_for(&notes_part)
                                .unwrap_or_else(|_| Relationships::empty()),
                        }),
                        Err(e) => {
                            record(notes_part.as_str(), &e);
                            None
                        },
                    },
                },
            };

            // Pictures carry `<a:blip r:embed="rIdN"/>`; charts and SmartArt
            // carry a reference to their own part. Parsing happens in
            // parallel below and can't use the OPC reader, so the bytes
            // and text are materialised here keyed by rId.
            let media = parts.media_for(&mut opc, &part_name, &slide_rels);
            let part_text = parts.part_text_for(&mut opc, &part_name, &slide_rels);

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
                match opc.read_part(&pn) {
                    Ok(data) => comments_data.push(data),
                    Err(e) => record(pn.as_str(), &e),
                }
            }

            bundles.push(SlideBundle {
                part_name: part_name.as_str().to_string(),
                slide_data,
                slide_rels,
                notes,
                comments_data,
                media,
                part_text,
                layout_index,
            });
        }
        let PartCache {
            layouts: layout_entries,
            masters: master_entries,
            ..
        } = parts;

        // Phase 2: parse slides (parallel when feature enabled). Each slide
        // yields its parts that could not be used, merged in order below.
        type ParsedSlide = (Slide, Vec<(String, String)>);
        let parsed = crate::core::parallel::map_collect(
            bundles,
            |b| -> std::result::Result<ParsedSlide, std::convert::Infallible> {
                let mut problems = Vec::new();
                let unreadable_slide = |err: String| Slide {
                    name: b.part_name.clone(),
                    parse_error: Some(err),
                    layout_index: b.layout_index,
                    ..Default::default()
                };
                let data = match &b.slide_data {
                    Ok(d) => d,
                    // Already recorded in phase 1.
                    Err(e) => return Ok((unreadable_slide(e.clone()), problems)),
                };
                let name = xml_csl_name(data);
                let mut parsed =
                    match Slide::parse(data, name, &b.slide_rels, &b.media, &b.part_text) {
                        Ok(s) => s,
                        Err(e) => {
                            problems.push((b.part_name.clone(), e.to_string()));
                            return Ok((unreadable_slide(e.to_string()), problems));
                        },
                    };
                parsed.layout_index = b.layout_index;
                if let Some(notes) = &b.notes {
                    match slide::extract_notes_body(&notes.data, &notes.rels) {
                        Ok(body) => parsed.notes = body,
                        Err(e) => problems.push((notes.part_name.clone(), e.to_string())),
                    }
                }
                for data in &b.comments_data {
                    parsed
                        .comments
                        .extend(slide::parse_comments(data, &comment_authors));
                }
                let chain = b
                    .layout_index
                    .and_then(|i| layout_entries.get(i))
                    .map(|l| &l.chain);
                if let Some(chain) = chain.filter(|c| !c.is_empty()) {
                    apply_inheritance(&mut parsed.shapes, chain);
                }
                Ok((parsed, problems))
            },
        );
        let parsed = match parsed {
            Ok(p) => p,
            Err(never) => match never {},
        };
        let mut slides = Vec::with_capacity(parsed.len());
        for (slide, problems) in parsed {
            for (part, err) in problems {
                record(&part, &err);
            }
            slides.push(slide);
        }
        if !slides.is_empty() && slides.iter().all(|s| s.parse_error.is_some()) {
            let (part, err) = &unreadable_parts[0];
            return Err(crate::core::Error::MalformedXml(format!(
                "no slide could be read; {part}: {err}"
            ))
            .into());
        }
        let layouts: Vec<layout::SlideLayout> =
            layout_entries.into_iter().map(|l| l.info).collect();
        let masters: Vec<layout::SlideMaster> =
            master_entries.into_iter().map(|m| m.info).collect();

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
            package_properties,
            has_macros,
            unreadable_parts,
            layouts,
            masters,
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
                if e.local_name().as_ref() == "cSld" =>
            {
                return crate::core::xml::optional_attr_str(e, "name")
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

/// Fill any unset (`None`) character/paragraph-formatting field on every
/// text shape's runs from the inheritance chain — layout placeholder,
/// master placeholder, master `txStyles`, presentation default — at each
/// paragraph's own outline level, never overriding a value the run,
/// paragraph or shape list style already specified (see [`master`]).
fn apply_inheritance(shapes: &mut [Shape], chain: &master::StyleChain) {
    for shape in shapes {
        match shape {
            Shape::Group(grp) => apply_inheritance(&mut grp.children, chain),
            Shape::AutoShape(auto) => {
                let ph = auto
                    .placeholder
                    .as_ref()
                    .map(|p| (p.ph_type.as_deref(), p.idx));
                let Some(ref mut tb) = auto.text_body else {
                    continue;
                };
                for para in &mut tb.paragraphs {
                    let defaults = chain.resolve(ph, para.level);
                    slide::apply_inherited_defaults(para, &defaults);
                }
            },
            _ => {},
        }
    }
}

/// A resolved slide layout plus the style chain its slides inherit.
struct LayoutEntry {
    info: layout::SlideLayout,
    chain: master::StyleChain,
}

/// A resolved slide master plus what its layouts inherit from it.
struct MasterEntry {
    info: layout::SlideMaster,
    placeholders: std::sync::Arc<Vec<master::PlaceholderStyle>>,
    text_styles: std::sync::Arc<master::MasterTextStyles>,
}

/// Parts shared between slides, each read, decompressed and parsed once:
/// layouts and their relationships, masters, images and chart/SmartArt
/// text. A logo on every slide, or thousands of slides sharing one layout,
/// is the normal shape of a deck.
#[derive(Default)]
struct PartCache {
    presentation_default: std::sync::Arc<master::LevelStyles>,
    layouts: Vec<LayoutEntry>,
    layout_index: std::collections::HashMap<crate::core::opc::PartName, Option<usize>>,
    masters: Vec<MasterEntry>,
    master_index: std::collections::HashMap<crate::core::opc::PartName, Option<usize>>,
    /// Target part → (bytes, extension); `None` when unreadable.
    images: std::collections::HashMap<crate::core::opc::PartName, Option<(Vec<u8>, String)>>,
    /// (target part, is-chart) → text lines.
    part_text: std::collections::HashMap<(crate::core::opc::PartName, bool), Vec<String>>,
    /// Every part this cache decompressed, in order — what the caching
    /// tests count.
    #[cfg(test)]
    reads: std::cell::RefCell<Vec<String>>,
}

impl PartCache {
    /// The index of the layout at `part`, reading it (and its master) the
    /// first time it is seen.
    fn layout<R: Read + Seek>(
        &mut self,
        opc: &mut OpcReader<R>,
        part: &crate::core::opc::PartName,
    ) -> Option<usize> {
        if let Some(&idx) = self.layout_index.get(part) {
            return idx;
        }
        let idx = self.load_layout(opc, part);
        self.layout_index.insert(part.clone(), idx);
        idx
    }

    fn load_layout<R: Read + Seek>(
        &mut self,
        opc: &mut OpcReader<R>,
        part: &crate::core::opc::PartName,
    ) -> Option<usize> {
        if !opc.has_part(part) {
            return None;
        }
        let data = opc.read_part(part).ok()?;
        #[cfg(test)]
        self.reads.borrow_mut().push(part.as_str().to_string());
        let rels = opc
            .read_rels_for(part)
            .unwrap_or_else(|_| Relationships::empty());
        let master_index = rels
            .first_by_type(rel_types::SLIDE_MASTER)
            .and_then(|rel| part.resolve_relative(&rel.target).ok())
            .and_then(|mp| self.master(opc, &mp));
        let (shapes, hide_master) = self.parse_shapes(opc, part, &rels, &data);
        let (master_placeholders, master_text_styles) = match master_index {
            Some(i) => (self.masters[i].placeholders.clone(), self.masters[i].text_styles.clone()),
            None => Default::default(),
        };
        self.layouts.push(LayoutEntry {
            info: layout::SlideLayout {
                part_name: part.as_str().to_string(),
                name: xml_csl_name(&data),
                shapes,
                master_index,
                show_master_shapes: !hide_master,
            },
            chain: master::StyleChain {
                layout_placeholders: std::sync::Arc::new(master::parse_placeholder_styles(&data)),
                master_placeholders,
                master_text_styles,
                presentation_default: self.presentation_default.clone(),
            },
        });
        Some(self.layouts.len() - 1)
    }

    fn master<R: Read + Seek>(
        &mut self,
        opc: &mut OpcReader<R>,
        part: &crate::core::opc::PartName,
    ) -> Option<usize> {
        if let Some(&idx) = self.master_index.get(part) {
            return idx;
        }
        let idx = (|| {
            if !opc.has_part(part) {
                return None;
            }
            let data = opc.read_part(part).ok()?;
            #[cfg(test)]
            self.reads.borrow_mut().push(part.as_str().to_string());
            let rels = opc
                .read_rels_for(part)
                .unwrap_or_else(|_| Relationships::empty());
            let (shapes, _) = self.parse_shapes(opc, part, &rels, &data);
            self.masters.push(MasterEntry {
                info: layout::SlideMaster {
                    part_name: part.as_str().to_string(),
                    name: xml_csl_name(&data),
                    shapes,
                },
                placeholders: std::sync::Arc::new(master::parse_placeholder_styles(&data)),
                text_styles: std::sync::Arc::new(master::parse_master_text_styles(&data)),
            });
            Some(self.masters.len() - 1)
        })();
        self.master_index.insert(part.clone(), idx);
        idx
    }

    /// The shapes of a layout or master part, and its `showMasterSp="0"`.
    /// Unparseable static content is skipped: the slides are what matter.
    fn parse_shapes<R: Read + Seek>(
        &mut self,
        opc: &mut OpcReader<R>,
        part: &crate::core::opc::PartName,
        rels: &Relationships,
        data: &[u8],
    ) -> (Vec<Shape>, bool) {
        let media = self.media_for(opc, part, rels);
        let part_text = self.part_text_for(opc, part, rels);
        match Slide::parse(data, String::new(), rels, &media, &part_text) {
            Ok(s) => (s.shapes, s.hide_master_shapes),
            Err(e) => {
                log::warn!("pptx: shapes of {part} are unreadable: {e}");
                (Vec::new(), false)
            },
        }
    }

    /// rId → (bytes, extension) for every image relationship of `source`.
    fn media_for<R: Read + Seek>(
        &mut self,
        opc: &mut OpcReader<R>,
        source: &crate::core::opc::PartName,
        rels: &Relationships,
    ) -> std::collections::HashMap<String, (Vec<u8>, String)> {
        let mut media = std::collections::HashMap::new();
        for rel in rels.all() {
            if rel.rel_type != rel_types::IMAGE
                || rel.target_mode == crate::core::relationships::TargetMode::External
            {
                continue;
            }
            let Ok(target) = source.resolve_relative(&rel.target) else {
                continue;
            };
            let entry = self.images.entry(target).or_insert_with_key(|target| {
                if !opc.has_part(target) {
                    return None;
                }
                let bytes = opc.read_part(target).ok()?;
                #[cfg(test)]
                self.reads.borrow_mut().push(target.as_str().to_string());
                let ext = std::path::Path::new(target.as_str())
                    .extension()
                    .and_then(|s| s.to_str())
                    .map(|s| s.to_ascii_lowercase())
                    .unwrap_or_else(|| guess_format_from_bytes(&bytes).to_string());
                Some((bytes, ext))
            });
            if let Some(found) = entry {
                media.insert(rel.id.clone(), found.clone());
            }
        }
        media
    }

    /// rId → text lines for every chart and SmartArt data relationship of
    /// `source`.
    fn part_text_for<R: Read + Seek>(
        &mut self,
        opc: &mut OpcReader<R>,
        source: &crate::core::opc::PartName,
        rels: &Relationships,
    ) -> std::collections::HashMap<String, Vec<String>> {
        let mut out = std::collections::HashMap::new();
        for rel in rels.all() {
            let is_chart = rel.rel_type == rel_types::CHART;
            if !is_chart && rel.rel_type != rel_types::DIAGRAM_DATA {
                continue;
            }
            let Ok(target) = source.resolve_relative(&rel.target) else {
                continue;
            };
            let lines =
                self.part_text
                    .entry((target, is_chart))
                    .or_insert_with_key(|(target, _)| {
                        if !opc.has_part(target) {
                            return Vec::new();
                        }
                        let Ok(data) = opc.read_part(target) else {
                            return Vec::new();
                        };
                        #[cfg(test)]
                        self.reads.borrow_mut().push(target.as_str().to_string());
                        if is_chart {
                            crate::core::chart::chart_text_lines(&data)
                        } else {
                            slide::diagram_data_text_lines(&data)
                        }
                    });
            if !lines.is_empty() {
                out.insert(rel.id.clone(), lines.clone());
            }
        }
        out
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
        minimal_package_with(content_type, &[], "")
    }

    /// As `minimal_package`, plus extra `(name, bytes)` entries and extra
    /// `<Relationship …/>` elements in the presentation part's own rels.
    fn minimal_package_with(
        content_type: &str,
        extra: &[(&str, &[u8])],
        pres_rels: &str,
    ) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts: zip::write::SimpleFileOptions = zip::write::SimpleFileOptions::default();
        for (name, data) in extra {
            zip.start_file(*name, opts).unwrap();
            zip.write_all(data).unwrap();
        }
        if !pres_rels.is_empty() {
            zip.start_file("ppt/_rels/presentation.xml.rels", opts)
                .unwrap();
            zip.write_all(
                format!(
                    r#"<?xml version="1.0"?><Relationships
                         xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{pres_rels}</Relationships>"#
                )
                .as_bytes(),
            )
            .unwrap();
        }

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

    /// `docProps/app.xml` is read on open for presentations too — the
    /// slide/notes/hidden-slide counts are the ones no other part carries.
    #[test]
    fn test_app_properties_are_read_on_open() {
        let app_xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Company>Acme Corp</Company>
  <Slides>12</Slides>
</Properties>"#;
        let bytes = minimal_package_with(
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            &[("docProps/app.xml", app_xml)],
            "",
        );
        let doc = super::PptxDocument::from_reader(Cursor::new(bytes)).unwrap();
        let app = doc
            .app_properties
            .expect("app_properties must be populated");
        assert_eq!(app.company.as_deref(), Some("Acme Corp"));
        assert_eq!(app.slides, Some(12));
    }

    /// The macro-presence signal is the presentation part's `vbaProject`
    /// relationship, under the type PowerPoint writes.
    #[test]
    fn test_vba_project_relationship_sets_has_macros() {
        let bytes = minimal_package_with(
            "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
            &[("ppt/vbaProject.bin", b"fake vba bytes")],
            r#"<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>"#,
        );
        let doc = super::PptxDocument::from_reader(Cursor::new(bytes)).unwrap();
        assert!(doc.has_macros);
        assert!(crate::convert_pptx::pptx_to_ir(&doc).metadata.has_macros);
        let plain = super::PptxDocument::from_reader(Cursor::new(minimal_package(
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
        )))
        .unwrap();
        assert!(!plain.has_macros);
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

#[cfg(test)]
mod part_cache_tests {
    use std::io::{Cursor, Write};

    use super::*;
    use crate::core::opc::PartName;

    fn rels(entries: &[(&str, &str, &str)]) -> Vec<u8> {
        let mut xml = String::from(
            r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        for (id, ty, target) in entries {
            xml.push_str(&format!(r#"<Relationship Id="{id}" Type="{ty}" Target="{target}"/>"#));
        }
        xml.push_str("</Relationships>");
        xml.into_bytes()
    }

    /// Three slides sharing one layout (and so one master), one image and
    /// one chart.
    fn shared_parts_package() -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = zip::write::SimpleFileOptions::default();
        let mut add = |name: &str, data: &[u8]| {
            zip.start_file(name, opts).unwrap();
            zip.write_all(data).unwrap();
        };
        add(
            "[Content_Types].xml",
            br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>"#,
        );
        let ns = r#"xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#;
        for n in 1..=3 {
            add(
                &format!("ppt/slides/slide{n}.xml"),
                format!(r#"<p:sld {ns}><p:cSld><p:spTree/></p:cSld></p:sld>"#).as_bytes(),
            );
            add(
                &format!("ppt/slides/_rels/slide{n}.xml.rels"),
                &rels(&[
                    ("rId1", rel_types::SLIDE_LAYOUT, "../slideLayouts/slideLayout1.xml"),
                    ("rId2", rel_types::IMAGE, "../media/logo.png"),
                    ("rId3", rel_types::CHART, "../charts/chart1.xml"),
                ]),
            );
        }
        add(
            "ppt/slideLayouts/slideLayout1.xml",
            format!(r#"<p:sldLayout {ns}><p:cSld><p:spTree/></p:cSld></p:sldLayout>"#).as_bytes(),
        );
        add(
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
            &rels(&[("rId1", rel_types::SLIDE_MASTER, "../slideMasters/slideMaster1.xml")]),
        );
        add(
            "ppt/slideMasters/slideMaster1.xml",
            format!(r#"<p:sldMaster {ns}><p:cSld><p:spTree/></p:cSld></p:sldMaster>"#).as_bytes(),
        );
        add("ppt/media/logo.png", &[0x89, b'P', b'N', b'G', 0, 0, 0, 0]);
        add(
            "ppt/charts/chart1.xml",
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>"#,
        );
        zip.finish().unwrap().into_inner()
    }

    /// The layout's relationships were re-read and every referenced image
    /// and chart re-decompressed once per slide; a logo on 500 slides was
    /// decompressed 500 times. Each shared part is now read once.
    #[test]
    fn test_shared_layout_image_and_chart_parts_are_read_once() {
        let mut opc = OpcReader::new(Cursor::new(shared_parts_package())).unwrap();
        let mut parts = PartCache::default();
        for n in 1..=3 {
            let slide = PartName::new(&format!("/ppt/slides/slide{n}.xml")).unwrap();
            let rels = opc.read_rels_for(&slide).unwrap();
            let layout = rels
                .first_by_type(rel_types::SLIDE_LAYOUT)
                .and_then(|r| slide.resolve_relative(&r.target).ok())
                .unwrap();
            assert_eq!(parts.layout(&mut opc, &layout), Some(0));
            let media = parts.media_for(&mut opc, &slide, &rels);
            assert_eq!(media["rId2"].1, "png");
            let text = parts.part_text_for(&mut opc, &slide, &rels);
            assert_eq!(text["rId3"], vec!["Title: Revenue".to_string()]);
        }
        let reads = parts.reads.borrow();
        let mut sorted = reads.clone();
        sorted.sort();
        assert_eq!(
            sorted,
            [
                "/ppt/charts/chart1.xml",
                "/ppt/media/logo.png",
                "/ppt/slideLayouts/slideLayout1.xml",
                "/ppt/slideMasters/slideMaster1.xml",
            ],
            "each shared part is read exactly once"
        );
        assert_eq!(parts.layouts.len(), 1);
        assert_eq!(parts.masters.len(), 1);
    }

    /// Byte-sniffed image formats, used when a relationship target has no
    /// extension.
    #[test]
    fn test_guess_format_from_bytes() {
        for (bytes, want) in [
            (&[0x89, b'P', b'N', b'G'][..], "png"),
            (&[0xFF, 0xD8, 0xFF, 0xE0], "jpeg"),
            (b"GIF89a..", "gif"),
            (b"GIF87a..", "gif"),
            (b"BM......", "bmp"),
            (&[0xD7, 0xCD, 0xC6, 0x9A], "wmf"),
            (&[0x01, 0x00, 0x00, 0x00], "emf"),
            (b"II*\0....", "tiff"),
            (b"MM\0*....", "tiff"),
            (b"????", "png"),
        ] {
            assert_eq!(guess_format_from_bytes(bytes), want, "{bytes:?}");
        }
    }
}
