//! PPTX creation (write) module.
//!
//! Provides a builder API for creating PPTX files from scratch.
//!
//! # Example
//!
//! ```rust,no_run
//! use office_oxide::pptx::write::{PptxWriter, Run};
//!
//! let mut writer = PptxWriter::new();
//! writer.add_slide()
//!     .set_title("Hello")
//!     .add_text("World")
//!     .add_rich_text(&[
//!         Run::new("Bold").bold(),
//!         Run::new(" and ").into(),
//!         Run::new("red").color("FF0000"),
//!     ])
//!     .add_bullet_list(&["First", "Second", "Third"])
//!     .add_text_box("Note", 1_000_000, 5_000_000, 3_000_000, 500_000);
//! writer.save("output.pptx").unwrap();
//! ```

use std::io::{Seek, Write};
use std::path::Path;

use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};

use crate::core::opc::{OpcWriter, PartName};
use crate::core::relationships::rel_types;

use super::Result;

// ---------------------------------------------------------------------------
// Content types
// ---------------------------------------------------------------------------

const CT_PRESENTATION: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const CT_SLIDE: &str = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const CT_SLIDE_LAYOUT: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
const CT_SLIDE_MASTER: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";

// ---------------------------------------------------------------------------
// Namespaces
// ---------------------------------------------------------------------------

use crate::core::xml::ns::{DRAWING_ML_STR as NS_DML, PML_STR as NS_PML, R_STR as NS_REL};

// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/// A styled text run for a PPTX paragraph.
///
/// # Example
/// ```rust,no_run
/// use office_oxide::pptx::write::Run;
///
/// let r = Run::new("Highlighted").bold().color("FFCC00").font_size(18.0);
/// ```
#[derive(Debug, Clone, Default)]
pub struct Run {
    /// The text content of this run.
    pub text: String,
    /// Apply bold weight.
    pub bold: bool,
    /// Apply italic style.
    pub italic: bool,
    /// Apply single underline.
    pub underline: bool,
    /// Apply strikethrough.
    pub strikethrough: bool,
    /// 6-char hex color string, e.g. `"FF0000"` (no leading `#`).
    pub color: Option<String>,
    /// Font size in points, e.g. `18.0`.
    pub font_size_pt: Option<f64>,
    /// Font name, e.g. `"Calibri"`.
    pub font_name: Option<String>,
}

impl Run {
    /// Create a plain text run.
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            ..Default::default()
        }
    }

    /// Enable bold weight.
    pub fn bold(mut self) -> Self {
        self.bold = true;
        self
    }
    /// Enable italic style.
    pub fn italic(mut self) -> Self {
        self.italic = true;
        self
    }
    /// Enable single underline.
    pub fn underline(mut self) -> Self {
        self.underline = true;
        self
    }
    /// Enable strikethrough.
    pub fn strikethrough(mut self) -> Self {
        self.strikethrough = true;
        self
    }

    /// Font color as a 6-char hex string (no `#`).
    pub fn color(mut self, hex: impl Into<String>) -> Self {
        self.color = Some(hex.into());
        self
    }

    /// Font size in points.
    pub fn font_size(mut self, pt: f64) -> Self {
        self.font_size_pt = Some(pt);
        self
    }

    /// Font family name.
    pub fn font(mut self, name: impl Into<String>) -> Self {
        self.font_name = Some(name.into());
        self
    }

    fn has_rpr(&self) -> bool {
        self.bold
            || self.italic
            || self.underline
            || self.strikethrough
            || self.color.is_some()
            || self.font_size_pt.is_some()
            || self.font_name.is_some()
    }
}

impl From<&str> for Run {
    fn from(s: &str) -> Self {
        Self::new(s)
    }
}

impl From<String> for Run {
    fn from(s: String) -> Self {
        Self::new(s)
    }
}

// ---------------------------------------------------------------------------
// Internal body content model
// ---------------------------------------------------------------------------

/// Paragraph-level properties carried through a `BodyItem::RichText`.
/// Present so the writer can emit `<a:pPr>` attributes (alignment,
/// space-before) that don't fit on per-run `<a:rPr>`.
#[derive(Debug, Clone, Default)]
pub struct ParaProps {
    /// Paragraph alignment written as `<a:pPr algn="…"/>`. `None`
    /// leaves the renderer-default left alignment in place.
    pub alignment: Option<crate::ir::ParagraphAlignment>,
    /// Space before the paragraph in points × 100. 1250 = 12.5pt.
    /// When set, written as `<a:spcBef><a:spcPts val="…"/></a:spcBef>`.
    pub space_before_hundredths_pt: Option<u32>,
}

#[derive(Debug, Clone)]
enum BodyItem {
    Text(String),
    RichText(Vec<Run>, ParaProps),
    BulletList(Vec<String>),
    /// Free-floating text box: (runs, x_emu, y_emu, cx_emu, cy_emu)
    TextBox(Vec<Run>, i64, i64, i64, i64),
    /// Embedded image: (data, format, x_emu, y_emu, cx_emu, cy_emu)
    Image(Vec<u8>, crate::ir::ImageFormat, i64, i64, u64, u64),
}

// ---------------------------------------------------------------------------
// SlideData
// ---------------------------------------------------------------------------

/// Data for a single slide being constructed.
#[derive(Debug, Clone)]
pub struct SlideData {
    /// The slide title (if set).
    pub title: Option<String>,
    /// Optional explicit alignment for the title placeholder. None
    /// leaves alignment to the slide layout default (typically
    /// centered for title placeholders).
    pub title_alignment: Option<crate::ir::ParagraphAlignment>,
    body_items: Vec<BodyItem>,
}

impl SlideData {
    fn new() -> Self {
        Self {
            title: None,
            title_alignment: None,
            body_items: Vec::new(),
        }
    }

    /// Set the slide title. Overwrites any previously set title.
    pub fn set_title(&mut self, title: &str) -> &mut Self {
        self.title = Some(title.to_string());
        self
    }

    /// Set the slide title and its alignment. Overwrites any
    /// previously set title.
    pub fn set_title_aligned(
        &mut self,
        title: &str,
        alignment: Option<crate::ir::ParagraphAlignment>,
    ) -> &mut Self {
        self.title = Some(title.to_string());
        self.title_alignment = alignment;
        self
    }

    /// Add a plain text paragraph to the body area.
    pub fn add_text(&mut self, text: &str) -> &mut Self {
        self.body_items.push(BodyItem::Text(text.to_string()));
        self
    }

    /// Add a paragraph of styled [`Run`]s to the body area.
    pub fn add_rich_text(&mut self, runs: &[Run]) -> &mut Self {
        self.body_items
            .push(BodyItem::RichText(runs.to_vec(), ParaProps::default()));
        self
    }

    /// Add a paragraph of styled [`Run`]s with an explicit alignment.
    pub fn add_rich_text_aligned(
        &mut self,
        runs: &[Run],
        alignment: Option<crate::ir::ParagraphAlignment>,
    ) -> &mut Self {
        self.body_items.push(BodyItem::RichText(
            runs.to_vec(),
            ParaProps {
                alignment,
                ..Default::default()
            },
        ));
        self
    }

    /// Add a paragraph of styled [`Run`]s with full paragraph
    /// properties (alignment, space-before).
    pub fn add_rich_text_with_props(&mut self, runs: &[Run], props: ParaProps) -> &mut Self {
        self.body_items
            .push(BodyItem::RichText(runs.to_vec(), props));
        self
    }

    /// Add a bullet list to the body area.
    pub fn add_bullet_list(&mut self, items: &[&str]) -> &mut Self {
        let owned: Vec<String> = items.iter().map(|s| s.to_string()).collect();
        self.body_items.push(BodyItem::BulletList(owned));
        self
    }

    /// Add a free-floating text box at an absolute position.
    ///
    /// All dimensions are in EMU (English Metric Units).
    /// 1 inch = 914 400 EMU; 1 cm ≈ 360 000 EMU.
    pub fn add_text_box(&mut self, text: &str, x: i64, y: i64, cx: i64, cy: i64) -> &mut Self {
        self.body_items
            .push(BodyItem::TextBox(vec![Run::new(text)], x, y, cx, cy));
        self
    }

    /// Add a free-floating text box with styled [`Run`]s.
    pub fn add_rich_text_box(
        &mut self,
        runs: &[Run],
        x: i64,
        y: i64,
        cx: i64,
        cy: i64,
    ) -> &mut Self {
        self.body_items
            .push(BodyItem::TextBox(runs.to_vec(), x, y, cx, cy));
        self
    }

    /// Embed an image at an absolute position on the slide.
    ///
    /// All coordinates are in EMU (English Metric Units; 914 400 EMU = 1 inch).
    pub fn add_image(
        &mut self,
        data: Vec<u8>,
        format: crate::ir::ImageFormat,
        x: i64,
        y: i64,
        cx: u64,
        cy: u64,
    ) -> &mut Self {
        self.body_items
            .push(BodyItem::Image(data, format, x, y, cx, cy));
        self
    }

    fn has_placeholder_body(&self) -> bool {
        self.body_items
            .iter()
            .any(|i| !matches!(i, BodyItem::TextBox(..) | BodyItem::Image(..)))
    }
}

// ---------------------------------------------------------------------------
// PptxWriter
// ---------------------------------------------------------------------------

/// Builder for creating PPTX files from scratch.
pub struct PptxWriter {
    slides: Vec<SlideData>,
    /// Presentation width in EMU (default: 12 192 000 — standard 16:9).
    cx: u64,
    /// Presentation height in EMU (default: 6 858 000 — standard 16:9).
    cy: u64,
    /// Embedded font programs to ship inside the package under `ppt/fonts/`.
    /// Mirrors `DocxWriter::embed_font` semantics: each `(name, bytes)` pair
    /// becomes one font part, used by PDF↔PPTX round-trips to preserve the
    /// source typeface.
    embedded_fonts: Vec<(String, Vec<u8>)>,
    /// Document metadata for `docProps/core.xml`. `None` means no
    /// core-properties part is written.
    metadata: Option<crate::ir::Metadata>,
}

impl PptxWriter {
    /// Create a new empty PPTX writer.
    pub fn new() -> Self {
        Self {
            slides: Vec::new(),
            cx: 12_192_000,
            cy: 6_858_000,
            embedded_fonts: Vec::new(),
            metadata: None,
        }
    }

    /// Set document metadata (written to `docProps/core.xml`).
    pub fn set_metadata(&mut self, meta: &crate::ir::Metadata) -> &mut Self {
        self.metadata = Some(meta.clone());
        self
    }

    /// Embed a font program (TrueType / OpenType bytes) under `ppt/fonts/`.
    ///
    /// `name` is used for both the on-disk file name and the human-readable
    /// font name in the presentation's font table. Deduplication is by
    /// `name` only — supplying different bytes for an already-registered
    /// name is a no-op. Pass distinct names (e.g. `Calibri-Bold` vs
    /// `Calibri`) when you need to ship multiple faces of the same family.
    pub fn embed_font(&mut self, name: impl Into<String>, data: Vec<u8>) -> &mut Self {
        let name = name.into();
        if !self.embedded_fonts.iter().any(|(n, _)| n == &name) {
            self.embedded_fonts.push((name, data));
        }
        self
    }

    /// Override the presentation canvas size (in EMU).
    ///
    /// Call before adding slides. 914 400 EMU = 1 inch.
    pub fn set_presentation_size(&mut self, cx: u64, cy: u64) -> &mut Self {
        self.cx = cx;
        self.cy = cy;
        self
    }

    /// Add a new slide and return a mutable reference for configuration.
    pub fn add_slide(&mut self) -> &mut SlideData {
        self.slides.push(SlideData::new());
        self.slides.last_mut().expect("just pushed")
    }

    /// Add a slide and return its 0-based index (for use with index-based API).
    pub fn add_slide_get_index(&mut self) -> usize {
        self.slides.push(SlideData::new());
        self.slides.len() - 1
    }

    /// Set the slide title by slide index.
    pub fn slide_set_title(&mut self, slide: usize, title: &str) {
        if let Some(s) = self.slides.get_mut(slide) {
            s.set_title(title);
        }
    }

    /// Add a plain text paragraph to the slide body by slide index.
    pub fn slide_add_text(&mut self, slide: usize, text: &str) {
        if let Some(s) = self.slides.get_mut(slide) {
            s.add_text(text);
        }
    }

    /// Embed an image on a slide by slide index.
    pub fn slide_add_image(
        &mut self,
        slide: usize,
        data: Vec<u8>,
        format: crate::ir::ImageFormat,
        x: i64,
        y: i64,
        cx: u64,
        cy: u64,
    ) {
        if let Some(s) = self.slides.get_mut(slide) {
            s.add_image(data, format, x, y, cx, cy);
        }
    }

    /// Save the presentation to a file path.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        let opc = OpcWriter::create(path)?;
        self.write_opc(opc)?;
        Ok(())
    }

    /// Write the presentation to any `Write + Seek` destination.
    pub fn write_to<W: Write + Seek>(&self, writer: W) -> Result<()> {
        let opc = OpcWriter::new(writer)?;
        self.write_opc(opc)?;
        Ok(())
    }

    fn write_opc<W: Write + Seek>(&self, mut opc: OpcWriter<W>) -> Result<()> {
        let pres_part = PartName::new("/ppt/presentation.xml")?;
        let master_part = PartName::new("/ppt/slideMasters/slideMaster1.xml")?;
        let layout_part = PartName::new("/ppt/slideLayouts/slideLayout1.xml")?;

        opc.add_package_rel(rel_types::OFFICE_DOCUMENT, "ppt/presentation.xml");
        opc.add_part_rel(&pres_part, rel_types::SLIDE_MASTER, "slideMasters/slideMaster1.xml");

        // Core properties (docProps/core.xml). Written only when the
        // caller supplied metadata so files generated through the
        // existing `add_slide` API stay byte-identical when no
        // metadata was set.
        if let Some(ref meta) = self.metadata {
            let core_part = PartName::new("/docProps/core.xml")?;
            opc.add_package_rel(rel_types::CORE_PROPERTIES, "docProps/core.xml");
            let core_xml = crate::core::core_properties::generate_xml(meta);
            opc.add_part(&core_part, crate::core::core_properties::CONTENT_TYPE, &core_xml)?;
        }

        let mut slide_parts = Vec::with_capacity(self.slides.len());
        for i in 0..self.slides.len() {
            let idx = i + 1;
            let slide_part = PartName::new(&format!("/ppt/slides/slide{idx}.xml"))?;
            opc.add_part_rel(&pres_part, rel_types::SLIDE, &format!("slides/slide{idx}.xml"));
            slide_parts.push(slide_part);
        }

        opc.add_part_rel(&master_part, rel_types::SLIDE_LAYOUT, "../slideLayouts/slideLayout1.xml");

        let pres_xml = generate_presentation_xml(self.slides.len(), self.cx, self.cy);
        opc.add_part(&pres_part, CT_PRESENTATION, &pres_xml)?;

        let master_xml = generate_slide_master_xml();
        opc.add_part(&master_part, CT_SLIDE_MASTER, &master_xml)?;

        let layout_xml = generate_slide_layout_xml();
        opc.add_part(&layout_part, CT_SLIDE_LAYOUT, &layout_xml)?;

        let mut global_img_idx = 1u32;
        for (i, slide) in self.slides.iter().enumerate() {
            let slide_part = &slide_parts[i];

            // rId1 = slide layout
            opc.add_part_rel(
                slide_part,
                rel_types::SLIDE_LAYOUT,
                "../slideLayouts/slideLayout1.xml",
            );

            // rId2+ = one per embedded image
            let mut img_rids: Vec<(String, i64, i64, u64, u64)> = Vec::new();
            for item in &slide.body_items {
                if let BodyItem::Image(data, fmt, x, y, cx, cy) = item {
                    let rid = format!("rId{}", img_rids.len() + 2);
                    let ext = fmt.extension();
                    opc.add_part_rel(
                        slide_part,
                        rel_types::IMAGE,
                        &format!("../media/image{global_img_idx}.{ext}"),
                    );
                    let media_part =
                        PartName::new(&format!("/ppt/media/image{global_img_idx}.{ext}"))?;
                    opc.add_part(&media_part, fmt.content_type(), data)?;
                    img_rids.push((rid, *x, *y, *cx, *cy));
                    global_img_idx += 1;
                }
            }

            let slide_xml = generate_slide_xml(slide, &img_rids);
            opc.add_part(slide_part, CT_SLIDE, &slide_xml)?;
        }

        // Embed fonts under `ppt/fonts/font_<n>_<safe_name>.ttf`. Mirrors
        // the DOCX `word/fonts/` layout. Other PowerPoint software may not
        // honor this without the full presentation-relationship machinery
        // for `<p:embeddedFontLst>`, but the in-process reader scans the
        // directory directly so PDF↔PPTX round-trips preserve fonts.
        crate::core::embedded_fonts::write_embedded_fonts(
            &mut opc,
            "/ppt/fonts/",
            &self.embedded_fonts,
        )?;

        opc.finish()?;
        Ok(())
    }
}

impl Default for PptxWriter {
    fn default() -> Self {
        Self::new()
    }
}

// ---------------------------------------------------------------------------
// XML generation helpers
// ---------------------------------------------------------------------------

fn write_decl(w: &mut Writer<Vec<u8>>) {
    w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .expect("write decl");
}

fn write_text_element(w: &mut Writer<Vec<u8>>, tag: &str, text: &str) {
    w.write_event(Event::Start(BytesStart::new(tag)))
        .expect("write start");
    w.write_event(Event::Text(BytesText::new(text)))
        .expect("write text");
    w.write_event(Event::End(BytesEnd::new(tag)))
        .expect("write end");
}

fn write_empty(w: &mut Writer<Vec<u8>>, tag: &str) {
    w.write_event(Event::Empty(BytesStart::new(tag)))
        .expect("write empty");
}

fn pml_root(tag: &str) -> BytesStart<'_> {
    let mut elem = BytesStart::new(tag);
    elem.push_attribute(("xmlns:p", NS_PML));
    elem.push_attribute(("xmlns:a", NS_DML));
    elem.push_attribute(("xmlns:r", NS_REL));
    elem
}

fn write_nv_grp_sp_pr(w: &mut Writer<Vec<u8>>) {
    w.write_event(Event::Start(BytesStart::new("p:nvGrpSpPr")))
        .expect("write");
    let mut cnv_pr = BytesStart::new("p:cNvPr");
    cnv_pr.push_attribute(("id", "1"));
    cnv_pr.push_attribute(("name", ""));
    w.write_event(Event::Empty(cnv_pr)).expect("write");
    write_empty(w, "p:cNvGrpSpPr");
    write_empty(w, "p:nvPr");
    w.write_event(Event::End(BytesEnd::new("p:nvGrpSpPr")))
        .expect("write");
}

// Write a DrawingML run (<a:r>) with optional rPr.
fn write_dml_run(w: &mut Writer<Vec<u8>>, run: &Run) {
    w.write_event(Event::Start(BytesStart::new("a:r")))
        .expect("write");

    if run.has_rpr() {
        let mut rpr = BytesStart::new("a:rPr");
        rpr.push_attribute(("lang", "en-US"));
        rpr.push_attribute(("dirty", "0"));
        if run.bold {
            rpr.push_attribute(("b", "1"));
        }
        if run.italic {
            rpr.push_attribute(("i", "1"));
        }
        if run.underline {
            rpr.push_attribute(("u", "sng"));
        }
        if run.strikethrough {
            rpr.push_attribute(("strike", "sngStrike"));
        }
        if let Some(pt) = run.font_size_pt {
            // DrawingML stores size in hundredths of a point
            let hundredths = (pt * 100.0).round() as u32;
            rpr.push_attribute(("sz", hundredths.to_string().as_str()));
        }

        if run.color.is_some() || run.font_name.is_some() {
            w.write_event(Event::Start(rpr)).expect("write rPr start");

            if let Some(ref hex) = run.color {
                w.write_event(Event::Start(BytesStart::new("a:solidFill")))
                    .expect("write");
                let mut clr = BytesStart::new("a:srgbClr");
                clr.push_attribute(("val", hex.as_str()));
                w.write_event(Event::Empty(clr)).expect("write");
                w.write_event(Event::End(BytesEnd::new("a:solidFill")))
                    .expect("write");
            }

            if let Some(ref name) = run.font_name {
                let mut latin = BytesStart::new("a:latin");
                latin.push_attribute(("typeface", name.as_str()));
                w.write_event(Event::Empty(latin)).expect("write");
            }

            w.write_event(Event::End(BytesEnd::new("a:rPr")))
                .expect("write rPr end");
        } else {
            w.write_event(Event::Empty(rpr)).expect("write rPr empty");
        }
    }

    write_text_element(w, "a:t", &run.text);
    w.write_event(Event::End(BytesEnd::new("a:r")))
        .expect("write");
}

// ---------------------------------------------------------------------------
// presentation.xml
// ---------------------------------------------------------------------------

fn generate_presentation_xml(slide_count: usize, cx: u64, cy: u64) -> Vec<u8> {
    let mut w = Writer::new(Vec::new());
    write_decl(&mut w);

    w.write_event(Event::Start(pml_root("p:presentation")))
        .expect("write");

    w.write_event(Event::Start(BytesStart::new("p:sldMasterIdLst")))
        .expect("write");
    let mut master_id = BytesStart::new("p:sldMasterId");
    master_id.push_attribute(("id", "2147483648"));
    master_id.push_attribute(("r:id", "rId1"));
    w.write_event(Event::Empty(master_id)).expect("write");
    w.write_event(Event::End(BytesEnd::new("p:sldMasterIdLst")))
        .expect("write");

    w.write_event(Event::Start(BytesStart::new("p:sldIdLst")))
        .expect("write");
    for i in 0..slide_count {
        let slide_id_val = 256 + i as u32;
        let r_id = format!("rId{}", i + 2);
        let mut slide_id = BytesStart::new("p:sldId");
        slide_id.push_attribute(("id", slide_id_val.to_string().as_str()));
        slide_id.push_attribute(("r:id", r_id.as_str()));
        w.write_event(Event::Empty(slide_id)).expect("write");
    }
    w.write_event(Event::End(BytesEnd::new("p:sldIdLst")))
        .expect("write");

    let mut sld_sz = BytesStart::new("p:sldSz");
    sld_sz.push_attribute(("cx", cx.to_string().as_str()));
    sld_sz.push_attribute(("cy", cy.to_string().as_str()));
    w.write_event(Event::Empty(sld_sz)).expect("write");

    // notesSz: PowerPoint expects this even when there are no notes
    // pages. Standard default is the same dimensions as the slide.
    let mut notes_sz = BytesStart::new("p:notesSz");
    notes_sz.push_attribute(("cx", cx.to_string().as_str()));
    notes_sz.push_attribute(("cy", cy.to_string().as_str()));
    w.write_event(Event::Empty(notes_sz)).expect("write");

    // defaultTextStyle: empty list of paragraph-level defaults is
    // legal and silences PowerPoint's "Reset Layout" command failure
    // when the user opens the deck.
    w.write_event(Event::Start(BytesStart::new("p:defaultTextStyle")))
        .expect("write");
    w.write_event(Event::End(BytesEnd::new("p:defaultTextStyle")))
        .expect("write");

    w.write_event(Event::End(BytesEnd::new("p:presentation")))
        .expect("write");
    w.into_inner()
}

// ---------------------------------------------------------------------------
// slideMasters/slideMaster1.xml
// ---------------------------------------------------------------------------

fn generate_slide_master_xml() -> Vec<u8> {
    let mut w = Writer::new(Vec::new());
    write_decl(&mut w);

    w.write_event(Event::Start(pml_root("p:sldMaster")))
        .expect("write");
    w.write_event(Event::Start(BytesStart::new("p:cSld")))
        .expect("write");
    w.write_event(Event::Start(BytesStart::new("p:spTree")))
        .expect("write");
    write_nv_grp_sp_pr(&mut w);
    write_empty(&mut w, "p:grpSpPr");
    w.write_event(Event::End(BytesEnd::new("p:spTree")))
        .expect("write");
    w.write_event(Event::End(BytesEnd::new("p:cSld")))
        .expect("write");

    w.write_event(Event::Start(BytesStart::new("p:sldLayoutIdLst")))
        .expect("write");
    let mut layout_id = BytesStart::new("p:sldLayoutId");
    layout_id.push_attribute(("id", "2147483649"));
    layout_id.push_attribute(("r:id", "rId1"));
    w.write_event(Event::Empty(layout_id)).expect("write");
    w.write_event(Event::End(BytesEnd::new("p:sldLayoutIdLst")))
        .expect("write");

    w.write_event(Event::End(BytesEnd::new("p:sldMaster")))
        .expect("write");
    w.into_inner()
}

// ---------------------------------------------------------------------------
// slideLayouts/slideLayout1.xml
// ---------------------------------------------------------------------------

fn generate_slide_layout_xml() -> Vec<u8> {
    let mut w = Writer::new(Vec::new());
    write_decl(&mut w);

    // Type "obj" = "Title and Content" — PowerPoint's standard
    // layout. Slides referencing this layout get a sized title
    // placeholder at the top and a body placeholder filling the
    // rest. Was `type="blank"` with empty spTree; that left
    // PowerPoint guessing at placeholder geometry.
    let mut root = pml_root("p:sldLayout");
    root.push_attribute(("type", "obj"));
    root.push_attribute(("preserve", "1"));
    w.write_event(Event::Start(root)).expect("write");
    w.write_event(Event::Start(BytesStart::new("p:cSld")))
        .expect("write");
    w.write_event(Event::Start(BytesStart::new("p:spTree")))
        .expect("write");
    write_nv_grp_sp_pr(&mut w);
    write_empty(&mut w, "p:grpSpPr");

    // Title placeholder — top of slide, ~5 % top inset, full width minus margin.
    write_layout_placeholder(
        &mut w,
        2,
        "Title 1",
        "title",
        None,
        // Geometry in EMU. Standard 16:9 @ 12 192 000 × 6 858 000:
        // place title at (914 400, 685 800) ≈ 1 in × 0.75 in,
        // size 10 363 200 × 1 143 000 ≈ 11.3 in × 1.25 in.
        Some((914_400, 685_800, 10_363_200, 1_143_000)),
    );

    // Body placeholder — fills the area below the title.
    write_layout_placeholder(
        &mut w,
        3,
        "Body 2",
        "body",
        Some(1),
        Some((914_400, 1_905_000, 10_363_200, 4_343_400)),
    );

    w.write_event(Event::End(BytesEnd::new("p:spTree")))
        .expect("write");
    w.write_event(Event::End(BytesEnd::new("p:cSld")))
        .expect("write");
    w.write_event(Event::End(BytesEnd::new("p:sldLayout")))
        .expect("write");
    w.into_inner()
}

/// Emit one placeholder `<p:sp>` inside the slide layout: an empty
/// shape carrying the placeholder type/idx + its xfrm rectangle.
/// Slides that reference this layout's `type` and `idx` inherit the
/// geometry — without it PowerPoint falls back to bare-default
/// positioning that often pushes content off the slide canvas.
fn write_layout_placeholder(
    w: &mut Writer<Vec<u8>>,
    id: u32,
    name: &str,
    ph_type: &str,
    ph_idx: Option<u32>,
    geometry_emu: Option<(i64, i64, i64, i64)>, // (x, y, cx, cy)
) {
    let id_str = id.to_string();
    w.write_event(Event::Start(BytesStart::new("p:sp")))
        .expect("sp start");

    w.write_event(Event::Start(BytesStart::new("p:nvSpPr")))
        .expect("nvSpPr start");
    let mut cnv_pr = BytesStart::new("p:cNvPr");
    cnv_pr.push_attribute(("id", id_str.as_str()));
    cnv_pr.push_attribute(("name", name));
    w.write_event(Event::Empty(cnv_pr)).expect("cNvPr");
    w.write_event(Event::Start(BytesStart::new("p:cNvSpPr")))
        .expect("cNvSpPr start");
    let mut locks = BytesStart::new("a:spLocks");
    locks.push_attribute(("noGrp", "1"));
    w.write_event(Event::Empty(locks)).expect("spLocks");
    w.write_event(Event::End(BytesEnd::new("p:cNvSpPr")))
        .expect("cNvSpPr end");
    w.write_event(Event::Start(BytesStart::new("p:nvPr")))
        .expect("nvPr start");
    let mut ph = BytesStart::new("p:ph");
    ph.push_attribute(("type", ph_type));
    let idx_buf;
    if let Some(idx) = ph_idx {
        idx_buf = idx.to_string();
        ph.push_attribute(("idx", idx_buf.as_str()));
    }
    w.write_event(Event::Empty(ph)).expect("ph");
    w.write_event(Event::End(BytesEnd::new("p:nvPr")))
        .expect("nvPr end");
    w.write_event(Event::End(BytesEnd::new("p:nvSpPr")))
        .expect("nvSpPr end");

    // spPr with optional xfrm geometry
    if let Some((x, y, cx, cy)) = geometry_emu {
        w.write_event(Event::Start(BytesStart::new("p:spPr")))
            .expect("spPr start");
        w.write_event(Event::Start(BytesStart::new("a:xfrm")))
            .expect("xfrm start");
        let mut off = BytesStart::new("a:off");
        let xs = x.to_string();
        let ys = y.to_string();
        off.push_attribute(("x", xs.as_str()));
        off.push_attribute(("y", ys.as_str()));
        w.write_event(Event::Empty(off)).expect("off");
        let mut ext = BytesStart::new("a:ext");
        let cxs = cx.to_string();
        let cys = cy.to_string();
        ext.push_attribute(("cx", cxs.as_str()));
        ext.push_attribute(("cy", cys.as_str()));
        w.write_event(Event::Empty(ext)).expect("ext");
        w.write_event(Event::End(BytesEnd::new("a:xfrm")))
            .expect("xfrm end");
        w.write_event(Event::End(BytesEnd::new("p:spPr")))
            .expect("spPr end");
    } else {
        write_empty(w, "p:spPr");
    }

    // Empty txBody — slides supply their own text.
    w.write_event(Event::Start(BytesStart::new("p:txBody")))
        .expect("txBody start");
    write_empty(w, "a:bodyPr");
    write_empty(w, "a:lstStyle");
    w.write_event(Event::Start(BytesStart::new("a:p")))
        .expect("a:p start");
    w.write_event(Event::End(BytesEnd::new("a:p")))
        .expect("a:p end");
    w.write_event(Event::End(BytesEnd::new("p:txBody")))
        .expect("txBody end");

    w.write_event(Event::End(BytesEnd::new("p:sp")))
        .expect("sp end");
}

// ---------------------------------------------------------------------------
// slides/slideN.xml
// ---------------------------------------------------------------------------

fn generate_slide_xml(slide: &SlideData, img_rids: &[(String, i64, i64, u64, u64)]) -> Vec<u8> {
    let mut w = Writer::new(Vec::new());
    write_decl(&mut w);

    w.write_event(Event::Start(pml_root("p:sld")))
        .expect("write");
    w.write_event(Event::Start(BytesStart::new("p:cSld")))
        .expect("write");
    w.write_event(Event::Start(BytesStart::new("p:spTree")))
        .expect("write");

    write_nv_grp_sp_pr(&mut w);
    write_empty(&mut w, "p:grpSpPr");

    let mut next_id: u32 = 2;

    if let Some(ref title) = slide.title {
        write_title_shape(&mut w, next_id, title, slide.title_alignment.as_ref());
        next_id += 1;
    }

    if slide.has_placeholder_body() {
        let placeholder_items: Vec<&BodyItem> = slide
            .body_items
            .iter()
            .filter(|i| !matches!(i, BodyItem::TextBox(..) | BodyItem::Image(..)))
            .collect();
        write_body_shape(&mut w, next_id, &placeholder_items);
        next_id += 1;
    }

    // Free-floating text boxes
    for item in &slide.body_items {
        if let BodyItem::TextBox(runs, x, y, cx, cy) = item {
            write_text_box_shape(&mut w, next_id, runs, *x, *y, *cx, *cy);
            next_id += 1;
        }
    }

    // Embedded images
    for (rid, x, y, cx, cy) in img_rids {
        write_pic_shape(&mut w, next_id, rid, *x, *y, *cx, *cy);
        next_id += 1;
    }

    w.write_event(Event::End(BytesEnd::new("p:spTree")))
        .expect("write");
    w.write_event(Event::End(BytesEnd::new("p:cSld")))
        .expect("write");
    w.write_event(Event::End(BytesEnd::new("p:sld")))
        .expect("write");

    w.into_inner()
}

fn write_title_shape(
    w: &mut Writer<Vec<u8>>,
    id: u32,
    title: &str,
    alignment: Option<&crate::ir::ParagraphAlignment>,
) {
    let id_str = id.to_string();
    w.write_event(Event::Start(BytesStart::new("p:sp")))
        .expect("write");

    w.write_event(Event::Start(BytesStart::new("p:nvSpPr")))
        .expect("write");
    let mut cnv_pr = BytesStart::new("p:cNvPr");
    cnv_pr.push_attribute(("id", id_str.as_str()));
    cnv_pr.push_attribute(("name", "Title 1"));
    w.write_event(Event::Empty(cnv_pr)).expect("write");
    w.write_event(Event::Start(BytesStart::new("p:cNvSpPr")))
        .expect("write");
    let mut locks = BytesStart::new("a:spLocks");
    locks.push_attribute(("noGrp", "1"));
    w.write_event(Event::Empty(locks)).expect("write");
    w.write_event(Event::End(BytesEnd::new("p:cNvSpPr")))
        .expect("write");
    w.write_event(Event::Start(BytesStart::new("p:nvPr")))
        .expect("write");
    let mut ph = BytesStart::new("p:ph");
    ph.push_attribute(("type", "title"));
    w.write_event(Event::Empty(ph)).expect("write");
    w.write_event(Event::End(BytesEnd::new("p:nvPr")))
        .expect("write");
    w.write_event(Event::End(BytesEnd::new("p:nvSpPr")))
        .expect("write");

    write_empty(w, "p:spPr");

    w.write_event(Event::Start(BytesStart::new("p:txBody")))
        .expect("write");
    write_empty(w, "a:bodyPr");
    if let Some(a) = alignment {
        let runs = vec![Run::new(title)];
        let props = ParaProps {
            alignment: Some(a.clone()),
            ..Default::default()
        };
        write_rich_paragraph(w, &runs, &props);
    } else {
        write_plain_paragraph(w, title);
    }
    w.write_event(Event::End(BytesEnd::new("p:txBody")))
        .expect("write");

    w.write_event(Event::End(BytesEnd::new("p:sp")))
        .expect("write");
}

fn write_body_shape(w: &mut Writer<Vec<u8>>, id: u32, items: &[&BodyItem]) {
    let id_str = id.to_string();
    w.write_event(Event::Start(BytesStart::new("p:sp")))
        .expect("write");

    w.write_event(Event::Start(BytesStart::new("p:nvSpPr")))
        .expect("write");
    let mut cnv_pr = BytesStart::new("p:cNvPr");
    cnv_pr.push_attribute(("id", id_str.as_str()));
    cnv_pr.push_attribute(("name", "Body 2"));
    w.write_event(Event::Empty(cnv_pr)).expect("write");
    w.write_event(Event::Start(BytesStart::new("p:cNvSpPr")))
        .expect("write");
    let mut locks = BytesStart::new("a:spLocks");
    locks.push_attribute(("noGrp", "1"));
    w.write_event(Event::Empty(locks)).expect("write");
    w.write_event(Event::End(BytesEnd::new("p:cNvSpPr")))
        .expect("write");
    w.write_event(Event::Start(BytesStart::new("p:nvPr")))
        .expect("write");
    let mut ph = BytesStart::new("p:ph");
    ph.push_attribute(("type", "body"));
    ph.push_attribute(("idx", "1"));
    w.write_event(Event::Empty(ph)).expect("write");
    w.write_event(Event::End(BytesEnd::new("p:nvPr")))
        .expect("write");
    w.write_event(Event::End(BytesEnd::new("p:nvSpPr")))
        .expect("write");

    write_empty(w, "p:spPr");

    w.write_event(Event::Start(BytesStart::new("p:txBody")))
        .expect("write");
    // <a:bodyPr><a:normAutofit/></a:bodyPr>: tell PowerPoint to
    // shrink-to-fit the body text. Without this, dense PDF pages
    // imported as slides overflow the placeholder and content
    // renders off-slide.
    w.write_event(Event::Start(BytesStart::new("a:bodyPr")))
        .expect("write bodyPr start");
    write_empty(w, "a:normAutofit");
    w.write_event(Event::End(BytesEnd::new("a:bodyPr")))
        .expect("write bodyPr end");

    for item in items {
        match item {
            BodyItem::Text(text) => write_plain_paragraph(w, text),
            BodyItem::RichText(runs, props) => write_rich_paragraph(w, runs, props),
            BodyItem::BulletList(bullets) => {
                for bullet in bullets {
                    write_bullet_paragraph(w, bullet);
                }
            },
            BodyItem::TextBox(..) | BodyItem::Image(..) => {}, // handled separately
        }
    }

    w.write_event(Event::End(BytesEnd::new("p:txBody")))
        .expect("write");
    w.write_event(Event::End(BytesEnd::new("p:sp")))
        .expect("write");
}

fn write_text_box_shape(
    w: &mut Writer<Vec<u8>>,
    id: u32,
    runs: &[Run],
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
) {
    let id_str = id.to_string();
    let name = format!("TextBox {id}");

    w.write_event(Event::Start(BytesStart::new("p:sp")))
        .expect("write");

    // nvSpPr — non-visual properties (txBox=1 = free-floating text box)
    w.write_event(Event::Start(BytesStart::new("p:nvSpPr")))
        .expect("write");
    let mut cnv_pr = BytesStart::new("p:cNvPr");
    cnv_pr.push_attribute(("id", id_str.as_str()));
    cnv_pr.push_attribute(("name", name.as_str()));
    w.write_event(Event::Empty(cnv_pr)).expect("write");
    let mut cnv_sp_pr = BytesStart::new("p:cNvSpPr");
    cnv_sp_pr.push_attribute(("txBox", "1"));
    w.write_event(Event::Empty(cnv_sp_pr)).expect("write");
    write_empty(w, "p:nvPr");
    w.write_event(Event::End(BytesEnd::new("p:nvSpPr")))
        .expect("write");

    // spPr — shape properties with position and size
    w.write_event(Event::Start(BytesStart::new("p:spPr")))
        .expect("write");

    w.write_event(Event::Start(BytesStart::new("a:xfrm")))
        .expect("write");
    let mut off = BytesStart::new("a:off");
    off.push_attribute(("x", x.to_string().as_str()));
    off.push_attribute(("y", y.to_string().as_str()));
    w.write_event(Event::Empty(off)).expect("write");
    let mut ext = BytesStart::new("a:ext");
    ext.push_attribute(("cx", cx.to_string().as_str()));
    ext.push_attribute(("cy", cy.to_string().as_str()));
    w.write_event(Event::Empty(ext)).expect("write");
    w.write_event(Event::End(BytesEnd::new("a:xfrm")))
        .expect("write");

    let mut geom = BytesStart::new("a:prstGeom");
    geom.push_attribute(("prst", "rect"));
    w.write_event(Event::Start(geom)).expect("write");
    write_empty(w, "a:avLst");
    w.write_event(Event::End(BytesEnd::new("a:prstGeom")))
        .expect("write");

    w.write_event(Event::End(BytesEnd::new("p:spPr")))
        .expect("write");

    // txBody — `wrap="none"` plus explicit zero insets so callers
    // sizing the shape rectangle to the exact text bbox (e.g. the
    // PDF→PPTX layout path) get the text rendered without
    // PowerPoint's default ~0.1" left/right padding silently eating
    // shape width and forcing visible glyph re-wrapping.
    w.write_event(Event::Start(BytesStart::new("p:txBody")))
        .expect("write");
    let mut body_pr = BytesStart::new("a:bodyPr");
    body_pr.push_attribute(("wrap", "none"));
    body_pr.push_attribute(("lIns", "0"));
    body_pr.push_attribute(("tIns", "0"));
    body_pr.push_attribute(("rIns", "0"));
    body_pr.push_attribute(("bIns", "0"));
    w.write_event(Event::Empty(body_pr)).expect("write");
    write_rich_paragraph(w, runs, &ParaProps::default());
    w.write_event(Event::End(BytesEnd::new("p:txBody")))
        .expect("write");

    w.write_event(Event::End(BytesEnd::new("p:sp")))
        .expect("write");
}

fn write_pic_shape(w: &mut Writer<Vec<u8>>, id: u32, rid: &str, x: i64, y: i64, cx: u64, cy: u64) {
    let id_str = id.to_string();
    let name = format!("Image {id}");

    w.write_event(Event::Start(BytesStart::new("p:pic")))
        .expect("write");

    w.write_event(Event::Start(BytesStart::new("p:nvPicPr")))
        .expect("write");
    let mut cnv_pr = BytesStart::new("p:cNvPr");
    cnv_pr.push_attribute(("id", id_str.as_str()));
    cnv_pr.push_attribute(("name", name.as_str()));
    w.write_event(Event::Empty(cnv_pr)).expect("write");
    write_empty(w, "p:cNvPicPr");
    write_empty(w, "p:nvPr");
    w.write_event(Event::End(BytesEnd::new("p:nvPicPr")))
        .expect("write");

    w.write_event(Event::Start(BytesStart::new("p:blipFill")))
        .expect("write");
    let mut blip = BytesStart::new("a:blip");
    blip.push_attribute(("r:embed", rid));
    w.write_event(Event::Empty(blip)).expect("write");
    w.write_event(Event::Start(BytesStart::new("a:stretch")))
        .expect("write");
    write_empty(w, "a:fillRect");
    w.write_event(Event::End(BytesEnd::new("a:stretch")))
        .expect("write");
    w.write_event(Event::End(BytesEnd::new("p:blipFill")))
        .expect("write");

    w.write_event(Event::Start(BytesStart::new("p:spPr")))
        .expect("write");
    w.write_event(Event::Start(BytesStart::new("a:xfrm")))
        .expect("write");
    let mut off = BytesStart::new("a:off");
    off.push_attribute(("x", x.to_string().as_str()));
    off.push_attribute(("y", y.to_string().as_str()));
    w.write_event(Event::Empty(off)).expect("write");
    let mut ext = BytesStart::new("a:ext");
    ext.push_attribute(("cx", cx.to_string().as_str()));
    ext.push_attribute(("cy", cy.to_string().as_str()));
    w.write_event(Event::Empty(ext)).expect("write");
    w.write_event(Event::End(BytesEnd::new("a:xfrm")))
        .expect("write");
    let mut geom = BytesStart::new("a:prstGeom");
    geom.push_attribute(("prst", "rect"));
    w.write_event(Event::Start(geom)).expect("write");
    write_empty(w, "a:avLst");
    w.write_event(Event::End(BytesEnd::new("a:prstGeom")))
        .expect("write");
    w.write_event(Event::End(BytesEnd::new("p:spPr")))
        .expect("write");

    w.write_event(Event::End(BytesEnd::new("p:pic")))
        .expect("write");
}

fn write_plain_paragraph(w: &mut Writer<Vec<u8>>, text: &str) {
    w.write_event(Event::Start(BytesStart::new("a:p")))
        .expect("write");
    w.write_event(Event::Start(BytesStart::new("a:r")))
        .expect("write");
    write_text_element(w, "a:t", text);
    w.write_event(Event::End(BytesEnd::new("a:r")))
        .expect("write");
    w.write_event(Event::End(BytesEnd::new("a:p")))
        .expect("write");
}

fn write_rich_paragraph(w: &mut Writer<Vec<u8>>, runs: &[Run], props: &ParaProps) {
    use crate::ir::ParagraphAlignment;
    w.write_event(Event::Start(BytesStart::new("a:p")))
        .expect("write");
    let algn = props.alignment.as_ref().map(|a| match a {
        ParagraphAlignment::Left => "l",
        ParagraphAlignment::Center => "ctr",
        ParagraphAlignment::Right => "r",
        ParagraphAlignment::Justify => "just",
        ParagraphAlignment::Distribute => "dist",
    });
    let need_ppr = algn.is_some() || props.space_before_hundredths_pt.is_some();
    if need_ppr {
        let mut p_pr = BytesStart::new("a:pPr");
        if let Some(v) = algn {
            p_pr.push_attribute(("algn", v));
        }
        if let Some(spc) = props.space_before_hundredths_pt {
            // <a:pPr ...><a:spcBef><a:spcPts val="N"/></a:spcBef></a:pPr>
            w.write_event(Event::Start(p_pr)).expect("write pPr start");
            w.write_event(Event::Start(BytesStart::new("a:spcBef")))
                .expect("write spcBef");
            let mut spc_pts = BytesStart::new("a:spcPts");
            spc_pts.push_attribute(("val", spc.to_string().as_str()));
            w.write_event(Event::Empty(spc_pts)).expect("write spcPts");
            w.write_event(Event::End(BytesEnd::new("a:spcBef")))
                .expect("write spcBef end");
            w.write_event(Event::End(BytesEnd::new("a:pPr")))
                .expect("write pPr end");
        } else {
            w.write_event(Event::Empty(p_pr)).expect("write pPr");
        }
    }
    for run in runs {
        write_dml_run(w, run);
    }
    w.write_event(Event::End(BytesEnd::new("a:p")))
        .expect("write");
}

fn write_bullet_paragraph(w: &mut Writer<Vec<u8>>, text: &str) {
    w.write_event(Event::Start(BytesStart::new("a:p")))
        .expect("write");
    w.write_event(Event::Start(BytesStart::new("a:pPr")))
        .expect("write");
    let mut bu = BytesStart::new("a:buChar");
    bu.push_attribute(("char", "\u{2022}"));
    w.write_event(Event::Empty(bu)).expect("write");
    w.write_event(Event::End(BytesEnd::new("a:pPr")))
        .expect("write");
    w.write_event(Event::Start(BytesStart::new("a:r")))
        .expect("write");
    write_text_element(w, "a:t", text);
    w.write_event(Event::End(BytesEnd::new("a:r")))
        .expect("write");
    w.write_event(Event::End(BytesEnd::new("a:p")))
        .expect("write");
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pptx::PptxDocument;
    use std::io::Cursor;

    fn roundtrip(writer: PptxWriter) -> PptxDocument {
        let mut buf = Cursor::new(Vec::new());
        writer.write_to(&mut buf).unwrap();
        buf.set_position(0);
        PptxDocument::from_reader(buf).unwrap()
    }

    #[test]
    fn rich_runs_roundtrip() {
        let mut writer = PptxWriter::new();
        writer
            .add_slide()
            .set_title("Test")
            .add_rich_text(&[Run::new("Bold").bold(), Run::new(" red").color("FF0000")]);
        let doc = roundtrip(writer);
        let text = doc.plain_text();
        assert!(text.contains("Bold"));
        assert!(text.contains("red"));
    }

    #[test]
    fn text_box_roundtrip() {
        let mut writer = PptxWriter::new();
        writer
            .add_slide()
            .add_text_box("Floating note", 1_000_000, 5_000_000, 3_000_000, 500_000);
        let doc = roundtrip(writer);
        let text = doc.plain_text();
        assert!(text.contains("Floating note"));
    }

    #[test]
    fn set_presentation_size_written() {
        let mut writer = PptxWriter::new();
        writer.set_presentation_size(9_144_000, 6_858_000);
        writer.add_slide().add_text("test");
        let mut buf = Cursor::new(Vec::new());
        writer.write_to(&mut buf).unwrap();
        let bytes = buf.into_inner();
        let cursor = Cursor::new(bytes.clone());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        let mut entry = zip.by_name("ppt/presentation.xml").unwrap();
        let mut xml = String::new();
        std::io::Read::read_to_string(&mut entry, &mut xml).unwrap();
        assert!(xml.contains("cx=\"9144000\""), "expected cx in presentation.xml");
    }

    #[test]
    fn add_image_embeds_media_part() {
        use crate::ir::ImageFormat;
        // Minimal 1x1 PNG
        let png_bytes: Vec<u8> = vec![
            0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, // PNG signature
            0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, // IHDR length + type
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xde, // bit depth, color, crc
            0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, // IDAT
            0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xe2, 0x21,
            0xbc, 0x33, // crc
            0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82, // IEND
        ];
        let mut writer = PptxWriter::new();
        writer
            .add_slide()
            .add_image(png_bytes, ImageFormat::Png, 0, 0, 3_000_000, 2_000_000);
        let mut buf = Cursor::new(Vec::new());
        writer.write_to(&mut buf).unwrap();
        let bytes = buf.into_inner();
        let cursor = Cursor::new(bytes);
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        assert!(zip.by_name("ppt/media/image1.png").is_ok(), "media part missing");
    }

    #[test]
    fn rich_text_box_roundtrip() {
        let mut writer = PptxWriter::new();
        writer.add_slide().add_rich_text_box(
            &[
                Run::new("Big").font_size(24.0).bold(),
                Run::new(" label").italic(),
            ],
            500_000,
            500_000,
            4_000_000,
            800_000,
        );
        let doc = roundtrip(writer);
        let text = doc.plain_text();
        assert!(text.contains("Big"));
        assert!(text.contains("label"));
    }
}
