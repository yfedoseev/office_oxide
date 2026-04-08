//! PPTX creation (write) module.
//!
//! Provides a builder API for creating PPTX files from scratch.
//!
//! # Example
//!
//! ```rust,no_run
//! use office_oxide::pptx::write::PptxWriter;
//!
//! let mut writer = PptxWriter::new();
//! writer.add_slide()
//!     .set_title("Hello")
//!     .add_text("World")
//!     .add_bullet_list(&["First", "Second", "Third"]);
//! writer.save("output.pptx").unwrap();
//! ```

use std::io::{Seek, Write};
use std::path::Path;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;

use crate::core::opc::{OpcWriter, PartName};
use crate::core::relationships::rel_types;

use super::Result;

// ---------------------------------------------------------------------------
// Content types
// ---------------------------------------------------------------------------

const CT_PRESENTATION: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const CT_SLIDE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const CT_SLIDE_LAYOUT: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
const CT_SLIDE_MASTER: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";

// ---------------------------------------------------------------------------
// Namespaces
// ---------------------------------------------------------------------------

const NS_PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

// ---------------------------------------------------------------------------
// Slide size (standard 16:9 in EMU)
// ---------------------------------------------------------------------------

const SLIDE_WIDTH: &str = "12192000";
const SLIDE_HEIGHT: &str = "6858000";

// ---------------------------------------------------------------------------
// Body content items
// ---------------------------------------------------------------------------

/// A single content item in the body of a slide.
#[derive(Debug, Clone)]
enum BodyItem {
    /// A plain text paragraph.
    Text(String),
    /// A list of bullet items, each rendered as a separate `<a:p>`.
    BulletList(Vec<String>),
}

// ---------------------------------------------------------------------------
// SlideData
// ---------------------------------------------------------------------------

/// Data for a single slide being constructed.
#[derive(Debug, Clone)]
pub struct SlideData {
    /// The slide title (if set).
    pub title: Option<String>,
    body_items: Vec<BodyItem>,
}

impl SlideData {
    fn new() -> Self {
        Self {
            title: None,
            body_items: Vec::new(),
        }
    }

    /// Set the slide title. Overwrites any previously set title.
    pub fn set_title(&mut self, title: &str) -> &mut Self {
        self.title = Some(title.to_string());
        self
    }

    /// Add a plain text paragraph to the body area.
    pub fn add_text(&mut self, text: &str) -> &mut Self {
        self.body_items.push(BodyItem::Text(text.to_string()));
        self
    }

    /// Add a bullet list to the body area. Each item becomes a separate paragraph.
    pub fn add_bullet_list(&mut self, items: &[&str]) -> &mut Self {
        let owned: Vec<String> = items.iter().map(|s| s.to_string()).collect();
        self.body_items.push(BodyItem::BulletList(owned));
        self
    }

    /// Returns `true` if the slide has any body content.
    fn has_body(&self) -> bool {
        !self.body_items.is_empty()
    }
}

// ---------------------------------------------------------------------------
// PptxWriter
// ---------------------------------------------------------------------------

/// Builder for creating PPTX files from scratch.
pub struct PptxWriter {
    slides: Vec<SlideData>,
}

impl PptxWriter {
    /// Create a new empty PPTX writer.
    pub fn new() -> Self {
        Self { slides: Vec::new() }
    }

    /// Add a new slide and return a mutable reference for configuration.
    pub fn add_slide(&mut self) -> &mut SlideData {
        self.slides.push(SlideData::new());
        self.slides.last_mut().expect("just pushed")
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

    /// Internal: write all parts through the OPC writer.
    fn write_opc<W: Write + Seek>(&self, mut opc: OpcWriter<W>) -> Result<()> {
        let pres_part = PartName::new("/ppt/presentation.xml")?;
        let master_part = PartName::new("/ppt/slideMasters/slideMaster1.xml")?;
        let layout_part = PartName::new("/ppt/slideLayouts/slideLayout1.xml")?;

        // --- Package relationship: presentation ---
        opc.add_package_rel(rel_types::OFFICE_DOCUMENT, "ppt/presentation.xml");

        // --- Presentation relationships ---
        // rId1 = slide master
        opc.add_part_rel(
            &pres_part,
            rel_types::SLIDE_MASTER,
            "slideMasters/slideMaster1.xml",
        );

        // rId2..rId(n+1) = slides
        let mut slide_parts = Vec::with_capacity(self.slides.len());
        for i in 0..self.slides.len() {
            let idx = i + 1;
            let slide_part = PartName::new(&format!("/ppt/slides/slide{idx}.xml"))?;
            opc.add_part_rel(
                &pres_part,
                rel_types::SLIDE,
                &format!("slides/slide{idx}.xml"),
            );
            slide_parts.push(slide_part);
        }

        // --- Slide master relationship: rId1 = slide layout ---
        opc.add_part_rel(
            &master_part,
            rel_types::SLIDE_LAYOUT,
            "../slideLayouts/slideLayout1.xml",
        );

        // --- Each slide relationship: rId1 = slide layout ---
        for slide_part in &slide_parts {
            opc.add_part_rel(
                slide_part,
                rel_types::SLIDE_LAYOUT,
                "../slideLayouts/slideLayout1.xml",
            );
        }

        // --- Generate and add parts ---
        let pres_xml = generate_presentation_xml(self.slides.len());
        opc.add_part(&pres_part, CT_PRESENTATION, &pres_xml)?;

        let master_xml = generate_slide_master_xml();
        opc.add_part(&master_part, CT_SLIDE_MASTER, &master_xml)?;

        let layout_xml = generate_slide_layout_xml();
        opc.add_part(&layout_part, CT_SLIDE_LAYOUT, &layout_xml)?;

        for (i, slide) in self.slides.iter().enumerate() {
            let slide_xml = generate_slide_xml(slide);
            opc.add_part(&slide_parts[i], CT_SLIDE, &slide_xml)?;
        }

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

/// Write the XML declaration.
fn write_decl(w: &mut Writer<Vec<u8>>) {
    w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .expect("write decl");
}

/// Write `<tag>text</tag>`.
fn write_text_element(w: &mut Writer<Vec<u8>>, tag: &str, text: &str) {
    w.write_event(Event::Start(BytesStart::new(tag)))
        .expect("write start");
    w.write_event(Event::Text(BytesText::new(text)))
        .expect("write text");
    w.write_event(Event::End(BytesEnd::new(tag)))
        .expect("write end");
}

/// Write `<tag/>` (empty element).
fn write_empty(w: &mut Writer<Vec<u8>>, tag: &str) {
    w.write_event(Event::Empty(BytesStart::new(tag)))
        .expect("write empty");
}

/// Create a `BytesStart` element with PML/DML/REL namespace attributes.
fn pml_root(tag: &str) -> BytesStart<'_> {
    let mut elem = BytesStart::new(tag);
    elem.push_attribute(("xmlns:p", NS_PML));
    elem.push_attribute(("xmlns:a", NS_DML));
    elem.push_attribute(("xmlns:r", NS_REL));
    elem
}

/// Write the standard `<p:nvGrpSpPr>` block for a shape tree root.
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

// ---------------------------------------------------------------------------
// presentation.xml
// ---------------------------------------------------------------------------

fn generate_presentation_xml(slide_count: usize) -> Vec<u8> {
    let mut w = Writer::new(Vec::new());
    write_decl(&mut w);

    w.write_event(Event::Start(pml_root("p:presentation")))
        .expect("write");

    // <p:sldMasterIdLst>
    w.write_event(Event::Start(BytesStart::new("p:sldMasterIdLst")))
        .expect("write");
    let mut master_id = BytesStart::new("p:sldMasterId");
    master_id.push_attribute(("id", "2147483648"));
    master_id.push_attribute(("r:id", "rId1"));
    w.write_event(Event::Empty(master_id)).expect("write");
    w.write_event(Event::End(BytesEnd::new("p:sldMasterIdLst")))
        .expect("write");

    // <p:sldIdLst>
    w.write_event(Event::Start(BytesStart::new("p:sldIdLst")))
        .expect("write");
    for i in 0..slide_count {
        let slide_id_val = 256 + i as u32;
        let r_id = format!("rId{}", i + 2); // rId2, rId3, ...
        let mut slide_id = BytesStart::new("p:sldId");
        slide_id.push_attribute(("id", slide_id_val.to_string().as_str()));
        slide_id.push_attribute(("r:id", r_id.as_str()));
        w.write_event(Event::Empty(slide_id)).expect("write");
    }
    w.write_event(Event::End(BytesEnd::new("p:sldIdLst")))
        .expect("write");

    // <p:sldSz>
    let mut sld_sz = BytesStart::new("p:sldSz");
    sld_sz.push_attribute(("cx", SLIDE_WIDTH));
    sld_sz.push_attribute(("cy", SLIDE_HEIGHT));
    w.write_event(Event::Empty(sld_sz)).expect("write");

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

    // <p:cSld><p:spTree>
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

    // <p:sldLayoutIdLst>
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

    let mut root = pml_root("p:sldLayout");
    root.push_attribute(("type", "blank"));
    w.write_event(Event::Start(root)).expect("write");

    // <p:cSld><p:spTree>
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

    w.write_event(Event::End(BytesEnd::new("p:sldLayout")))
        .expect("write");

    w.into_inner()
}

// ---------------------------------------------------------------------------
// slides/slideN.xml
// ---------------------------------------------------------------------------

fn generate_slide_xml(slide: &SlideData) -> Vec<u8> {
    let mut w = Writer::new(Vec::new());
    write_decl(&mut w);

    w.write_event(Event::Start(pml_root("p:sld")))
        .expect("write");

    // <p:cSld><p:spTree>
    w.write_event(Event::Start(BytesStart::new("p:cSld")))
        .expect("write");
    w.write_event(Event::Start(BytesStart::new("p:spTree")))
        .expect("write");

    write_nv_grp_sp_pr(&mut w);
    write_empty(&mut w, "p:grpSpPr");

    // Shape IDs: 1 is reserved for the group, 2+ for shapes.
    let mut next_id: u32 = 2;

    // Title shape (only if title is set)
    if let Some(ref title) = slide.title {
        write_title_shape(&mut w, next_id, title);
        next_id += 1;
    }

    // Body shape (only if there is body content)
    if slide.has_body() {
        write_body_shape(&mut w, next_id, &slide.body_items);
    }

    // Close spTree, cSld, sld
    w.write_event(Event::End(BytesEnd::new("p:spTree")))
        .expect("write");
    w.write_event(Event::End(BytesEnd::new("p:cSld")))
        .expect("write");
    w.write_event(Event::End(BytesEnd::new("p:sld")))
        .expect("write");

    w.into_inner()
}

/// Write a title placeholder shape.
fn write_title_shape(w: &mut Writer<Vec<u8>>, id: u32, title: &str) {
    let id_str = id.to_string();

    // <p:sp>
    w.write_event(Event::Start(BytesStart::new("p:sp")))
        .expect("write");

    // <p:nvSpPr>
    w.write_event(Event::Start(BytesStart::new("p:nvSpPr")))
        .expect("write");

    let mut cnv_pr = BytesStart::new("p:cNvPr");
    cnv_pr.push_attribute(("id", id_str.as_str()));
    cnv_pr.push_attribute(("name", "Title 1"));
    w.write_event(Event::Empty(cnv_pr)).expect("write");

    // <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    w.write_event(Event::Start(BytesStart::new("p:cNvSpPr")))
        .expect("write");
    let mut locks = BytesStart::new("a:spLocks");
    locks.push_attribute(("noGrp", "1"));
    w.write_event(Event::Empty(locks)).expect("write");
    w.write_event(Event::End(BytesEnd::new("p:cNvSpPr")))
        .expect("write");

    // <p:nvPr><p:ph type="title"/></p:nvPr>
    w.write_event(Event::Start(BytesStart::new("p:nvPr")))
        .expect("write");
    let mut ph = BytesStart::new("p:ph");
    ph.push_attribute(("type", "title"));
    w.write_event(Event::Empty(ph)).expect("write");
    w.write_event(Event::End(BytesEnd::new("p:nvPr")))
        .expect("write");

    w.write_event(Event::End(BytesEnd::new("p:nvSpPr")))
        .expect("write");

    // <p:spPr/>
    write_empty(w, "p:spPr");

    // <p:txBody>
    w.write_event(Event::Start(BytesStart::new("p:txBody")))
        .expect("write");
    write_empty(w, "a:bodyPr");
    write_text_paragraph(w, title);
    w.write_event(Event::End(BytesEnd::new("p:txBody")))
        .expect("write");

    // </p:sp>
    w.write_event(Event::End(BytesEnd::new("p:sp")))
        .expect("write");
}

/// Write a body placeholder shape containing all body items.
fn write_body_shape(w: &mut Writer<Vec<u8>>, id: u32, items: &[BodyItem]) {
    let id_str = id.to_string();

    // <p:sp>
    w.write_event(Event::Start(BytesStart::new("p:sp")))
        .expect("write");

    // <p:nvSpPr>
    w.write_event(Event::Start(BytesStart::new("p:nvSpPr")))
        .expect("write");

    let mut cnv_pr = BytesStart::new("p:cNvPr");
    cnv_pr.push_attribute(("id", id_str.as_str()));
    cnv_pr.push_attribute(("name", "Body 2"));
    w.write_event(Event::Empty(cnv_pr)).expect("write");

    // <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    w.write_event(Event::Start(BytesStart::new("p:cNvSpPr")))
        .expect("write");
    let mut locks = BytesStart::new("a:spLocks");
    locks.push_attribute(("noGrp", "1"));
    w.write_event(Event::Empty(locks)).expect("write");
    w.write_event(Event::End(BytesEnd::new("p:cNvSpPr")))
        .expect("write");

    // <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
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

    // <p:spPr/>
    write_empty(w, "p:spPr");

    // <p:txBody>
    w.write_event(Event::Start(BytesStart::new("p:txBody")))
        .expect("write");
    write_empty(w, "a:bodyPr");

    for item in items {
        match item {
            BodyItem::Text(text) => {
                write_text_paragraph(w, text);
            }
            BodyItem::BulletList(bullets) => {
                for bullet in bullets {
                    write_bullet_paragraph(w, bullet);
                }
            }
        }
    }

    w.write_event(Event::End(BytesEnd::new("p:txBody")))
        .expect("write");

    // </p:sp>
    w.write_event(Event::End(BytesEnd::new("p:sp")))
        .expect("write");
}

/// Write a single `<a:p><a:r><a:t>text</a:t></a:r></a:p>`.
fn write_text_paragraph(w: &mut Writer<Vec<u8>>, text: &str) {
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

/// Write a bullet paragraph: `<a:p><a:pPr><a:buChar char="&#x2022;"/></a:pPr><a:r><a:t>text</a:t></a:r></a:p>`.
fn write_bullet_paragraph(w: &mut Writer<Vec<u8>>, text: &str) {
    w.write_event(Event::Start(BytesStart::new("a:p")))
        .expect("write");

    // <a:pPr><a:buChar char="&#x2022;"/></a:pPr>
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
