//! DOCX creation (write) module.
//!
//! Provides a builder API for creating DOCX files from scratch.
//!
//! # Example
//!
//! ```rust,no_run
//! use office_oxide::docx::write::{DocxWriter, Run, Alignment};
//!
//! let mut doc = DocxWriter::new();
//! doc.add_heading("Report", 1)
//!    .add_paragraph("This is a paragraph.")
//!    .add_rich_paragraph(&[
//!        Run::new("Bold text").bold(),
//!        Run::new(" and ").into(),
//!        Run::new("red italic").italic().color("FF0000"),
//!    ])
//!    .add_paragraph_aligned("Centred text", Alignment::Center)
//!    .add_list(&["First", "Second", "Third"], false)
//!    .add_table(&[
//!        vec!["Name", "Age"],
//!        vec!["Alice", "30"],
//!    ]);
//! doc.save("report.docx").unwrap();
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

const CT_DOCUMENT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_STYLES: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const CT_NUMBERING: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";

use crate::core::xml::ns::{R_STR as R_NS, WML_STR as WML_NS};

// ---------------------------------------------------------------------------
// Public data types
// ---------------------------------------------------------------------------

/// Paragraph text alignment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Alignment {
    Left,
    Center,
    Right,
    Justify,
}

impl Alignment {
    fn as_wml_val(&self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::Justify => "both",
        }
    }
}

/// A styled text run within a paragraph.
///
/// Build with the builder methods; plain text uses `Run::new("text")`.
///
/// # Example
/// ```rust,no_run
/// use office_oxide::docx::write::Run;
///
/// let r = Run::new("Hello, world!")
///     .bold()
///     .font_size(14.0)
///     .color("1F497D");
/// ```
#[derive(Debug, Clone, Default)]
pub struct Run {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    /// 6-char hex color string, e.g. `"FF0000"` (no leading `#`).
    pub color: Option<String>,
    /// Font size in points, e.g. `12.0`.
    pub font_size_pt: Option<f64>,
    /// Font name, e.g. `"Arial"`.
    pub font_name: Option<String>,
}

impl Run {
    /// Create a run with the given text and default (unstyled) properties.
    pub fn new(text: impl Into<String>) -> Self {
        Self { text: text.into(), ..Default::default() }
    }

    pub fn bold(mut self) -> Self { self.bold = true; self }
    pub fn italic(mut self) -> Self { self.italic = true; self }
    pub fn underline(mut self) -> Self { self.underline = true; self }
    pub fn strikethrough(mut self) -> Self { self.strikethrough = true; self }

    /// Set the font color. `hex` is a 6-character hex string without `#`.
    pub fn color(mut self, hex: impl Into<String>) -> Self {
        self.color = Some(hex.into());
        self
    }

    /// Set font size in points (e.g. `12.0`).
    pub fn font_size(mut self, pt: f64) -> Self {
        self.font_size_pt = Some(pt);
        self
    }

    /// Set the font family name (e.g. `"Arial"`).
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
    fn from(s: &str) -> Self { Self::new(s) }
}

impl From<String> for Run {
    fn from(s: String) -> Self { Self::new(s) }
}

// ---------------------------------------------------------------------------
// Internal data model
// ---------------------------------------------------------------------------

struct DocxParagraph {
    style: Option<String>,
    runs: Vec<Run>,
    numbering: Option<(u32, u8)>,
    alignment: Option<Alignment>,
}

impl DocxParagraph {
    fn plain(text: &str, style: Option<String>, numbering: Option<(u32, u8)>) -> Self {
        Self {
            style,
            runs: vec![Run::new(text)],
            numbering,
            alignment: None,
        }
    }
}

struct DocxTable {
    rows: Vec<Vec<String>>,
}

enum DocxElement {
    Paragraph(DocxParagraph),
    Table(DocxTable),
    PageBreak,
}

// ---------------------------------------------------------------------------
// Builder API
// ---------------------------------------------------------------------------

/// Builder for creating DOCX files from scratch.
pub struct DocxWriter {
    elements: Vec<DocxElement>,
}

impl DocxWriter {
    /// Create a new empty DOCX builder.
    pub fn new() -> Self {
        Self { elements: Vec::new() }
    }

    /// Add a plain paragraph with the given text.
    pub fn add_paragraph(&mut self, text: &str) -> &mut Self {
        self.elements
            .push(DocxElement::Paragraph(DocxParagraph::plain(text, None, None)));
        self
    }

    /// Add a paragraph with explicit alignment.
    pub fn add_paragraph_aligned(&mut self, text: &str, alignment: Alignment) -> &mut Self {
        self.elements.push(DocxElement::Paragraph(DocxParagraph {
            style: None,
            runs: vec![Run::new(text)],
            numbering: None,
            alignment: Some(alignment),
        }));
        self
    }

    /// Add a paragraph built from styled [`Run`]s.
    pub fn add_rich_paragraph(&mut self, runs: &[Run]) -> &mut Self {
        self.elements.push(DocxElement::Paragraph(DocxParagraph {
            style: None,
            runs: runs.to_vec(),
            numbering: None,
            alignment: None,
        }));
        self
    }

    /// Add a rich paragraph with both custom runs and explicit alignment.
    pub fn add_rich_paragraph_aligned(
        &mut self,
        runs: &[Run],
        alignment: Alignment,
    ) -> &mut Self {
        self.elements.push(DocxElement::Paragraph(DocxParagraph {
            style: None,
            runs: runs.to_vec(),
            numbering: None,
            alignment: Some(alignment),
        }));
        self
    }

    /// Add a heading paragraph at the given level (1-6).
    ///
    /// Level is clamped to `1..=6`.
    pub fn add_heading(&mut self, text: &str, level: u8) -> &mut Self {
        let level = level.clamp(1, 6);
        self.elements.push(DocxElement::Paragraph(DocxParagraph::plain(
            text,
            Some(format!("Heading{level}")),
            None,
        )));
        self
    }

    /// Add a table. The first row is treated as the header row.
    pub fn add_table(&mut self, rows: &[Vec<&str>]) -> &mut Self {
        let owned: Vec<Vec<String>> = rows
            .iter()
            .map(|row| row.iter().map(|s| s.to_string()).collect())
            .collect();
        self.elements.push(DocxElement::Table(DocxTable { rows: owned }));
        self
    }

    /// Add a list of items.
    ///
    /// `ordered = false` → bullet list; `ordered = true` → numbered list.
    pub fn add_list(&mut self, items: &[&str], ordered: bool) -> &mut Self {
        let num_id: u32 = if ordered { 2 } else { 1 };
        for item in items {
            self.elements.push(DocxElement::Paragraph(DocxParagraph::plain(
                item,
                Some("ListParagraph".to_string()),
                Some((num_id, 0)),
            )));
        }
        self
    }

    /// Insert a page break.
    pub fn add_page_break(&mut self) -> &mut Self {
        self.elements.push(DocxElement::PageBreak);
        self
    }

    /// Save the document to a file at `path`.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        let opc = OpcWriter::create(path)?;
        self.write_package(opc)?;
        Ok(())
    }

    /// Write the document to an arbitrary `Write + Seek` destination.
    pub fn write_to<W: Write + Seek>(&self, writer: W) -> Result<()> {
        let opc = OpcWriter::new(writer)?;
        self.write_package(opc)?;
        Ok(())
    }

    // -----------------------------------------------------------------------
    // Internals
    // -----------------------------------------------------------------------

    fn write_package<W: Write + Seek>(&self, mut opc: OpcWriter<W>) -> Result<()> {
        let doc_part = PartName::new("/word/document.xml")?;
        let styles_part = PartName::new("/word/styles.xml")?;

        opc.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
        opc.add_part_rel(&doc_part, rel_types::STYLES, "styles.xml");

        let document_xml = self.generate_document_xml();
        opc.add_part(&doc_part, CT_DOCUMENT, &document_xml)?;

        let styles_xml = generate_styles_xml();
        opc.add_part(&styles_part, CT_STYLES, &styles_xml)?;

        if self.has_lists() {
            let numbering_part = PartName::new("/word/numbering.xml")?;
            opc.add_part_rel(&doc_part, rel_types::NUMBERING, "numbering.xml");
            let numbering_xml = generate_numbering_xml();
            opc.add_part(&numbering_part, CT_NUMBERING, &numbering_xml)?;
        }

        opc.finish()?;
        Ok(())
    }

    fn has_lists(&self) -> bool {
        self.elements
            .iter()
            .any(|e| matches!(e, DocxElement::Paragraph(p) if p.numbering.is_some()))
    }

    fn generate_document_xml(&self) -> Vec<u8> {
        let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

        w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .expect("write decl");

        let mut root = BytesStart::new("w:document");
        root.push_attribute(("xmlns:w", WML_NS));
        root.push_attribute(("xmlns:r", R_NS));
        w.write_event(Event::Start(root)).expect("write document start");
        w.write_event(Event::Start(BytesStart::new("w:body"))).expect("write body start");

        for element in &self.elements {
            match element {
                DocxElement::Paragraph(p) => write_paragraph(&mut w, p),
                DocxElement::Table(t) => write_table(&mut w, t),
                DocxElement::PageBreak => write_page_break(&mut w),
            }
        }

        w.write_event(Event::End(BytesEnd::new("w:body"))).expect("write body end");
        w.write_event(Event::End(BytesEnd::new("w:document"))).expect("write document end");

        w.into_inner()
    }
}

impl Default for DocxWriter {
    fn default() -> Self { Self::new() }
}

// ---------------------------------------------------------------------------
// XML serialisation helpers
// ---------------------------------------------------------------------------

fn write_paragraph(w: &mut Writer<Vec<u8>>, p: &DocxParagraph) {
    w.write_event(Event::Start(BytesStart::new("w:p"))).expect("write p start");

    let has_ppr = p.style.is_some() || p.numbering.is_some() || p.alignment.is_some();
    if has_ppr {
        w.write_event(Event::Start(BytesStart::new("w:pPr"))).expect("write pPr start");

        if let Some(ref style) = p.style {
            let mut elem = BytesStart::new("w:pStyle");
            elem.push_attribute(("w:val", style.as_str()));
            w.write_event(Event::Empty(elem)).expect("write pStyle");
        }

        if let Some(ref align) = p.alignment {
            let mut elem = BytesStart::new("w:jc");
            elem.push_attribute(("w:val", align.as_wml_val()));
            w.write_event(Event::Empty(elem)).expect("write jc");
        }

        if let Some((num_id, ilvl)) = p.numbering {
            w.write_event(Event::Start(BytesStart::new("w:numPr"))).expect("write numPr start");

            let mut ilvl_elem = BytesStart::new("w:ilvl");
            ilvl_elem.push_attribute(("w:val", ilvl.to_string().as_str()));
            w.write_event(Event::Empty(ilvl_elem)).expect("write ilvl");

            let mut num_id_elem = BytesStart::new("w:numId");
            num_id_elem.push_attribute(("w:val", num_id.to_string().as_str()));
            w.write_event(Event::Empty(num_id_elem)).expect("write numId");

            w.write_event(Event::End(BytesEnd::new("w:numPr"))).expect("write numPr end");
        }

        w.write_event(Event::End(BytesEnd::new("w:pPr"))).expect("write pPr end");
    }

    for run in &p.runs {
        write_run(w, run);
    }

    w.write_event(Event::End(BytesEnd::new("w:p"))).expect("write p end");
}

fn write_run(w: &mut Writer<Vec<u8>>, run: &Run) {
    w.write_event(Event::Start(BytesStart::new("w:r"))).expect("write r start");

    if run.has_rpr() {
        w.write_event(Event::Start(BytesStart::new("w:rPr"))).expect("write rPr start");

        if let Some(ref name) = run.font_name {
            let mut elem = BytesStart::new("w:rFonts");
            elem.push_attribute(("w:ascii", name.as_str()));
            elem.push_attribute(("w:hAnsi", name.as_str()));
            w.write_event(Event::Empty(elem)).expect("write rFonts");
        }

        if run.bold {
            w.write_event(Event::Empty(BytesStart::new("w:b"))).expect("write bold");
        }
        if run.italic {
            w.write_event(Event::Empty(BytesStart::new("w:i"))).expect("write italic");
        }
        if run.underline {
            let mut elem = BytesStart::new("w:u");
            elem.push_attribute(("w:val", "single"));
            w.write_event(Event::Empty(elem)).expect("write underline");
        }
        if run.strikethrough {
            w.write_event(Event::Empty(BytesStart::new("w:strike"))).expect("write strike");
        }
        if let Some(ref hex) = run.color {
            let mut elem = BytesStart::new("w:color");
            elem.push_attribute(("w:val", hex.as_str()));
            w.write_event(Event::Empty(elem)).expect("write color");
        }
        if let Some(pt) = run.font_size_pt {
            // WML stores size in half-points
            let half_pts = (pt * 2.0).round() as u32;
            let val = half_pts.to_string();
            let mut sz = BytesStart::new("w:sz");
            sz.push_attribute(("w:val", val.as_str()));
            w.write_event(Event::Empty(sz)).expect("write sz");
            let mut sz_cs = BytesStart::new("w:szCs");
            sz_cs.push_attribute(("w:val", val.as_str()));
            w.write_event(Event::Empty(sz_cs)).expect("write szCs");
        }

        w.write_event(Event::End(BytesEnd::new("w:rPr"))).expect("write rPr end");
    }

    let text = &run.text;
    let mut t_elem = BytesStart::new("w:t");
    if text.starts_with(' ') || text.ends_with(' ') || text.contains("  ") {
        t_elem.push_attribute(("xml:space", "preserve"));
    }
    w.write_event(Event::Start(t_elem)).expect("write t start");
    w.write_event(Event::Text(BytesText::new(text))).expect("write text");
    w.write_event(Event::End(BytesEnd::new("w:t"))).expect("write t end");

    w.write_event(Event::End(BytesEnd::new("w:r"))).expect("write r end");
}

fn write_page_break(w: &mut Writer<Vec<u8>>) {
    w.write_event(Event::Start(BytesStart::new("w:p"))).expect("write p start");
    w.write_event(Event::Start(BytesStart::new("w:r"))).expect("write r start");
    let mut br = BytesStart::new("w:br");
    br.push_attribute(("w:type", "page"));
    w.write_event(Event::Empty(br)).expect("write br");
    w.write_event(Event::End(BytesEnd::new("w:r"))).expect("write r end");
    w.write_event(Event::End(BytesEnd::new("w:p"))).expect("write p end");
}

fn write_table(w: &mut Writer<Vec<u8>>, table: &DocxTable) {
    w.write_event(Event::Start(BytesStart::new("w:tbl"))).expect("write tbl start");

    for row in &table.rows {
        w.write_event(Event::Start(BytesStart::new("w:tr"))).expect("write tr start");

        for cell_text in row {
            w.write_event(Event::Start(BytesStart::new("w:tc"))).expect("write tc start");
            let p = DocxParagraph::plain(cell_text, None, None);
            write_paragraph(w, &p);
            w.write_event(Event::End(BytesEnd::new("w:tc"))).expect("write tc end");
        }

        w.write_event(Event::End(BytesEnd::new("w:tr"))).expect("write tr end");
    }

    w.write_event(Event::End(BytesEnd::new("w:tbl"))).expect("write tbl end");
}

// ---------------------------------------------------------------------------
// Styles and numbering generators
// ---------------------------------------------------------------------------

fn generate_styles_xml() -> Vec<u8> {
    let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

    w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .expect("write decl");

    let mut root = BytesStart::new("w:styles");
    root.push_attribute(("xmlns:w", WML_NS));
    w.write_event(Event::Start(root)).expect("write styles start");

    write_paragraph_style(&mut w, "Normal", "Normal", None);
    for level in 1u8..=6 {
        let style_id = format!("Heading{level}");
        let name = format!("heading {level}");
        write_paragraph_style(&mut w, &style_id, &name, Some(level - 1));
    }
    write_paragraph_style(&mut w, "ListParagraph", "List Paragraph", None);

    w.write_event(Event::End(BytesEnd::new("w:styles"))).expect("write styles end");

    w.into_inner()
}

fn write_paragraph_style(
    w: &mut Writer<Vec<u8>>,
    style_id: &str,
    name: &str,
    outline_level: Option<u8>,
) {
    let mut elem = BytesStart::new("w:style");
    elem.push_attribute(("w:type", "paragraph"));
    elem.push_attribute(("w:styleId", style_id));
    w.write_event(Event::Start(elem)).expect("write style start");

    let mut name_elem = BytesStart::new("w:name");
    name_elem.push_attribute(("w:val", name));
    w.write_event(Event::Empty(name_elem)).expect("write style name");

    if let Some(level) = outline_level {
        w.write_event(Event::Start(BytesStart::new("w:pPr"))).expect("write pPr start");
        let mut lvl = BytesStart::new("w:outlineLvl");
        lvl.push_attribute(("w:val", level.to_string().as_str()));
        w.write_event(Event::Empty(lvl)).expect("write outlineLvl");
        w.write_event(Event::End(BytesEnd::new("w:pPr"))).expect("write pPr end");
    }

    w.write_event(Event::End(BytesEnd::new("w:style"))).expect("write style end");
}

fn generate_numbering_xml() -> Vec<u8> {
    let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

    w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .expect("write decl");

    let mut root = BytesStart::new("w:numbering");
    root.push_attribute(("xmlns:w", WML_NS));
    w.write_event(Event::Start(root)).expect("write numbering start");

    write_abstract_num(&mut w, 0, "bullet", "\u{2022}");
    write_abstract_num(&mut w, 1, "decimal", "%1.");
    write_num(&mut w, 1, 0);
    write_num(&mut w, 2, 1);

    w.write_event(Event::End(BytesEnd::new("w:numbering"))).expect("write numbering end");

    w.into_inner()
}

fn write_abstract_num(
    w: &mut Writer<Vec<u8>>,
    abstract_num_id: u32,
    num_fmt: &str,
    lvl_text: &str,
) {
    let mut elem = BytesStart::new("w:abstractNum");
    elem.push_attribute(("w:abstractNumId", abstract_num_id.to_string().as_str()));
    w.write_event(Event::Start(elem)).expect("write abstractNum start");

    let mut lvl = BytesStart::new("w:lvl");
    lvl.push_attribute(("w:ilvl", "0"));
    w.write_event(Event::Start(lvl)).expect("write lvl start");

    let mut fmt = BytesStart::new("w:numFmt");
    fmt.push_attribute(("w:val", num_fmt));
    w.write_event(Event::Empty(fmt)).expect("write numFmt");

    let mut text = BytesStart::new("w:lvlText");
    text.push_attribute(("w:val", lvl_text));
    w.write_event(Event::Empty(text)).expect("write lvlText");

    w.write_event(Event::End(BytesEnd::new("w:lvl"))).expect("write lvl end");
    w.write_event(Event::End(BytesEnd::new("w:abstractNum"))).expect("write abstractNum end");
}

fn write_num(w: &mut Writer<Vec<u8>>, num_id: u32, abstract_num_id: u32) {
    let mut elem = BytesStart::new("w:num");
    elem.push_attribute(("w:numId", num_id.to_string().as_str()));
    w.write_event(Event::Start(elem)).expect("write num start");

    let mut abs = BytesStart::new("w:abstractNumId");
    abs.push_attribute(("w:val", abstract_num_id.to_string().as_str()));
    w.write_event(Event::Empty(abs)).expect("write abstractNumId");

    w.write_event(Event::End(BytesEnd::new("w:num"))).expect("write num end");
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;
    use crate::docx::DocxDocument;

    fn roundtrip(doc: DocxWriter) -> DocxDocument {
        let mut buf = Cursor::new(Vec::new());
        doc.write_to(&mut buf).unwrap();
        buf.set_position(0);
        DocxDocument::from_reader(buf).unwrap()
    }

    #[test]
    fn rich_run_bold_italic() {
        let mut doc = DocxWriter::new();
        doc.add_rich_paragraph(&[
            Run::new("Hello ").bold(),
            Run::new("world").italic().color("FF0000"),
        ]);
        let parsed = roundtrip(doc);
        let text = parsed.plain_text();
        assert!(text.contains("Hello"));
        assert!(text.contains("world"));
    }

    #[test]
    fn alignment_center() {
        let mut doc = DocxWriter::new();
        doc.add_paragraph_aligned("Centred", Alignment::Center);
        let parsed = roundtrip(doc);
        assert!(parsed.plain_text().contains("Centred"));
    }

    #[test]
    fn page_break_roundtrip() {
        let mut doc = DocxWriter::new();
        doc.add_paragraph("Before");
        doc.add_page_break();
        doc.add_paragraph("After");
        let parsed = roundtrip(doc);
        let text = parsed.plain_text();
        assert!(text.contains("Before"));
        assert!(text.contains("After"));
    }

    #[test]
    fn font_size_and_name() {
        let mut doc = DocxWriter::new();
        doc.add_rich_paragraph(&[
            Run::new("Big text").font_size(24.0).font("Arial"),
        ]);
        let parsed = roundtrip(doc);
        assert!(parsed.plain_text().contains("Big text"));
    }

    #[test]
    fn underline_strikethrough() {
        let mut doc = DocxWriter::new();
        doc.add_rich_paragraph(&[
            Run::new("under").underline(),
            Run::new(" strike").strikethrough(),
        ]);
        let parsed = roundtrip(doc);
        let text = parsed.plain_text();
        assert!(text.contains("under"));
        assert!(text.contains("strike"));
    }
}
