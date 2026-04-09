//! DOCX creation (write) module.
//!
//! Provides a builder API for creating DOCX files from scratch.
//!
//! # Example
//!
//! ```rust,no_run
//! use office_oxide::docx::write::DocxWriter;
//!
//! let mut doc = DocxWriter::new();
//! doc.add_heading("Report", 1)
//!    .add_paragraph("This is a paragraph.")
//!    .add_list(&["First", "Second", "Third"], false)
//!    .add_table(&[
//!        vec!["Name", "Age"],
//!        vec!["Alice", "30"],
//!    ]);
//! doc.save("report.docx").unwrap();
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

const CT_DOCUMENT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_STYLES: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const CT_NUMBERING: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";

use crate::core::xml::ns::{WML_STR as WML_NS, R_STR as R_NS};

// ---------------------------------------------------------------------------
// Data model
// ---------------------------------------------------------------------------

/// Builder for creating DOCX files from scratch.
pub struct DocxWriter {
    elements: Vec<DocxElement>,
}

enum DocxElement {
    Paragraph(DocxParagraph),
    Table(DocxTable),
}

struct DocxParagraph {
    style: Option<String>,
    items: Vec<DocxRun>,
    numbering: Option<(u32, u8)>, // (numId, ilvl)
}

struct DocxRun {
    text: String,
    bold: bool,
    italic: bool,
}

struct DocxTable {
    rows: Vec<Vec<String>>,
}

// ---------------------------------------------------------------------------
// Builder API
// ---------------------------------------------------------------------------

impl DocxWriter {
    /// Create a new empty DOCX builder.
    pub fn new() -> Self {
        Self {
            elements: Vec::new(),
        }
    }

    /// Add a plain paragraph with the given text.
    pub fn add_paragraph(&mut self, text: &str) -> &mut Self {
        self.elements.push(DocxElement::Paragraph(DocxParagraph {
            style: None,
            items: vec![DocxRun {
                text: text.to_string(),
                bold: false,
                italic: false,
            }],
            numbering: None,
        }));
        self
    }

    /// Add a heading paragraph at the given level (1-6).
    ///
    /// The level is clamped to the range 1..=6.
    pub fn add_heading(&mut self, text: &str, level: u8) -> &mut Self {
        let level = level.clamp(1, 6);
        let style = format!("Heading{level}");
        self.elements.push(DocxElement::Paragraph(DocxParagraph {
            style: Some(style),
            items: vec![DocxRun {
                text: text.to_string(),
                bold: false,
                italic: false,
            }],
            numbering: None,
        }));
        self
    }

    /// Add a table. The first row is treated as the header row.
    ///
    /// Each inner `Vec<&str>` is one row of cells.
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
    /// When `ordered` is false a bullet list is produced (numId=1);
    /// when `ordered` is true a decimal-numbered list is produced (numId=2).
    pub fn add_list(&mut self, items: &[&str], ordered: bool) -> &mut Self {
        let num_id: u32 = if ordered { 2 } else { 1 };
        for item in items {
            self.elements.push(DocxElement::Paragraph(DocxParagraph {
                style: Some("ListParagraph".to_string()),
                items: vec![DocxRun {
                    text: item.to_string(),
                    bold: false,
                    italic: false,
                }],
                numbering: Some((num_id, 0)),
            }));
        }
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

    /// Write the complete OPC package.
    fn write_package<W: Write + Seek>(&self, mut opc: OpcWriter<W>) -> Result<()> {
        let doc_part = PartName::new("/word/document.xml")?;
        let styles_part = PartName::new("/word/styles.xml")?;

        // Package relationship: officeDocument -> word/document.xml
        opc.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");

        // Part relationship: document.xml -> styles.xml
        opc.add_part_rel(&doc_part, rel_types::STYLES, "styles.xml");

        // Generate and add document.xml
        let document_xml = self.generate_document_xml();
        opc.add_part(&doc_part, CT_DOCUMENT, &document_xml)?;

        // Generate and add styles.xml
        let styles_xml = generate_styles_xml();
        opc.add_part(&styles_part, CT_STYLES, &styles_xml)?;

        // If any lists are present, add numbering.xml
        if self.has_lists() {
            let numbering_part = PartName::new("/word/numbering.xml")?;
            opc.add_part_rel(&doc_part, rel_types::NUMBERING, "numbering.xml");
            let numbering_xml = generate_numbering_xml();
            opc.add_part(&numbering_part, CT_NUMBERING, &numbering_xml)?;
        }

        opc.finish()?;
        Ok(())
    }

    /// Check whether any element uses numbering (lists).
    fn has_lists(&self) -> bool {
        self.elements.iter().any(|e| matches!(e, DocxElement::Paragraph(p) if p.numbering.is_some()))
    }

    /// Generate the `word/document.xml` content.
    fn generate_document_xml(&self) -> Vec<u8> {
        let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

        // XML declaration
        w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .expect("write decl");

        // <w:document>
        let mut root = BytesStart::new("w:document");
        root.push_attribute(("xmlns:w", WML_NS));
        root.push_attribute(("xmlns:r", R_NS));
        w.write_event(Event::Start(root)).expect("write document start");

        // <w:body>
        w.write_event(Event::Start(BytesStart::new("w:body")))
            .expect("write body start");

        for element in &self.elements {
            match element {
                DocxElement::Paragraph(p) => write_paragraph(&mut w, p),
                DocxElement::Table(t) => write_table(&mut w, t),
            }
        }

        // </w:body>
        w.write_event(Event::End(BytesEnd::new("w:body")))
            .expect("write body end");

        // </w:document>
        w.write_event(Event::End(BytesEnd::new("w:document")))
            .expect("write document end");

        w.into_inner()
    }
}

impl Default for DocxWriter {
    fn default() -> Self {
        Self::new()
    }
}

// ---------------------------------------------------------------------------
// XML serialisation helpers
// ---------------------------------------------------------------------------

/// Write a single `<w:p>` element.
fn write_paragraph(w: &mut Writer<Vec<u8>>, p: &DocxParagraph) {
    w.write_event(Event::Start(BytesStart::new("w:p")))
        .expect("write p start");

    // Paragraph properties
    if p.style.is_some() || p.numbering.is_some() {
        w.write_event(Event::Start(BytesStart::new("w:pPr")))
            .expect("write pPr start");

        if let Some(ref style) = p.style {
            let mut elem = BytesStart::new("w:pStyle");
            elem.push_attribute(("w:val", style.as_str()));
            w.write_event(Event::Empty(elem)).expect("write pStyle");
        }

        if let Some((num_id, ilvl)) = p.numbering {
            w.write_event(Event::Start(BytesStart::new("w:numPr")))
                .expect("write numPr start");

            let mut ilvl_elem = BytesStart::new("w:ilvl");
            ilvl_elem.push_attribute(("w:val", ilvl.to_string().as_str()));
            w.write_event(Event::Empty(ilvl_elem)).expect("write ilvl");

            let mut num_id_elem = BytesStart::new("w:numId");
            num_id_elem.push_attribute(("w:val", num_id.to_string().as_str()));
            w.write_event(Event::Empty(num_id_elem)).expect("write numId");

            w.write_event(Event::End(BytesEnd::new("w:numPr")))
                .expect("write numPr end");
        }

        w.write_event(Event::End(BytesEnd::new("w:pPr")))
            .expect("write pPr end");
    }

    // Runs
    for run in &p.items {
        write_run(w, run);
    }

    w.write_event(Event::End(BytesEnd::new("w:p")))
        .expect("write p end");
}

/// Write a single `<w:r>` element.
fn write_run(w: &mut Writer<Vec<u8>>, run: &DocxRun) {
    w.write_event(Event::Start(BytesStart::new("w:r")))
        .expect("write r start");

    // Run properties (bold / italic)
    if run.bold || run.italic {
        w.write_event(Event::Start(BytesStart::new("w:rPr")))
            .expect("write rPr start");

        if run.bold {
            w.write_event(Event::Empty(BytesStart::new("w:b")))
                .expect("write bold");
        }
        if run.italic {
            w.write_event(Event::Empty(BytesStart::new("w:i")))
                .expect("write italic");
        }

        w.write_event(Event::End(BytesEnd::new("w:rPr")))
            .expect("write rPr end");
    }

    // Text element — use xml:space="preserve" to keep leading/trailing whitespace
    let mut t_elem = BytesStart::new("w:t");
    let text = &run.text;
    if text.starts_with(' ') || text.ends_with(' ') || text.contains("  ") {
        t_elem.push_attribute(("xml:space", "preserve"));
    }
    w.write_event(Event::Start(t_elem)).expect("write t start");
    w.write_event(Event::Text(BytesText::new(text)))
        .expect("write text");
    w.write_event(Event::End(BytesEnd::new("w:t")))
        .expect("write t end");

    w.write_event(Event::End(BytesEnd::new("w:r")))
        .expect("write r end");
}

/// Write a `<w:tbl>` element.
fn write_table(w: &mut Writer<Vec<u8>>, table: &DocxTable) {
    w.write_event(Event::Start(BytesStart::new("w:tbl")))
        .expect("write tbl start");

    for row in &table.rows {
        w.write_event(Event::Start(BytesStart::new("w:tr")))
            .expect("write tr start");

        for cell_text in row {
            w.write_event(Event::Start(BytesStart::new("w:tc")))
                .expect("write tc start");

            // Each cell must contain at least one paragraph
            let p = DocxParagraph {
                style: None,
                items: vec![DocxRun {
                    text: cell_text.clone(),
                    bold: false,
                    italic: false,
                }],
                numbering: None,
            };
            write_paragraph(w, &p);

            w.write_event(Event::End(BytesEnd::new("w:tc")))
                .expect("write tc end");
        }

        w.write_event(Event::End(BytesEnd::new("w:tr")))
            .expect("write tr end");
    }

    w.write_event(Event::End(BytesEnd::new("w:tbl")))
        .expect("write tbl end");
}

// ---------------------------------------------------------------------------
// Styles and numbering generators
// ---------------------------------------------------------------------------

/// Generate a minimal `word/styles.xml` with Normal, Heading1-6, and ListParagraph styles.
fn generate_styles_xml() -> Vec<u8> {
    let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

    w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .expect("write decl");

    let mut root = BytesStart::new("w:styles");
    root.push_attribute(("xmlns:w", WML_NS));
    w.write_event(Event::Start(root)).expect("write styles start");

    // Normal style
    write_paragraph_style(&mut w, "Normal", "Normal", None);

    // Heading styles (1–6)
    for level in 1u8..=6 {
        let style_id = format!("Heading{level}");
        let name = format!("heading {level}");
        write_paragraph_style(&mut w, &style_id, &name, Some(level - 1));
    }

    // ListParagraph style
    write_paragraph_style(&mut w, "ListParagraph", "List Paragraph", None);

    w.write_event(Event::End(BytesEnd::new("w:styles")))
        .expect("write styles end");

    w.into_inner()
}

/// Write a single `<w:style>` element.
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

    // <w:name w:val="..."/>
    let mut name_elem = BytesStart::new("w:name");
    name_elem.push_attribute(("w:val", name));
    w.write_event(Event::Empty(name_elem)).expect("write style name");

    // <w:pPr><w:outlineLvl w:val="N"/></w:pPr> for headings
    if let Some(level) = outline_level {
        w.write_event(Event::Start(BytesStart::new("w:pPr")))
            .expect("write pPr start");

        let mut lvl = BytesStart::new("w:outlineLvl");
        lvl.push_attribute(("w:val", level.to_string().as_str()));
        w.write_event(Event::Empty(lvl)).expect("write outlineLvl");

        w.write_event(Event::End(BytesEnd::new("w:pPr")))
            .expect("write pPr end");
    }

    w.write_event(Event::End(BytesEnd::new("w:style")))
        .expect("write style end");
}

/// Generate `word/numbering.xml` with bullet (abstractNumId=0, numId=1)
/// and decimal (abstractNumId=1, numId=2) definitions.
fn generate_numbering_xml() -> Vec<u8> {
    let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

    w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .expect("write decl");

    let mut root = BytesStart::new("w:numbering");
    root.push_attribute(("xmlns:w", WML_NS));
    w.write_event(Event::Start(root)).expect("write numbering start");

    // Abstract numbering definition 0: bullet
    write_abstract_num(&mut w, 0, "bullet", "\u{2022}");

    // Abstract numbering definition 1: decimal
    write_abstract_num(&mut w, 1, "decimal", "%1.");

    // Concrete numbering instances
    // numId=1 -> abstractNumId=0 (bullet)
    write_num(&mut w, 1, 0);
    // numId=2 -> abstractNumId=1 (decimal)
    write_num(&mut w, 2, 1);

    w.write_event(Event::End(BytesEnd::new("w:numbering")))
        .expect("write numbering end");

    w.into_inner()
}

/// Write a `<w:abstractNum>` element with a single level.
fn write_abstract_num(
    w: &mut Writer<Vec<u8>>,
    abstract_num_id: u32,
    num_fmt: &str,
    lvl_text: &str,
) {
    let mut elem = BytesStart::new("w:abstractNum");
    elem.push_attribute(("w:abstractNumId", abstract_num_id.to_string().as_str()));
    w.write_event(Event::Start(elem)).expect("write abstractNum start");

    // Single level: ilvl="0"
    let mut lvl = BytesStart::new("w:lvl");
    lvl.push_attribute(("w:ilvl", "0"));
    w.write_event(Event::Start(lvl)).expect("write lvl start");

    let mut fmt = BytesStart::new("w:numFmt");
    fmt.push_attribute(("w:val", num_fmt));
    w.write_event(Event::Empty(fmt)).expect("write numFmt");

    let mut text = BytesStart::new("w:lvlText");
    text.push_attribute(("w:val", lvl_text));
    w.write_event(Event::Empty(text)).expect("write lvlText");

    w.write_event(Event::End(BytesEnd::new("w:lvl")))
        .expect("write lvl end");

    w.write_event(Event::End(BytesEnd::new("w:abstractNum")))
        .expect("write abstractNum end");
}

/// Write a `<w:num>` element mapping numId to abstractNumId.
fn write_num(w: &mut Writer<Vec<u8>>, num_id: u32, abstract_num_id: u32) {
    let mut elem = BytesStart::new("w:num");
    elem.push_attribute(("w:numId", num_id.to_string().as_str()));
    w.write_event(Event::Start(elem)).expect("write num start");

    let mut abs = BytesStart::new("w:abstractNumId");
    abs.push_attribute(("w:val", abstract_num_id.to_string().as_str()));
    w.write_event(Event::Empty(abs)).expect("write abstractNumId");

    w.write_event(Event::End(BytesEnd::new("w:num")))
        .expect("write num end");
}
