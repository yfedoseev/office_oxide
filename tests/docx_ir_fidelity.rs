//! Conversion fidelity for DOCX → `DocumentIR`.
//!
//! Every test here builds a minimal DOCX **in code** (AGENTS.md rule #4 —
//! no committed third-party fixtures) and asserts that a construct the
//! parser already understood actually reaches the IR. The defects these
//! guard against were all of the "parsed then discarded" shape: the parser
//! read the value correctly and the converter dropped it, so a consumer saw
//! a plausible-looking document with the formatting silently missing.

use std::io::Cursor;

use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::rel_types;
use office_oxide::ir::*;
use office_oxide::{Document, DocumentFormat};

// ---------------------------------------------------------------------------
// Fixture builder
// ---------------------------------------------------------------------------

const CT_DOC: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_STYLES: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const CT_NUMBERING: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const CT_CORE: &str = "application/vnd.openxmlformats-package.core-properties+xml";
const CT_HF: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";

struct Docx {
    w: OpcWriter<Cursor<Vec<u8>>>,
    doc_part: PartName,
    body: String,
}

impl Docx {
    fn new(body: &str) -> Self {
        let mut w = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
        let doc_part = PartName::new("/word/document.xml").unwrap();
        w.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
        Self {
            w,
            doc_part,
            body: body.to_string(),
        }
    }

    fn styles(mut self, inner: &str) -> Self {
        let part = PartName::new("/word/styles.xml").unwrap();
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{inner}</w:styles>"#
        );
        self.w.add_part(&part, CT_STYLES, xml.as_bytes()).unwrap();
        self.w
            .add_part_rel(&self.doc_part, rel_types::STYLES, "styles.xml");
        self
    }

    fn numbering(mut self, inner: &str) -> Self {
        let part = PartName::new("/word/numbering.xml").unwrap();
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{inner}</w:numbering>"#
        );
        self.w
            .add_part(&part, CT_NUMBERING, xml.as_bytes())
            .unwrap();
        self.w
            .add_part_rel(&self.doc_part, rel_types::NUMBERING, "numbering.xml");
        self
    }

    /// Add a header or footer part. The relationship id is assigned by the
    /// OPC writer, so the body XML uses the `placeholder` token and this
    /// substitutes the real id into it.
    fn hf(mut self, file: &str, placeholder: &str, text: &str) -> Self {
        let tag = if file.starts_with("header") {
            "hdr"
        } else {
            "ftr"
        };
        let part = PartName::new(&format!("/word/{file}")).unwrap();
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<w:{tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>{text}</w:t></w:r></w:p>
</w:{tag}>"#
        );
        self.w.add_part(&part, CT_HF, xml.as_bytes()).unwrap();
        let rel_type = if tag == "hdr" {
            rel_types::HEADER
        } else {
            rel_types::FOOTER
        };
        let rid = self.w.add_part_rel(&self.doc_part, rel_type, file);
        self.body = self.body.replace(placeholder, &rid);
        self
    }

    fn core_props(mut self, inner: &str) -> Self {
        let part = PartName::new("/docProps/core.xml").unwrap();
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/">{inner}</cp:coreProperties>"#
        );
        self.w.add_part(&part, CT_CORE, xml.as_bytes()).unwrap();
        self.w
            .add_package_rel(rel_types::CORE_PROPERTIES, "docProps/core.xml");
        self
    }

    fn ir(mut self) -> DocumentIR {
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>{}</w:body>
</w:document>"#,
            self.body
        );
        self.w
            .add_part(&self.doc_part, CT_DOC, xml.as_bytes())
            .unwrap();
        let bytes = self.w.finish().unwrap().into_inner();
        Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
            .expect("parse")
            .to_ir()
    }
}

/// First element of the first section.
fn first(ir: &DocumentIR) -> &Element {
    &ir.sections[0].elements[0]
}

fn para(ir: &DocumentIR, idx: usize) -> &Paragraph {
    match &ir.sections[0].elements[idx] {
        Element::Paragraph(p) => p,
        other => panic!("element {idx} is not a paragraph: {other:?}"),
    }
}

fn first_span(p: &Paragraph) -> &TextSpan {
    match &p.content[0] {
        InlineContent::Text(s) => s,
        other => panic!("not a text span: {other:?}"),
    }
}

fn table(ir: &DocumentIR) -> &Table {
    match first(ir) {
        Element::Table(t) => t,
        other => panic!("not a table: {other:?}"),
    }
}

// ---------------------------------------------------------------------------
// #182 — run properties beyond bold/italic
// ---------------------------------------------------------------------------

#[test]
fn run_underline_reaches_the_ir() {
    let ir =
        Docx::new(r#"<w:p><w:r><w:rPr><w:u w:val="double"/></w:rPr><w:t>x</w:t></w:r></w:p>"#).ir();
    assert_eq!(first_span(para(&ir, 0)).underline, Some(UnderlineStyle::Double));
}

#[test]
fn run_caps_smallcaps_and_spacing_reach_the_ir() {
    let ir = Docx::new(
        r#"<w:p><w:r><w:rPr><w:caps/><w:smallCaps/><w:spacing w:val="-20"/></w:rPr>
             <w:t>x</w:t></w:r></w:p>"#,
    )
    .ir();
    let s = first_span(para(&ir, 0));
    assert!(s.all_caps, "w:caps");
    assert!(s.small_caps, "w:smallCaps");
    assert_eq!(s.char_spacing_half_pt, Some(-20));
}

#[test]
fn run_vert_align_reaches_the_ir() {
    let ir = Docx::new(
        r#"<w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r></w:p>"#,
    )
    .ir();
    assert_eq!(first_span(para(&ir, 0)).vertical_align, Some(VerticalAlign::Superscript));
}

#[test]
fn run_highlight_reads_both_encodings() {
    // Word's named palette …
    let ir = Docx::new(
        r#"<w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>x</w:t></w:r></w:p>"#,
    )
    .ir();
    assert_eq!(first_span(para(&ir, 0)).highlight, Some([0xFF, 0xFF, 0x00]));

    // … and the `w:shd` form this crate's own writer emits.
    let ir = Docx::new(
        r#"<w:p><w:r><w:rPr><w:shd w:val="clear" w:fill="00FF00"/></w:rPr><w:t>x</w:t></w:r></w:p>"#,
    )
    .ir();
    assert_eq!(first_span(para(&ir, 0)).highlight, Some([0x00, 0xFF, 0x00]));
}

// ---------------------------------------------------------------------------
// #175 — theme colours and the `w:val` fallback
// ---------------------------------------------------------------------------

#[test]
fn theme_color_falls_back_to_w_val_when_no_theme_part() {
    // `w:themeColor` with no theme part in the package: the literal `w:val`
    // is the producer-supplied fallback and must be used, not discarded.
    let ir = Docx::new(
        r#"<w:p><w:r><w:rPr><w:color w:val="4472C4" w:themeColor="accent1"/></w:rPr>
             <w:t>x</w:t></w:r></w:p>"#,
    )
    .ir();
    assert_eq!(first_span(para(&ir, 0)).color, Some([0x44, 0x72, 0xC4]));
}

#[test]
fn plain_rgb_color_still_reaches_the_ir() {
    let ir =
        Docx::new(r#"<w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>x</w:t></w:r></w:p>"#)
            .ir();
    assert_eq!(first_span(para(&ir, 0)).color, Some([0xFF, 0x00, 0x00]));
}

#[test]
fn auto_color_stays_none() {
    let ir =
        Docx::new(r#"<w:p><w:r><w:rPr><w:color w:val="auto"/></w:rPr><w:t>x</w:t></w:r></w:p>"#)
            .ir();
    assert_eq!(first_span(para(&ir, 0)).color, None);
}

// ---------------------------------------------------------------------------
// #181 — paragraph geometry
// ---------------------------------------------------------------------------

#[test]
fn paragraph_indent_spacing_and_keep_flags_reach_the_ir() {
    let ir = Docx::new(
        r#"<w:p><w:pPr>
             <w:ind w:left="720" w:right="360" w:firstLine="240"/>
             <w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/>
             <w:keepNext/><w:keepLines/><w:pageBreakBefore/>
           </w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#,
    )
    .ir();
    let p = para(&ir, 0);
    assert_eq!(p.indent_left_twips, Some(720));
    assert_eq!(p.indent_right_twips, Some(360));
    assert_eq!(p.first_line_indent_twips, Some(240));
    assert_eq!(p.space_before_twips, Some(120));
    assert_eq!(p.space_after_twips, Some(240));
    assert_eq!(p.line_spacing, Some(LineSpacing::Auto(360)));
    assert!(p.keep_with_next && p.keep_together && p.page_break_before);
}

#[test]
fn hanging_indent_is_a_negative_first_line_indent() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
           <w:r><w:t>x</w:t></w:r></w:p>"#,
    )
    .ir();
    assert_eq!(para(&ir, 0).first_line_indent_twips, Some(-360));
}

#[test]
fn exact_line_spacing_keeps_its_rule() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr>
           <w:r><w:t>x</w:t></w:r></w:p>"#,
    )
    .ir();
    assert_eq!(para(&ir, 0).line_spacing, Some(LineSpacing::Exact(240)));
}

#[test]
fn paragraph_tabs_and_shading_reach_the_ir() {
    let ir = Docx::new(
        r#"<w:p><w:pPr>
             <w:shd w:val="clear" w:fill="EEEEEE"/>
             <w:tabs><w:tab w:val="right" w:pos="9000" w:leader="dot"/></w:tabs>
           </w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#,
    )
    .ir();
    let p = para(&ir, 0);
    assert_eq!(p.background_color, Some([0xEE, 0xEE, 0xEE]));
    assert_eq!(
        p.tabs,
        vec![TabStop {
            position_twips: 9000,
            alignment: TabAlignment::Right,
            leader: TabLeader::Dot,
        }]
    );
}

// ---------------------------------------------------------------------------
// #190 — paragraph borders keep their styling
// ---------------------------------------------------------------------------

#[test]
fn paragraph_borders_keep_style_size_and_colour() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pBdr>
             <w:top w:val="double" w:sz="12" w:space="4" w:color="FF0000"/>
             <w:bottom w:val="single" w:sz="6" w:space="1" w:color="0000FF"/>
           </w:pBdr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#,
    )
    .ir();
    let b = para(&ir, 0).border.as_ref().expect("paragraph border");
    let top = b.top.as_ref().expect("top edge");
    assert_eq!(top.style, BorderStyle::Double);
    assert_eq!(top.size, Some(12));
    assert_eq!(top.space, Some(4));
    assert_eq!(top.color, Some([0xFF, 0x00, 0x00]));
    assert_eq!(b.bottom.as_ref().unwrap().color, Some([0x00, 0x00, 0xFF]));
}

#[test]
fn empty_paragraph_with_only_a_bottom_border_is_still_a_thematic_break() {
    // The horizontal-rule encoding must keep working now that `w:pBdr` is
    // parsed in full rather than narrowed to a boolean.
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6"/></w:pBdr></w:pPr></w:p>"#,
    )
    .ir();
    assert!(matches!(first(&ir), Element::ThematicBreak));
}

// ---------------------------------------------------------------------------
// #180 / #190 — table geometry and borders
// ---------------------------------------------------------------------------

#[test]
fn table_geometry_reaches_the_ir() {
    let ir = Docx::new(
        r#"<w:tbl>
             <w:tblPr>
               <w:tblW w:w="9000" w:type="dxa"/>
               <w:tblInd w:w="360" w:type="dxa"/>
               <w:jc w:val="center"/>
               <w:tblCaption w:val="Quarterly figures"/>
               <w:tblBorders>
                 <w:top w:val="single" w:sz="4" w:color="000000"/>
                 <w:insideV w:val="dotted" w:sz="2" w:color="808080"/>
               </w:tblBorders>
               <w:tblCellMar>
                 <w:top w:w="60" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>
                 <w:bottom w:w="60" w:type="dxa"/><w:right w:w="120" w:type="dxa"/>
               </w:tblCellMar>
             </w:tblPr>
             <w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="5000"/></w:tblGrid>
             <w:tr>
               <w:trPr><w:trHeight w:val="480"/><w:cantSplit/><w:tblHeader/></w:trPr>
               <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>
                       <w:shd w:val="clear" w:fill="DDEEFF"/>
                       <w:vAlign w:val="center"/>
                       <w:textDirection w:val="tbRl"/>
                       <w:tcBorders><w:left w:val="thick" w:sz="24" w:color="00FF00"/></w:tcBorders>
                       <w:tcMar><w:left w:w="200" w:type="dxa"/></w:tcMar>
                     </w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
               <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
             </w:tr>
           </w:tbl>"#,
    )
    .ir();
    let t = table(&ir);
    assert_eq!(t.width_twips, Some(9000));
    assert_eq!(t.indent_left_twips, Some(360));
    assert_eq!(t.alignment, Some(TableAlignment::Center));
    assert_eq!(t.caption.as_deref(), Some("Quarterly figures"));
    assert_eq!(t.column_widths_twips, vec![4000, 5000]);
    assert_eq!(t.cell_padding_twips, Some(120));
    let tb = t.border.as_ref().expect("table border");
    assert_eq!(tb.top.as_ref().unwrap().style, BorderStyle::Single);
    assert_eq!(tb.inside_v.as_ref().unwrap().style, BorderStyle::Dotted);

    let row = &t.rows[0];
    assert_eq!(row.height_twips, Some(480));
    assert!(!row.allow_break, "w:cantSplit");
    assert!(row.is_header && row.repeat_as_header);

    let cell = &row.cells[0];
    assert_eq!(cell.width_twips, Some(4000));
    assert_eq!(cell.background_color, Some([0xDD, 0xEE, 0xFF]));
    assert_eq!(cell.vertical_align, Some(CellVerticalAlign::Center));
    assert_eq!(cell.text_direction, Some(TextDirection::TbRl));
    assert_eq!(cell.border.as_ref().unwrap().left.as_ref().unwrap().style, BorderStyle::Thick);
    assert_eq!(cell.padding.as_ref().unwrap().left_twips, Some(200));
}

#[test]
fn auto_and_percentage_table_widths_report_no_twips() {
    // `w:type="auto"`/`"pct"` carry no absolute measure; reporting one would
    // be a confidently wrong number.
    let ir = Docx::new(
        r#"<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
             <w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    )
    .ir();
    assert_eq!(table(&ir).width_twips, None);
}

// ---------------------------------------------------------------------------
// #185 — page and column breaks
// ---------------------------------------------------------------------------

#[test]
fn page_break_is_a_page_break_not_a_thematic_break() {
    let ir = Docx::new(
        r#"<w:p><w:r><w:t>before</w:t></w:r></w:p>
           <w:p><w:r><w:br w:type="page"/></w:r></w:p>
           <w:p><w:r><w:t>after</w:t></w:r></w:p>"#,
    )
    .ir();
    let kinds: Vec<&str> = ir.sections[0]
        .elements
        .iter()
        .map(|e| match e {
            Element::Paragraph(_) => "P",
            Element::PageBreak => "PB",
            Element::ColumnBreak => "CB",
            Element::ThematicBreak => "TB",
            _ => "?",
        })
        .collect();
    assert_eq!(kinds, vec!["P", "PB", "P"]);
}

#[test]
fn column_break_reaches_the_ir() {
    let ir = Docx::new(r#"<w:p><w:r><w:br w:type="column"/></w:r></w:p>"#).ir();
    assert!(
        ir.sections[0]
            .elements
            .iter()
            .any(|e| matches!(e, Element::ColumnBreak)),
        "expected a ColumnBreak, got {:?}",
        ir.sections[0].elements
    );
}

// ---------------------------------------------------------------------------
// #191 / #177 — section break type and column layout
// ---------------------------------------------------------------------------

#[test]
fn section_break_type_comes_from_w_type_not_the_section_index() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/>
                <w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>
           <w:r><w:t>one</w:t></w:r></w:p>
           <w:p><w:r><w:t>two</w:t></w:r></w:p>
           <w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>"#,
    )
    .ir();
    assert_eq!(ir.sections.len(), 2);
    assert_eq!(ir.sections[0].break_type, SectionBreakType::NextPage);
    assert_eq!(ir.sections[1].break_type, SectionBreakType::Continuous);
}

#[test]
fn column_layout_keeps_space_separator_and_widths() {
    let ir = Docx::new(
        r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>
           <w:sectPr>
             <w:cols w:num="2" w:space="480" w:sep="1">
               <w:col w:w="4320" w:space="480"/><w:col w:w="4320"/>
             </w:cols>
           </w:sectPr>"#,
    )
    .ir();
    let c = ir.sections[0].columns.as_ref().expect("columns");
    assert_eq!(c.count, 2);
    assert_eq!(c.space_twips, Some(480));
    assert!(c.separator);
    assert_eq!(c.column_widths_twips, vec![4320, 4320]);
}

#[test]
fn a_sect_pr_with_no_page_size_reports_no_page_setup() {
    // Synthesising a Letter page for a section that states none would hand
    // the consumer a measurement the document never made.
    let ir = Docx::new(
        r#"<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:type w:val="continuous"/></w:sectPr>"#,
    )
    .ir();
    assert!(ir.sections[0].page_setup.is_none());
}

// ---------------------------------------------------------------------------
// #178 — header / footer routing
// ---------------------------------------------------------------------------

#[test]
fn first_default_and_even_headers_land_in_distinct_slots() {
    let ir = Docx::new(
        r#"<w:p><w:r><w:t>body</w:t></w:r></w:p>
           <w:sectPr>
             <w:headerReference w:type="default" r:id="rIdH1"/>
             <w:headerReference w:type="first" r:id="rIdH2"/>
             <w:footerReference w:type="default" r:id="rIdF1"/>
           </w:sectPr>"#,
    )
    .hf("header1.xml", "rIdH1", "DEFAULT HEADER")
    .hf("header2.xml", "rIdH2", "FIRST HEADER")
    .hf("footer1.xml", "rIdF1", "DEFAULT FOOTER")
    .ir();

    let s = &ir.sections[0];
    let text = |hf: &Option<HeaderFooter>| -> String {
        hf.as_ref()
            .map(|h| {
                h.content
                    .iter()
                    .map(|e| match e {
                        Element::Paragraph(p) => p
                            .content
                            .iter()
                            .filter_map(|c| match c {
                                InlineContent::Text(t) => Some(t.text.as_str()),
                                _ => None,
                            })
                            .collect::<String>(),
                        _ => String::new(),
                    })
                    .collect()
            })
            .unwrap_or_default()
    };
    assert_eq!(text(&s.header), "DEFAULT HEADER");
    assert_eq!(text(&s.first_page_header), "FIRST HEADER");
    assert_eq!(text(&s.footer), "DEFAULT FOOTER");
    // The footer must not have leaked into the header slot, which is what
    // the old cumulative-index split did as soon as the counts were unequal.
    assert!(!text(&s.header).contains("FOOTER"));
}

// ---------------------------------------------------------------------------
// #174 — document metadata
// ---------------------------------------------------------------------------

#[test]
fn core_properties_populate_the_ir_metadata() {
    let ir = Docx::new(r#"<w:p><w:r><w:t>body</w:t></w:r></w:p>"#)
        .core_props(
            r#"<dc:title>Quarterly Report</dc:title>
               <dc:creator>A. Author</dc:creator>
               <dc:subject>Finance</dc:subject>
               <dc:description>Q3 numbers</dc:description>
               <cp:keywords>finance, q3; revenue</cp:keywords>
               <dcterms:created>2026-01-02T03:04:05Z</dcterms:created>
               <dcterms:modified>2026-02-03T04:05:06Z</dcterms:modified>"#,
        )
        .ir();
    let m = &ir.metadata;
    assert_eq!(m.title.as_deref(), Some("Quarterly Report"));
    assert_eq!(m.author.as_deref(), Some("A. Author"));
    assert_eq!(m.subject.as_deref(), Some("Finance"));
    assert_eq!(m.description.as_deref(), Some("Q3 numbers"));
    assert_eq!(m.keywords, vec!["finance", "q3", "revenue"]);
    assert_eq!(m.created.as_deref(), Some("2026-01-02T03:04:05Z"));
    assert_eq!(m.modified.as_deref(), Some("2026-02-03T04:05:06Z"));
}

#[test]
fn title_falls_back_to_the_first_heading_without_core_properties() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>Fallback</w:t></w:r></w:p>"#,
    )
    .ir();
    assert_eq!(ir.metadata.title.as_deref(), Some("Fallback"));
}

// ---------------------------------------------------------------------------
// #187 / #188 — numbering
// ---------------------------------------------------------------------------

const NUMBERING: &str = r#"
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="5"/><w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%1)"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#;

#[test]
fn list_start_number_and_style_reach_the_ir() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
             <w:r><w:t>one</w:t></w:r></w:p>
           <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
             <w:r><w:t>two</w:t></w:r></w:p>"#,
    )
    .numbering(NUMBERING)
    .ir();
    let list = match first(&ir) {
        Element::List(l) => l,
        other => panic!("not a list: {other:?}"),
    };
    assert!(list.ordered);
    assert_eq!(list.start_number, Some(5));
    assert_eq!(list.style, Some(ListStyle::LowerAlpha));
}

#[test]
fn num_id_zero_is_not_a_list() {
    // `<w:numId w:val="0"/>` explicitly removes numbering. Treating it as a
    // list put a bullet in front of an ordinary paragraph.
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr>
             <w:r><w:t>not a bullet</w:t></w:r></w:p>"#,
    )
    .numbering(NUMBERING)
    .ir();
    assert!(matches!(first(&ir), Element::Paragraph(_)), "got {:?}", first(&ir));
}

// ---------------------------------------------------------------------------
// #194 — style-based formatting
// ---------------------------------------------------------------------------

#[test]
fn style_based_run_formatting_reaches_the_ir() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pStyle w:val="Emphatic"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p>"#,
    )
    .styles(
        r#"<w:style w:type="paragraph" w:styleId="Emphatic">
             <w:name w:val="Emphatic"/>
             <w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="112233"/></w:rPr>
           </w:style>"#,
    )
    .ir();
    let s = first_span(para(&ir, 0));
    assert!(s.bold, "bold from the paragraph style");
    assert_eq!(s.font_size_half_pt, Some(28));
    assert_eq!(s.color, Some([0x11, 0x22, 0x33]));
}

#[test]
fn style_inheritance_walks_based_on_and_direct_formatting_wins() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr>
             <w:r><w:rPr><w:i w:val="0"/></w:rPr><w:t>x</w:t></w:r></w:p>"#,
    )
    .styles(
        r#"<w:style w:type="paragraph" w:styleId="Parent">
             <w:name w:val="Parent"/><w:rPr><w:b/><w:i/></w:rPr>
           </w:style>
           <w:style w:type="paragraph" w:styleId="Child">
             <w:name w:val="Child"/><w:basedOn w:val="Parent"/>
           </w:style>"#,
    )
    .ir();
    let s = first_span(para(&ir, 0));
    assert!(s.bold, "inherited from Parent through basedOn");
    assert!(!s.italic, "direct <w:i w:val=\"0\"/> overrides the style");
}

#[test]
fn document_defaults_apply_when_no_style_is_named() {
    let ir = Docx::new(r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>"#)
        .styles(
            r#"<w:docDefaults><w:rPrDefault><w:rPr>
                 <w:rFonts w:ascii="Garamond"/><w:sz w:val="24"/>
               </w:rPr></w:rPrDefault></w:docDefaults>"#,
        )
        .ir();
    let s = first_span(para(&ir, 0));
    assert_eq!(s.font_name.as_deref(), Some("Garamond"));
    assert_eq!(s.font_size_half_pt, Some(24));
}

#[test]
fn character_style_sits_between_paragraph_style_and_direct_formatting() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr>
             <w:r><w:rPr><w:rStyle w:val="Code"/></w:rPr><w:t>x</w:t></w:r></w:p>"#,
    )
    .styles(
        r#"<w:style w:type="paragraph" w:styleId="Body">
             <w:name w:val="Body"/><w:rPr><w:rFonts w:ascii="Georgia"/><w:b/></w:rPr>
           </w:style>
           <w:style w:type="character" w:styleId="Code">
             <w:name w:val="Code"/><w:rPr><w:rFonts w:ascii="Consolas"/></w:rPr>
           </w:style>"#,
    )
    .ir();
    let s = first_span(para(&ir, 0));
    assert_eq!(s.font_name.as_deref(), Some("Consolas"), "character style wins");
    assert!(s.bold, "paragraph style's bold survives");
}

#[test]
fn style_based_paragraph_geometry_reaches_the_ir() {
    let ir =
        Docx::new(r#"<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#)
            .styles(
                r#"<w:style w:type="paragraph" w:styleId="Quote">
                 <w:name w:val="Quote"/>
                 <w:pPr><w:ind w:left="1440"/><w:jc w:val="center"/></w:pPr>
               </w:style>"#,
            )
            .ir();
    let p = para(&ir, 0);
    assert_eq!(p.indent_left_twips, Some(1440));
    assert_eq!(p.alignment, Some(ParagraphAlignment::Center));
}

#[test]
fn a_based_on_cycle_does_not_hang() {
    let ir = Docx::new(r#"<w:p><w:pPr><w:pStyle w:val="A"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#)
        .styles(
            r#"<w:style w:type="paragraph" w:styleId="A">
                 <w:name w:val="A"/><w:basedOn w:val="B"/><w:rPr><w:b/></w:rPr></w:style>
               <w:style w:type="paragraph" w:styleId="B">
                 <w:name w:val="B"/><w:basedOn w:val="A"/></w:style>"#,
        )
        .ir();
    assert!(first_span(para(&ir, 0)).bold);
}

// ---------------------------------------------------------------------------
// #154 — heading detection via style id / name
// ---------------------------------------------------------------------------

#[test]
fn heading_style_id_without_outline_level_is_still_a_heading() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Sub</w:t></w:r></w:p>"#,
    )
    .styles(
        r#"<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style>"#,
    )
    .ir();
    match first(&ir) {
        Element::Heading(h) => assert_eq!(h.level, 2),
        other => panic!("not a heading: {other:?}"),
    }
}

#[test]
fn heading_style_name_without_a_conventional_id_is_still_a_heading() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pStyle w:val="berschrift1"/></w:pPr><w:r><w:t>Top</w:t></w:r></w:p>"#,
    )
    .styles(
        r#"<w:style w:type="paragraph" w:styleId="berschrift1"><w:name w:val="heading 1"/></w:style>"#,
    )
    .ir();
    match first(&ir) {
        Element::Heading(h) => assert_eq!(h.level, 1),
        other => panic!("not a heading: {other:?}"),
    }
}

#[test]
fn a_non_heading_style_stays_a_paragraph() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#,
    )
    .styles(
        r#"<w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/></w:style>"#,
    )
    .ir();
    assert!(matches!(first(&ir), Element::Paragraph(_)));
}

// ---------------------------------------------------------------------------
// #189 — internal hyperlinks
// ---------------------------------------------------------------------------

#[test]
fn internal_anchor_hyperlinks_become_fragment_links() {
    let ir = Docx::new(
        r#"<w:p><w:hyperlink w:anchor="section2"><w:r><w:t>Jump</w:t></w:r></w:hyperlink></w:p>"#,
    )
    .ir();
    assert_eq!(first_span(para(&ir, 0)).hyperlink.as_deref(), Some("#section2"));
}
