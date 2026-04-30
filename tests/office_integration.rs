use std::io::Cursor;

use office_oxide::ir::*;
use office_oxide::{Document, DocumentFormat, OfficeError};

use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::{TargetMode, rel_types};

// ===========================================================================
// Builders
// ===========================================================================

fn make_minimal_docx(document_xml: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
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
    writer.finish().unwrap().into_inner()
}

fn make_docx_with_styles(document_xml: &[u8], styles_xml: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
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

    let styles_part = PartName::new("/word/styles.xml").unwrap();
    writer
        .add_part(
            &styles_part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
            styles_xml,
        )
        .unwrap();
    writer.add_part_rel(&doc_part, rel_types::STYLES, "styles.xml");

    writer.finish().unwrap().into_inner()
}

fn make_docx_with_numbering(document_xml: &[u8], numbering_xml: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
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

    let num_part = PartName::new("/word/numbering.xml").unwrap();
    writer
        .add_part(
            &num_part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
            numbering_xml,
        )
        .unwrap();
    writer.add_part_rel(&doc_part, rel_types::NUMBERING, "numbering.xml");

    writer.finish().unwrap().into_inner()
}

struct XlsxBuilder {
    writer: OpcWriter<Cursor<Vec<u8>>>,
    workbook_part: PartName,
}

impl XlsxBuilder {
    fn new() -> Self {
        let cursor = Cursor::new(Vec::new());
        let mut writer = OpcWriter::new(cursor).unwrap();
        let workbook_part = PartName::new("/xl/workbook.xml").unwrap();
        writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "xl/workbook.xml");
        Self {
            writer,
            workbook_part,
        }
    }

    fn with_workbook(mut self, xml: &[u8]) -> Self {
        self.writer
            .add_part(
                &self.workbook_part,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                xml,
            )
            .unwrap();
        self
    }

    fn with_worksheet(mut self, rel_target: &str, xml: &[u8]) -> Self {
        let part_path = format!("/xl/{rel_target}");
        let part = PartName::new(&part_path).unwrap();
        self.writer
            .add_part(
                &part,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                xml,
            )
            .unwrap();
        self.writer
            .add_part_rel(&self.workbook_part, rel_types::WORKSHEET, rel_target);
        self
    }

    fn build(self) -> Vec<u8> {
        self.writer.finish().unwrap().into_inner()
    }
}

struct PptxBuilder {
    writer: OpcWriter<Cursor<Vec<u8>>>,
    pres_part: PartName,
    slide_count: u32,
}

impl PptxBuilder {
    fn new() -> Self {
        let cursor = Cursor::new(Vec::new());
        let mut writer = OpcWriter::new(cursor).unwrap();
        let pres_part = PartName::new("/ppt/presentation.xml").unwrap();
        writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "ppt/presentation.xml");
        Self {
            writer,
            pres_part,
            slide_count: 0,
        }
    }

    fn with_presentation(mut self, xml: &[u8]) -> Self {
        self.writer
            .add_part(
                &self.pres_part,
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                xml,
            )
            .unwrap();
        self
    }

    fn with_slide(mut self, xml: &[u8]) -> Self {
        self.slide_count += 1;
        let n = self.slide_count;
        let part_path = format!("/ppt/slides/slide{n}.xml");
        let part = PartName::new(&part_path).unwrap();
        self.writer
            .add_part(
                &part,
                "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
                xml,
            )
            .unwrap();
        self.writer.add_part_rel(
            &self.pres_part,
            rel_types::SLIDE,
            &format!("slides/slide{n}.xml"),
        );
        self
    }

    fn build(self) -> Vec<u8> {
        self.writer.finish().unwrap().into_inner()
    }
}

fn pres_xml(slide_ids: &[(u32, &str)]) -> Vec<u8> {
    let mut ids = String::new();
    for (id, rid) in slide_ids {
        ids.push_str(&format!(r#"    <p:sldId id="{id}" r:id="{rid}"/>"#));
        ids.push('\n');
    }
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
{ids}  </p:sldIdLst>
</p:presentation>"#
    )
    .into_bytes()
}

// ===========================================================================
// 1. DOCX via unified API — open, plain_text, to_markdown
// ===========================================================================

#[test]
fn docx_via_unified_api() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Hello from DOCX</w:t></w:r></w:p>
  </w:body>
</w:document>"#;

    let data = make_minimal_docx(xml);
    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Docx).unwrap();

    assert_eq!(doc.format(), DocumentFormat::Docx);
    assert_eq!(doc.plain_text(), "Hello from DOCX");
    assert_eq!(doc.to_markdown(), "Hello from DOCX");
}

// ===========================================================================
// 2. XLSX via unified API — open, plain_text, to_markdown
// ===========================================================================

#[test]
fn xlsx_via_unified_api() {
    let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let ws_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Name</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Score</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Alice</t></is></c>
      <c r="B2"><v>95</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let data = XlsxBuilder::new()
        .with_workbook(wb_xml)
        .with_worksheet("worksheets/sheet1.xml", ws_xml)
        .build();

    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Xlsx).unwrap();
    assert_eq!(doc.format(), DocumentFormat::Xlsx);

    let text = doc.plain_text();
    assert!(text.contains("Name\tScore"), "text was: {text}");
    assert!(text.contains("Alice\t95"), "text was: {text}");

    let md = doc.to_markdown();
    assert!(md.contains("| Name | Score |"), "md was: {md}");
}

// ===========================================================================
// 3. PPTX via unified API — open, plain_text, to_markdown
// ===========================================================================

#[test]
fn pptx_via_unified_api() {
    let slide_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:p><a:r><a:t>Slide Title</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Content"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="2000000"/><a:ext cx="9144000" cy="4000000"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:p><a:r><a:t>Body content here</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>"#;

    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(slide_xml)
        .build();

    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Pptx).unwrap();
    assert_eq!(doc.format(), DocumentFormat::Pptx);

    let text = doc.plain_text();
    assert!(text.contains("Slide Title"), "text was: {text}");
    assert!(text.contains("Body content here"), "text was: {text}");

    let md = doc.to_markdown();
    assert!(md.contains("## Slide Title"), "md was: {md}");
}

// ===========================================================================
// 4. Format detection
// ===========================================================================

#[test]
fn format_detection() {
    assert_eq!(DocumentFormat::from_extension("docx"), Some(DocumentFormat::Docx));
    assert_eq!(DocumentFormat::from_extension("xlsx"), Some(DocumentFormat::Xlsx));
    assert_eq!(DocumentFormat::from_extension("pptx"), Some(DocumentFormat::Pptx));
    assert_eq!(DocumentFormat::from_extension("txt"), None);
    assert_eq!(DocumentFormat::from_extension("pdf"), None);
}

// ===========================================================================
// 5. from_reader with format — works without file extension
// ===========================================================================

#[test]
fn from_reader_with_format() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Reader test</w:t></w:r></w:p>
  </w:body>
</w:document>"#;
    let data = make_minimal_docx(xml);
    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Docx).unwrap();
    assert_eq!(doc.plain_text(), "Reader test");
}

// ===========================================================================
// 6. extract_text() convenience
// ===========================================================================

// This uses open() which requires a file path — we test the function exists
// by calling from_reader and plain_text which is equivalent.
#[test]
fn extract_text_equivalent() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Convenience function</w:t></w:r></w:p>
  </w:body>
</w:document>"#;
    let data = make_minimal_docx(xml);
    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Docx).unwrap();
    assert_eq!(doc.plain_text(), "Convenience function");
}

// ===========================================================================
// 7. to_ir() → DOCX with headings, paragraphs, formatting
// ===========================================================================

#[test]
fn docx_to_ir_headings_and_formatting() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
      <w:r><w:t>Main Title</w:t></w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>Bold text</w:t>
      </w:r>
      <w:r>
        <w:rPr><w:i/></w:rPr>
        <w:t> italic text</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let data = make_minimal_docx(xml);
    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Docx).unwrap();
    let ir = doc.to_ir();

    assert_eq!(ir.sections.len(), 1);
    let section = &ir.sections[0];

    // First element should be heading level 1
    assert!(matches!(&section.elements[0], Element::Heading(h) if h.level == 1));

    // Second element should be paragraph with bold + italic
    if let Element::Paragraph(p) = &section.elements[1] {
        // Should have bold text
        assert!(
            p.content
                .iter()
                .any(|c| matches!(c, InlineContent::Text(s) if s.bold && s.text == "Bold text"))
        );
        // Should have italic text
        assert!(
            p.content
                .iter()
                .any(|c| matches!(c, InlineContent::Text(s) if s.italic))
        );
    } else {
        panic!("expected paragraph, got {:?}", section.elements[1]);
    }
}

// ===========================================================================
// 8. to_ir() → XLSX — sheets as sections, cell grid as table
// ===========================================================================

#[test]
fn xlsx_to_ir_sheets_as_sections() {
    let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sales" sheetId="1" r:id="rId1"/>
    <sheet name="Costs" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"#;

    let ws1 = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Product</t></is></c>
      <c r="B1"><v>100</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let ws2 = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Item</t></is></c>
      <c r="B1"><v>50</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let data = XlsxBuilder::new()
        .with_workbook(wb_xml)
        .with_worksheet("worksheets/sheet1.xml", ws1)
        .with_worksheet("worksheets/sheet2.xml", ws2)
        .build();

    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Xlsx).unwrap();
    let ir = doc.to_ir();

    assert_eq!(ir.sections.len(), 2);
    assert_eq!(ir.sections[0].title.as_deref(), Some("Sales"));
    assert_eq!(ir.sections[1].title.as_deref(), Some("Costs"));

    // Each section should have a table element
    assert!(matches!(&ir.sections[0].elements[0], Element::Table(_)));
    assert!(matches!(&ir.sections[1].elements[0], Element::Table(_)));
}

// ===========================================================================
// 9. to_ir() → PPTX — slides as sections, shapes as elements
// ===========================================================================

#[test]
fn pptx_to_ir_slides_as_sections() {
    let slide_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:p><a:r><a:t>Intro</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Body"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="2000000"/><a:ext cx="5000000" cy="3000000"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:p><a:r><a:t>Welcome</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>"#;

    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(slide_xml)
        .build();

    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Pptx).unwrap();
    let ir = doc.to_ir();

    assert_eq!(ir.sections.len(), 1);
    assert_eq!(ir.sections[0].title.as_deref(), Some("Intro"));

    // Body content should be a paragraph (title is used as section title, not element)
    assert!(
        ir.sections[0]
            .elements
            .iter()
            .any(|e| matches!(e, Element::Paragraph(p) if
                p.content.iter().any(|c| matches!(c, InlineContent::Text(s) if s.text == "Welcome"))
            ))
    );
}

// ===========================================================================
// 10. IR round-trip — to_ir().plain_text() produces reasonable output
// ===========================================================================

#[test]
fn ir_round_trip_plain_text() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
      <w:r><w:t>Report Title</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>Some body text here.</w:t></w:r></w:p>
  </w:body>
</w:document>"#;

    let data = make_minimal_docx(xml);
    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Docx).unwrap();
    let ir_text = doc.to_ir().plain_text();

    assert!(ir_text.contains("Report Title"), "ir_text was: {ir_text}");
    assert!(ir_text.contains("Some body text here."), "ir_text was: {ir_text}");
}

// ===========================================================================
// 11. IR markdown — to_ir().to_markdown() with formatting
// ===========================================================================

#[test]
fn ir_markdown_with_formatting() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:outlineLvl w:val="1"/></w:pPr>
      <w:r><w:t>Section Header</w:t></w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>Important</w:t>
      </w:r>
      <w:r><w:t xml:space="preserve"> note</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let data = make_minimal_docx(xml);
    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Docx).unwrap();
    let ir_md = doc.to_ir().to_markdown();

    assert!(ir_md.contains("## Section Header"), "ir_md was: {ir_md}");
    assert!(ir_md.contains("**Important**"), "ir_md was: {ir_md}");
    assert!(ir_md.contains(" note"), "ir_md was: {ir_md}");
}

// ===========================================================================
// 12. DOCX lists → IR — numbered paragraphs become List elements
// ===========================================================================

#[test]
fn docx_lists_to_ir() {
    let numbering_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#61623;"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>"#;

    let doc_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      </w:pPr>
      <w:r><w:t>First item</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      </w:pPr>
      <w:r><w:t>Second item</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>Not a list</w:t></w:r></w:p>
  </w:body>
</w:document>"#;

    let data = make_docx_with_numbering(doc_xml, numbering_xml);
    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Docx).unwrap();
    let ir = doc.to_ir();

    // First element should be a List with 2 items
    assert!(matches!(&ir.sections[0].elements[0], Element::List(l) if l.items.len() == 2));

    if let Element::List(list) = &ir.sections[0].elements[0] {
        assert!(!list.ordered); // bullet = unordered
        // Check first item text (content is now Vec<Element>)
        let item0_text: String = list.items[0]
            .content
            .iter()
            .filter_map(|e| {
                if let Element::Paragraph(p) = e {
                    Some(
                        p.content
                            .iter()
                            .filter_map(|c| {
                                if let InlineContent::Text(s) = c {
                                    Some(s.text.as_str())
                                } else {
                                    None
                                }
                            })
                            .collect::<String>(),
                    )
                } else {
                    None
                }
            })
            .collect();
        assert!(item0_text.contains("First item"), "item0: {item0_text}");
        let item1_text: String = list.items[1]
            .content
            .iter()
            .filter_map(|e| {
                if let Element::Paragraph(p) = e {
                    Some(
                        p.content
                            .iter()
                            .filter_map(|c| {
                                if let InlineContent::Text(s) = c {
                                    Some(s.text.as_str())
                                } else {
                                    None
                                }
                            })
                            .collect::<String>(),
                    )
                } else {
                    None
                }
            })
            .collect();
        assert!(item1_text.contains("Second item"), "item1: {item1_text}");
    }

    // Second element should be a paragraph
    assert!(matches!(&ir.sections[0].elements[1], Element::Paragraph(_)));
}

// ===========================================================================
// 13. DOCX table merges → IR
// ===========================================================================

#[test]
fn docx_table_merges_to_ir() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:trPr><w:tblHeader/></w:trPr>
        <w:tc>
          <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
          <w:p><w:r><w:t>Merged Header</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
          <w:p><w:r><w:t>Span</w:t></w:r></w:p>
        </w:tc>
        <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:tcPr><w:vMerge/></w:tcPr>
          <w:p/>
        </w:tc>
        <w:tc><w:p><w:r><w:t>B3</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#;

    let data = make_minimal_docx(xml);
    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Docx).unwrap();
    let ir = doc.to_ir();

    if let Element::Table(table) = &ir.sections[0].elements[0] {
        // Header row: single cell with col_span=2
        assert!(table.rows[0].is_header);
        assert_eq!(table.rows[0].cells.len(), 1);
        assert_eq!(table.rows[0].cells[0].col_span, 2);

        // Row 2: cell with row_span=2
        assert_eq!(table.rows[1].cells.len(), 2);
        assert_eq!(table.rows[1].cells[0].row_span, 2);

        // Row 3: vMerge continue cell should be skipped
        assert_eq!(table.rows[2].cells.len(), 1); // only B3
    } else {
        panic!("expected table");
    }
}

// ===========================================================================
// 14. as_docx() / as_xlsx() / as_pptx() — format-specific access
// ===========================================================================

#[test]
fn format_specific_access() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Test</w:t></w:r></w:p>
  </w:body>
</w:document>"#;

    let data = make_minimal_docx(xml);
    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Docx).unwrap();

    assert!(doc.as_docx().is_some());
    assert!(doc.as_xlsx().is_none());
    assert!(doc.as_pptx().is_none());
}

// ===========================================================================
// 15. Unsupported format error
// ===========================================================================

#[test]
fn unsupported_format_error() {
    let result = Document::open("notes.txt");
    assert!(result.is_err());
    match result {
        Err(OfficeError::UnsupportedFormat(_)) => {},
        _ => panic!("expected UnsupportedFormat error"),
    }
}

// ===========================================================================
// 16. DOCX heading via style resolution → IR
// ===========================================================================

#[test]
fn docx_heading_via_style() {
    let styles_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
  </w:style>
</w:styles>"#;

    let doc_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Styled Heading</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let data = make_docx_with_styles(doc_xml, styles_xml);
    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Docx).unwrap();
    let ir = doc.to_ir();

    // Should be detected as heading level 1 via style resolution
    assert!(matches!(&ir.sections[0].elements[0], Element::Heading(h) if h.level == 1));
}

// ===========================================================================
// 17. DOCX table → IR with plain text round-trip
// ===========================================================================

#[test]
fn docx_table_to_ir_plain_text() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#;

    let data = make_minimal_docx(xml);
    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Docx).unwrap();
    let ir = doc.to_ir();
    let text = ir.plain_text();

    assert!(text.contains("A\tB"), "text was: {text}");
    assert!(text.contains("1\t2"), "text was: {text}");
}

// ===========================================================================
// 18. XLSX to IR → plain text
// ===========================================================================

#[test]
fn xlsx_to_ir_plain_text() {
    let wb_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Report" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let ws_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>X</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Y</t></is></c>
    </row>
    <row r="2">
      <c r="A2"><v>1</v></c>
      <c r="B2"><v>2</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let data = XlsxBuilder::new()
        .with_workbook(wb_xml)
        .with_worksheet("worksheets/sheet1.xml", ws_xml)
        .build();

    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Xlsx).unwrap();
    let ir = doc.to_ir();

    assert_eq!(ir.sections.len(), 1);
    assert_eq!(ir.sections[0].title.as_deref(), Some("Report"));

    let md = ir.to_markdown();
    assert!(md.contains("| X | Y |"), "md was: {md}");
}

// ===========================================================================
// 19. PPTX with image → IR has Image element
// ===========================================================================

#[test]
fn pptx_image_to_ir() {
    let slide_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="2" name="Photo" descr="A scenic view"/>
          <p:cNvPicPr/>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill>
          <a:blip r:embed="rId2"/>
        </p:blipFill>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="5000000" cy="3000000"/></a:xfrm>
        </p:spPr>
      </p:pic>
    </p:spTree>
  </p:cSld>
</p:sld>"#;

    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(slide_xml)
        .build();

    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Pptx).unwrap();
    let ir = doc.to_ir();

    assert!(ir.sections[0].elements.iter().any(
        |e| matches!(e, Element::Image(img) if img.alt_text.as_deref() == Some("A scenic view"))
    ));
}

// ===========================================================================
// 20. DOCX strikethrough + hyperlink in IR
// ===========================================================================

#[test]
fn docx_strikethrough_and_hyperlink_in_ir() {
    let cursor = Cursor::new(Vec::new());
    let mut writer = OpcWriter::new(cursor).unwrap();
    let doc_part = PartName::new("/word/document.xml").unwrap();

    let doc_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr><w:strike/></w:rPr>
        <w:t>deleted</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:hyperlink r:id="rId1">
        <w:r><w:t>link text</w:t></w:r>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>"#;

    writer
        .add_part(
            &doc_part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            doc_xml,
        )
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    writer.add_part_rel_with_mode(
        &doc_part,
        rel_types::HYPERLINK,
        "https://example.com",
        TargetMode::External,
    );

    let data = writer.finish().unwrap().into_inner();
    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Docx).unwrap();
    let ir = doc.to_ir();

    // First paragraph has strikethrough
    if let Element::Paragraph(p) = &ir.sections[0].elements[0] {
        assert!(p.content.iter().any(
            |c| matches!(c, InlineContent::Text(s) if s.strikethrough && s.text == "deleted")
        ));
    } else {
        panic!("expected paragraph");
    }

    // Second paragraph has hyperlink
    if let Element::Paragraph(p) = &ir.sections[0].elements[1] {
        assert!(p.content.iter().any(|c| matches!(c, InlineContent::Text(s) if s.hyperlink.as_deref() == Some("https://example.com"))));
    } else {
        panic!("expected paragraph");
    }

    // Check IR markdown output
    let md = ir.to_markdown();
    assert!(md.contains("~~deleted~~"), "md was: {md}");
    assert!(md.contains("[link text](https://example.com)"), "md was: {md}");
}
