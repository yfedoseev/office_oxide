use std::io::Cursor;

use docx_oxide::{
    BlockElement, BreakType, DocxDocument, ParagraphContent, RunContent,
};
use office_core::opc::{OpcWriter, PartName};
use office_core::relationships::{rel_types, TargetMode};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

struct DocxBuilder {
    cursor: Option<OpcWriter<Cursor<Vec<u8>>>>,
    doc_part: PartName,
}

impl DocxBuilder {
    fn new() -> Self {
        let cursor = Cursor::new(Vec::new());
        let mut writer = OpcWriter::new(cursor).unwrap();
        let doc_part = PartName::new("/word/document.xml").unwrap();
        writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
        Self {
            cursor: Some(writer),
            doc_part,
        }
    }

    fn with_document(mut self, xml: &[u8]) -> Self {
        self.cursor.as_mut().unwrap().add_part(
            &self.doc_part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            xml,
        ).unwrap();
        self
    }

    fn with_styles(mut self, xml: &[u8]) -> Self {
        let part = PartName::new("/word/styles.xml").unwrap();
        self.cursor.as_mut().unwrap().add_part(
            &part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
            xml,
        ).unwrap();
        self.cursor.as_mut().unwrap().add_part_rel(
            &self.doc_part,
            rel_types::STYLES,
            "styles.xml",
        );
        self
    }

    fn with_numbering(mut self, xml: &[u8]) -> Self {
        let part = PartName::new("/word/numbering.xml").unwrap();
        self.cursor.as_mut().unwrap().add_part(
            &part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
            xml,
        ).unwrap();
        self.cursor.as_mut().unwrap().add_part_rel(
            &self.doc_part,
            rel_types::NUMBERING,
            "numbering.xml",
        );
        self
    }

    fn with_hyperlink(mut self, url: &str) -> Self {
        self.cursor.as_mut().unwrap().add_part_rel_with_mode(
            &self.doc_part,
            rel_types::HYPERLINK,
            url,
            TargetMode::External,
        );
        self
    }

    fn build(mut self) -> Vec<u8> {
        let writer = self.cursor.take().unwrap();
        let result = writer.finish().unwrap();
        result.into_inner()
    }
}

fn parse(data: &[u8]) -> DocxDocument {
    DocxDocument::from_reader(Cursor::new(data.to_vec())).unwrap()
}

// ---------------------------------------------------------------------------
// Integration tests: Round-trip (create then parse)
// ---------------------------------------------------------------------------

#[test]
fn round_trip_simple_paragraphs() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Hello World</w:t></w:r></w:p>
    <w:p><w:r><w:t>Second paragraph</w:t></w:r></w:p>
    <w:p><w:r><w:t>Third paragraph</w:t></w:r></w:p>
  </w:body>
</w:document>"#;
    let data = DocxBuilder::new().with_document(xml).build();
    let doc = parse(&data);

    assert_eq!(doc.body.elements.len(), 3);
    assert_eq!(doc.plain_text(), "Hello World\nSecond paragraph\nThird paragraph");
}

#[test]
fn round_trip_with_styles() {
    let doc_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>My Heading</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Normal text.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let styles_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr><w:sz w:val="24"/></w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
</w:styles>"#;

    let data = DocxBuilder::new()
        .with_document(doc_xml)
        .with_styles(styles_xml)
        .build();
    let doc = parse(&data);

    // Styles parsed
    let styles = doc.styles.as_ref().unwrap();
    assert!(styles.styles.contains_key("Heading1"));
    assert_eq!(styles.resolve_outline_level("Heading1"), Some(0));

    // Heading should produce markdown heading
    let md = doc.to_markdown();
    assert!(md.starts_with("# My Heading"), "markdown was: {md}");
    assert!(md.contains("Normal text."));
}

#[test]
fn round_trip_with_numbering() {
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
  </w:body>
</w:document>"#;

    let num_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="-"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>"#;

    let data = DocxBuilder::new()
        .with_document(doc_xml)
        .with_numbering(num_xml)
        .build();
    let doc = parse(&data);

    assert!(doc.numbering.is_some());
    let md = doc.to_markdown();
    assert!(md.contains("- First item"), "markdown was: {md}");
    assert!(md.contains("- Second item"), "markdown was: {md}");
}

#[test]
fn round_trip_table_with_merged_cells() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblGrid>
        <w:gridCol w:w="3000"/>
        <w:gridCol w:w="3000"/>
        <w:gridCol w:w="3000"/>
      </w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
          <w:p><w:r><w:t>Merged</w:t></w:r></w:p>
        </w:tc>
        <w:tc><w:p><w:r><w:t>Single</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
          <w:p><w:r><w:t>Vert</w:t></w:r></w:p>
        </w:tc>
        <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>C2</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:tcPr><w:vMerge/></w:tcPr>
          <w:p/>
        </w:tc>
        <w:tc><w:p><w:r><w:t>B3</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>C3</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#;
    let data = DocxBuilder::new().with_document(xml).build();
    let doc = parse(&data);

    if let BlockElement::Table(ref t) = doc.body.elements[0] {
        assert_eq!(t.grid.len(), 3);
        assert_eq!(t.rows.len(), 3);

        // First row: merged cell has gridSpan=2
        let cell0 = &t.rows[0].cells[0];
        assert_eq!(
            cell0.properties.as_ref().unwrap().grid_span,
            Some(2)
        );

        // Second row: first cell starts vertical merge
        let cell_v = &t.rows[1].cells[0];
        assert_eq!(
            cell_v.properties.as_ref().unwrap().vertical_merge,
            Some(docx_oxide::table::MergeType::Restart)
        );

        // Third row: first cell continues vertical merge
        let cell_vc = &t.rows[2].cells[0];
        assert_eq!(
            cell_vc.properties.as_ref().unwrap().vertical_merge,
            Some(docx_oxide::table::MergeType::Continue)
        );
    } else {
        panic!("expected table");
    }
}

#[test]
fn round_trip_hyperlink() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink r:id="rId1">
        <w:r><w:t>Click here</w:t></w:r>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>"#;
    let data = DocxBuilder::new()
        .with_document(xml)
        .with_hyperlink("https://example.com")
        .build();
    let doc = parse(&data);

    if let BlockElement::Paragraph(ref p) = doc.body.elements[0] {
        if let ParagraphContent::Hyperlink(ref hl) = p.content[0] {
            match &hl.target {
                docx_oxide::HyperlinkTarget::External(url) => {
                    assert_eq!(url, "https://example.com");
                }
                _ => panic!("expected external hyperlink"),
            }
            assert_eq!(hl.runs.len(), 1);
        } else {
            panic!("expected hyperlink");
        }
    } else {
        panic!("expected paragraph");
    }

    let md = doc.to_markdown();
    assert!(md.contains("[Click here](https://example.com)"), "markdown was: {md}");
}

#[test]
fn round_trip_section_properties() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Content</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
      <w:pgMar w:top="720" w:bottom="720" w:left="1440" w:right="1440"/>
      <w:cols w:num="2"/>
    </w:sectPr>
  </w:body>
</w:document>"#;
    let data = DocxBuilder::new().with_document(xml).build();
    let doc = parse(&data);

    assert_eq!(doc.sections.len(), 1);
    let sect = &doc.sections[0];
    let ps = sect.page_size.as_ref().unwrap();
    assert_eq!(ps.width.0, 15840);
    assert_eq!(ps.height.0, 12240);
    assert_eq!(
        ps.orient,
        Some(docx_oxide::headers::PageOrientation::Landscape)
    );
    let margins = sect.margins.as_ref().unwrap();
    assert_eq!(margins.top.0, 720);
    assert_eq!(margins.left.0, 1440);
    assert_eq!(sect.columns, Some(2));
}

// ---------------------------------------------------------------------------
// Edge cases
// ---------------------------------------------------------------------------

#[test]
fn empty_body() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body/>
</w:document>"#;
    let data = DocxBuilder::new().with_document(xml).build();
    let doc = parse(&data);
    assert!(doc.body.elements.is_empty());
    assert_eq!(doc.plain_text(), "");
    assert_eq!(doc.to_markdown(), "");
}

#[test]
fn only_tables() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Only</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Tables</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#;
    let data = DocxBuilder::new().with_document(xml).build();
    let doc = parse(&data);
    assert_eq!(doc.body.elements.len(), 1);
    assert!(matches!(doc.body.elements[0], BlockElement::Table(_)));
    assert_eq!(doc.plain_text(), "Only\tTables");
}

#[test]
fn paragraph_with_page_break() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Before break</w:t>
        <w:br w:type="page"/>
        <w:t>After break</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;
    let data = DocxBuilder::new().with_document(xml).build();
    let doc = parse(&data);

    if let BlockElement::Paragraph(ref p) = doc.body.elements[0] {
        if let ParagraphContent::Run(ref run) = p.content[0] {
            assert_eq!(run.content.len(), 3);
            assert!(matches!(run.content[1], RunContent::Break(BreakType::Page)));
        } else {
            panic!("expected run");
        }
    } else {
        panic!("expected paragraph");
    }
    assert!(doc.plain_text().contains("Before break"));
}

#[test]
fn formatting_bold_italic_strikethrough() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr><w:b/><w:i/></w:rPr>
        <w:t>bold italic</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr><w:strike/></w:rPr>
        <w:t>struck</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;
    let data = DocxBuilder::new().with_document(xml).build();
    let doc = parse(&data);
    let md = doc.to_markdown();
    assert!(md.contains("***bold italic***"), "markdown was: {md}");
    assert!(md.contains("~~struck~~"), "markdown was: {md}");
}

#[test]
fn internal_bookmark_hyperlink() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:hyperlink w:anchor="section1">
        <w:r><w:t>Jump to section</w:t></w:r>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>"#;
    let data = DocxBuilder::new().with_document(xml).build();
    let doc = parse(&data);

    if let BlockElement::Paragraph(ref p) = doc.body.elements[0] {
        if let ParagraphContent::Hyperlink(ref hl) = p.content[0] {
            match &hl.target {
                docx_oxide::HyperlinkTarget::Internal(anchor) => {
                    assert_eq!(anchor, "section1");
                }
                _ => panic!("expected internal hyperlink"),
            }
        }
    }
    let md = doc.to_markdown();
    assert!(md.contains("[Jump to section](#section1)"), "markdown was: {md}");
}

#[test]
fn nested_table_in_cell() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p><w:r><w:t>Outer</w:t></w:r></w:p>
          <w:tbl>
            <w:tr>
              <w:tc><w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:tc>
            </w:tr>
          </w:tbl>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#;
    let data = DocxBuilder::new().with_document(xml).build();
    let doc = parse(&data);

    if let BlockElement::Table(ref outer) = doc.body.elements[0] {
        let cell = &outer.rows[0].cells[0];
        assert_eq!(cell.content.len(), 2);
        assert!(matches!(cell.content[0], BlockElement::Paragraph(_)));
        assert!(matches!(cell.content[1], BlockElement::Table(_)));
    } else {
        panic!("expected table");
    }
}

#[test]
fn complex_document_with_everything() {
    let doc_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Document Title</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t xml:space="preserve">This is </w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>
      <w:r><w:t xml:space="preserve"> and </w:t></w:r>
      <w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>
      <w:r><w:t xml:space="preserve"> text.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      </w:pPr>
      <w:r><w:t>List item one</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      </w:pPr>
      <w:r><w:t>List item two</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Name</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Alpha</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>100</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p>
      <w:hyperlink r:id="rId3">
        <w:r><w:t>Link text</w:t></w:r>
      </w:hyperlink>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>
    </w:sectPr>
  </w:body>
</w:document>"#;

    let styles_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/></w:rPr>
  </w:style>
</w:styles>"#;

    let num_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

    let data = DocxBuilder::new()
        .with_document(doc_xml)
        .with_styles(styles_xml)
        .with_numbering(num_xml)
        .with_hyperlink("https://example.com")
        .build();
    let doc = parse(&data);

    // Verify structure
    assert_eq!(doc.body.elements.len(), 6); // heading + text + 2 list items + table + link para
    assert!(doc.styles.is_some());
    assert!(doc.numbering.is_some());
    assert_eq!(doc.sections.len(), 1);

    // Plain text
    let text = doc.plain_text();
    assert!(text.contains("Document Title"));
    assert!(text.contains("bold"));
    assert!(text.contains("italic"));
    assert!(text.contains("List item one"));
    assert!(text.contains("Name\tValue"));
    assert!(text.contains("Link text"));

    // Markdown
    let md = doc.to_markdown();
    assert!(md.contains("# **Document Title**"), "markdown was: {md}");
    assert!(md.contains("**bold**"), "markdown was: {md}");
    assert!(md.contains("*italic*"), "markdown was: {md}");
    assert!(md.contains("1. List item one"), "markdown was: {md}");
    assert!(md.contains("| Name | Value |"), "markdown was: {md}");
    assert!(md.contains("[Link text](https://example.com)"), "markdown was: {md}");
}

// ---------------------------------------------------------------------------
// Error handling
// ---------------------------------------------------------------------------

#[test]
fn missing_document_part() {
    // Create a package without document.xml — should error
    let cursor = Cursor::new(Vec::new());
    let mut writer = OpcWriter::new(cursor).unwrap();
    let part = PartName::new("/word/other.xml").unwrap();
    writer.add_part(&part, "application/xml", b"<root/>").unwrap();
    // Note: no officeDocument relationship
    let result = writer.finish().unwrap();
    let data = result.into_inner();

    let result = DocxDocument::from_reader(Cursor::new(data));
    assert!(result.is_err());
}
