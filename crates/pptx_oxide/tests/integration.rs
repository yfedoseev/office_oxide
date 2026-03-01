use std::io::Cursor;

use office_core::opc::{OpcWriter, PartName};
use office_core::relationships::{rel_types, TargetMode};
use pptx_oxide::PptxDocument;

// ---------------------------------------------------------------------------
// PptxBuilder helper
// ---------------------------------------------------------------------------

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
        let rel_target = format!("slides/slide{n}.xml");
        self.writer
            .add_part_rel(&self.pres_part, rel_types::SLIDE, &rel_target);
        self
    }

    fn with_slide_notes(mut self, slide_num: u32, notes_xml: &[u8]) -> Self {
        let notes_path = format!("/ppt/notesSlides/notesSlide{slide_num}.xml");
        let notes_part = PartName::new(&notes_path).unwrap();
        self.writer
            .add_part(
                &notes_part,
                "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
                notes_xml,
            )
            .unwrap();
        let slide_part = PartName::new(&format!("/ppt/slides/slide{slide_num}.xml")).unwrap();
        self.writer.add_part_rel(
            &slide_part,
            rel_types::NOTES_SLIDE,
            &format!("../notesSlides/notesSlide{slide_num}.xml"),
        );
        self
    }

    fn with_slide_hyperlink(mut self, slide_num: u32, url: &str) -> Self {
        let slide_part = PartName::new(&format!("/ppt/slides/slide{slide_num}.xml")).unwrap();
        self.writer.add_part_rel_with_mode(
            &slide_part,
            rel_types::HYPERLINK,
            url,
            TargetMode::External,
        );
        self
    }

    fn with_theme(mut self, xml: &[u8]) -> Self {
        let part = PartName::new("/ppt/theme/theme1.xml").unwrap();
        self.writer
            .add_part(
                &part,
                "application/vnd.openxmlformats-officedocument.theme+xml",
                xml,
            )
            .unwrap();
        self.writer
            .add_part_rel(&self.pres_part, rel_types::THEME, "theme/theme1.xml");
        self
    }

    fn build(self) -> Vec<u8> {
        let result = self.writer.finish().unwrap();
        result.into_inner()
    }
}

fn parse(data: &[u8]) -> PptxDocument {
    PptxDocument::from_reader(Cursor::new(data.to_vec())).unwrap()
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
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#
    )
    .into_bytes()
}

fn slide_xml(shapes: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
      {shapes}
    </p:spTree>
  </p:cSld>
</p:sld>"#
    )
    .into_bytes()
}

fn notes_xml(text: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Slide Image"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph type="sldImg"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Notes"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:p><a:r><a:t>{text}</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>"#
    )
    .into_bytes()
}

fn auto_shape(id: u32, name: &str, text: &str, x: i64, y: i64, cx: i64, cy: i64) -> String {
    format!(
        r#"<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{id}" name="{name}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="{x}" y="{y}"/>
      <a:ext cx="{cx}" cy="{cy}"/>
    </a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:p><a:r><a:t>{text}</a:t></a:r></a:p>
  </p:txBody>
</p:sp>"#
    )
}

fn title_shape(id: u32, text: &str) -> String {
    format!(
        r#"<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{id}" name="Title"/>
    <p:cNvSpPr/>
    <p:nvPr><p:ph type="title"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="457200" y="274638"/>
      <a:ext cx="8229600" cy="1143000"/>
    </a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:p><a:r><a:t>{text}</a:t></a:r></a:p>
  </p:txBody>
</p:sp>"#
    )
}

// ---------------------------------------------------------------------------
// 1. Simple slide with text
// ---------------------------------------------------------------------------

#[test]
fn simple_slide_with_text() {
    let shapes = auto_shape(2, "TextBox 1", "Hello World", 457200, 1600200, 8229600, 4525963);
    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(&slide_xml(&shapes))
        .build();

    let doc = parse(&data);
    assert_eq!(doc.slides.len(), 1);
    assert_eq!(doc.plain_text(), "Hello World");
}

// ---------------------------------------------------------------------------
// 2. Multiple shapes with spatial sorting
// ---------------------------------------------------------------------------

#[test]
fn multiple_shapes_spatial_sort() {
    let shapes = format!(
        "{}{}{}",
        auto_shape(2, "Bottom", "Third", 100, 5000000, 4000000, 500000),
        auto_shape(3, "Top", "First", 200, 100000, 4000000, 500000),
        auto_shape(4, "Middle", "Second", 50, 2500000, 4000000, 500000),
    );
    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(&slide_xml(&shapes))
        .build();

    let doc = parse(&data);
    let text = doc.plain_text();
    assert_eq!(text, "First\n\nSecond\n\nThird");
}

// ---------------------------------------------------------------------------
// 3. Rich text formatting
// ---------------------------------------------------------------------------

#[test]
fn rich_text_formatting() {
    let shapes = r#"<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Text"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="9000" cy="1000"/></a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:p>
      <a:r><a:rPr b="1"/><a:t>bold</a:t></a:r>
      <a:r><a:t> and </a:t></a:r>
      <a:r><a:rPr i="1"/><a:t>italic</a:t></a:r>
      <a:r><a:t> and </a:t></a:r>
      <a:r><a:rPr strike="sngStrike"/><a:t>struck</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>"#;

    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(&slide_xml(shapes))
        .build();

    let doc = parse(&data);
    let md = doc.to_markdown();
    assert!(md.contains("**bold** and *italic* and ~~struck~~"));
    assert_eq!(doc.plain_text(), "bold and italic and struck");
}

// ---------------------------------------------------------------------------
// 4. Group shapes
// ---------------------------------------------------------------------------

#[test]
fn group_shapes() {
    let shapes = format!(
        r#"<p:grpSp>
  <p:nvGrpSpPr>
    <p:cNvPr id="5" name="Group 1"/>
    <p:cNvGrpSpPr/>
    <p:nvPr/>
  </p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm>
      <a:off x="100" y="100"/>
      <a:ext cx="5000" cy="3000"/>
    </a:xfrm>
  </p:grpSpPr>
  {}
  {}
</p:grpSp>"#,
        auto_shape(6, "Child 1", "First child", 100, 200, 2000, 500),
        auto_shape(7, "Child 2", "Second child", 100, 800, 2000, 500),
    );

    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(&slide_xml(&shapes))
        .build();

    let doc = parse(&data);
    let text = doc.plain_text();
    assert!(text.contains("First child"));
    assert!(text.contains("Second child"));
}

// ---------------------------------------------------------------------------
// 5. Table in GraphicFrame
// ---------------------------------------------------------------------------

#[test]
fn table_graphic_frame() {
    let shapes = r#"<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="10" name="Table 1"/>
    <p:cNvGraphicFramePr/>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm>
    <a:off x="0" y="0"/>
    <a:ext cx="9144000" cy="3000000"/>
  </p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl>
        <a:tblGrid>
          <a:gridCol w="3048000"/>
          <a:gridCol w="3048000"/>
        </a:tblGrid>
        <a:tr h="370840">
          <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Header1</a:t></a:r></a:p></a:txBody></a:tc>
          <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Header2</a:t></a:r></a:p></a:txBody></a:tc>
        </a:tr>
        <a:tr h="370840">
          <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Cell1</a:t></a:r></a:p></a:txBody></a:tc>
          <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Cell2</a:t></a:r></a:p></a:txBody></a:tc>
        </a:tr>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>"#;

    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(&slide_xml(shapes))
        .build();

    let doc = parse(&data);
    let plain = doc.plain_text();
    assert!(plain.contains("Header1\tHeader2"));
    assert!(plain.contains("Cell1\tCell2"));

    let md = doc.to_markdown();
    assert!(md.contains("| Header1 | Header2 |"));
    assert!(md.contains("| --- | --- |"));
    assert!(md.contains("| Cell1 | Cell2 |"));
}

// ---------------------------------------------------------------------------
// 6. Hyperlinks
// ---------------------------------------------------------------------------

#[test]
fn hyperlinks() {
    // Hyperlink rels are added to slide1, generating rId1 for the hyperlink
    // (since notes rels aren't added, the first rel for slide1 is the hyperlink)
    let shapes = r#"<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Text"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="9000" cy="1000"/></a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:p>
      <a:r>
        <a:rPr>
          <a:hlinkClick r:id="rId1"/>
        </a:rPr>
        <a:t>Click me</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>"#;

    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide_hyperlink(1, "https://example.com")
        .with_slide(&slide_xml(shapes))
        .build();

    let doc = parse(&data);
    let md = doc.to_markdown();
    assert!(md.contains("[Click me](https://example.com)"));
    assert_eq!(doc.plain_text(), "Click me");
}

// ---------------------------------------------------------------------------
// 7. Notes slide
// ---------------------------------------------------------------------------

#[test]
fn notes_slide() {
    let shapes = auto_shape(2, "Content", "Main content", 0, 0, 9000, 5000);
    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(&slide_xml(&shapes))
        .with_slide_notes(1, &notes_xml("These are speaker notes"))
        .build();

    let doc = parse(&data);
    assert_eq!(doc.slides[0].notes.as_deref(), Some("These are speaker notes"));

    let text = doc.plain_text();
    assert!(text.contains("Main content"));
    assert!(text.contains("[Notes]\nThese are speaker notes"));

    let md = doc.to_markdown();
    assert!(md.contains("> These are speaker notes"));
}

// ---------------------------------------------------------------------------
// 8. Multiple slides
// ---------------------------------------------------------------------------

#[test]
fn multiple_slides() {
    let s1 = format!(
        "{}{}",
        title_shape(2, "Introduction"),
        auto_shape(3, "Body", "Welcome everyone", 0, 2000000, 9000000, 4000000)
    );
    let s2 = format!(
        "{}{}",
        title_shape(2, "Main Point"),
        auto_shape(3, "Body", "Here is the point", 0, 2000000, 9000000, 4000000)
    );
    let s3 = format!(
        "{}{}",
        title_shape(2, "Conclusion"),
        auto_shape(3, "Body", "Thank you", 0, 2000000, 9000000, 4000000)
    );

    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[
            (256, "rId1"),
            (257, "rId2"),
            (258, "rId3"),
        ]))
        .with_slide(&slide_xml(&s1))
        .with_slide(&slide_xml(&s2))
        .with_slide(&slide_xml(&s3))
        .build();

    let doc = parse(&data);
    assert_eq!(doc.slides.len(), 3);

    let text = doc.plain_text();
    assert!(text.contains("Introduction"));
    assert!(text.contains("Welcome everyone"));
    assert!(text.contains("---"));
    assert!(text.contains("Thank you"));

    let md = doc.to_markdown();
    assert!(md.contains("## Introduction"));
    assert!(md.contains("## Main Point"));
    assert!(md.contains("## Conclusion"));
}

// ---------------------------------------------------------------------------
// 9. Placeholder types (title, subtitle, body)
// ---------------------------------------------------------------------------

#[test]
fn placeholder_types() {
    let shapes = r#"<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Title"/>
    <p:cNvSpPr/>
    <p:nvPr><p:ph type="ctrTitle"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="9000" cy="2000"/></a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:p><a:r><a:t>Centered Title</a:t></a:r></a:p>
  </p:txBody>
</p:sp>
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="3" name="Subtitle"/>
    <p:cNvSpPr/>
    <p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="3000"/><a:ext cx="9000" cy="1500"/></a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:p><a:r><a:t>A subtitle</a:t></a:r></a:p>
  </p:txBody>
</p:sp>"#;

    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(&slide_xml(shapes))
        .build();

    let doc = parse(&data);
    // ctrTitle is treated as title for markdown heading
    let md = doc.to_markdown();
    assert!(md.contains("## Centered Title"));
    assert!(md.contains("A subtitle"));
}

// ---------------------------------------------------------------------------
// 10. Text fields (slidenum, datetime)
// ---------------------------------------------------------------------------

#[test]
fn text_fields() {
    let shapes = r#"<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Slide Number"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="9000" cy="1000"/></a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:p>
      <a:r><a:t>Slide </a:t></a:r>
      <a:fld type="slidenum">
        <a:rPr/>
        <a:t>5</a:t>
      </a:fld>
    </a:p>
  </p:txBody>
</p:sp>"#;

    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(&slide_xml(shapes))
        .build();

    let doc = parse(&data);
    assert_eq!(doc.plain_text(), "Slide 5");
}

// ---------------------------------------------------------------------------
// 11. Empty shapes
// ---------------------------------------------------------------------------

#[test]
fn empty_shapes() {
    let shapes = r#"<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Empty"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr/>
    <a:p/>
  </p:txBody>
</p:sp>"#;

    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(&slide_xml(shapes))
        .build();

    let doc = parse(&data);
    assert_eq!(doc.plain_text(), "");
}

// ---------------------------------------------------------------------------
// 12. Picture with alt text
// ---------------------------------------------------------------------------

#[test]
fn picture_alt_text() {
    let shapes = r#"<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="3" name="Picture 1" descr="A cute cat"/>
    <p:cNvPicPr/>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId99"/>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="0" y="0"/>
      <a:ext cx="5000" cy="3000"/>
    </a:xfrm>
  </p:spPr>
</p:pic>"#;

    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(&slide_xml(shapes))
        .build();

    let doc = parse(&data);
    assert_eq!(doc.plain_text(), "A cute cat");
    let md = doc.to_markdown();
    assert!(md.contains("![A cute cat]()"));
}

// ---------------------------------------------------------------------------
// 13. Plain text output formatting
// ---------------------------------------------------------------------------

#[test]
fn plain_text_output() {
    let shapes = format!(
        "{}{}",
        title_shape(2, "My Slide"),
        auto_shape(3, "Body", "Line one\nLine two", 0, 2000000, 9000000, 4000000)
    );

    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(&slide_xml(&shapes))
        .build();

    let doc = parse(&data);
    let text = doc.slide_plain_text(0).unwrap();
    // Title comes first (lower y), then body
    assert!(text.starts_with("My Slide"));
    assert!(text.contains("Line one\nLine two"));
}

// ---------------------------------------------------------------------------
// 14. Markdown output with all features
// ---------------------------------------------------------------------------

#[test]
fn markdown_output_combined() {
    let shapes = format!(
        r#"{}
{}
<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="10" name="Table"/>
    <p:cNvGraphicFramePr/>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm>
    <a:off x="0" y="5000000"/>
    <a:ext cx="9000000" cy="2000000"/>
  </p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl>
        <a:tr h="370840">
          <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Col A</a:t></a:r></a:p></a:txBody></a:tc>
          <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Col B</a:t></a:r></a:p></a:txBody></a:tc>
        </a:tr>
        <a:tr h="370840">
          <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>1</a:t></a:r></a:p></a:txBody></a:tc>
          <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>2</a:t></a:r></a:p></a:txBody></a:tc>
        </a:tr>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>"#,
        title_shape(2, "Data Slide"),
        auto_shape(3, "Body", "Some data below:", 0, 2000000, 9000000, 2000000),
    );

    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(&slide_xml(&shapes))
        .with_slide_notes(1, &notes_xml("Remember to explain the data"))
        .build();

    let doc = parse(&data);
    let md = doc.to_markdown();
    assert!(md.contains("## Data Slide"));
    assert!(md.contains("Some data below:"));
    assert!(md.contains("| Col A | Col B |"));
    assert!(md.contains("| --- | --- |"));
    assert!(md.contains("| 1 | 2 |"));
    assert!(md.contains("> Remember to explain the data"));
}

// ---------------------------------------------------------------------------
// 15. Missing notes
// ---------------------------------------------------------------------------

#[test]
fn missing_notes() {
    let shapes = auto_shape(2, "Content", "Just content", 0, 0, 9000, 5000);
    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(&slide_xml(&shapes))
        .build();

    let doc = parse(&data);
    assert!(doc.slides[0].notes.is_none());
}

// ---------------------------------------------------------------------------
// 16. Presentation info (slide count, slide size)
// ---------------------------------------------------------------------------

#[test]
fn presentation_info() {
    let s1 = auto_shape(2, "T", "A", 0, 0, 1, 1);
    let s2 = auto_shape(2, "T", "B", 0, 0, 1, 1);
    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1"), (257, "rId2")]))
        .with_slide(&slide_xml(&s1))
        .with_slide(&slide_xml(&s2))
        .build();

    let doc = parse(&data);
    assert_eq!(doc.presentation.slides.len(), 2);
    let size = doc.presentation.slide_size.as_ref().unwrap();
    assert_eq!(size.cx, 9144000);
    assert_eq!(size.cy, 6858000);
}

// ---------------------------------------------------------------------------
// 17. Connector shapes (no text extracted)
// ---------------------------------------------------------------------------

#[test]
fn connector_no_text() {
    let shapes = format!(
        r#"{}
<p:cxnSp>
  <p:nvCxnSpPr>
    <p:cNvPr id="5" name="Connector"/>
    <p:cNvCxnSpPr/>
    <p:nvPr/>
  </p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="5000"/><a:ext cx="5000" cy="0"/></a:xfrm>
  </p:spPr>
</p:cxnSp>"#,
        auto_shape(2, "Text", "Visible text", 0, 0, 9000, 2000)
    );

    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(&slide_xml(&shapes))
        .build();

    let doc = parse(&data);
    // Connector should not contribute text
    assert_eq!(doc.plain_text(), "Visible text");
}

// ---------------------------------------------------------------------------
// 18. Theme parsing
// ---------------------------------------------------------------------------

#[test]
fn theme_parsing() {
    let theme_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>"#;

    let shapes = auto_shape(2, "Text", "Hello", 0, 0, 9000, 5000);
    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(&slide_xml(&shapes))
        .with_theme(theme_xml)
        .build();

    let doc = parse(&data);
    assert!(doc.theme.is_some());
}

// ---------------------------------------------------------------------------
// 19. Line breaks in text
// ---------------------------------------------------------------------------

#[test]
fn line_breaks_in_text() {
    let shapes = r#"<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Text"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="9000" cy="2000"/></a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:p>
      <a:r><a:t>First</a:t></a:r>
      <a:br/>
      <a:r><a:t>Second</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>"#;

    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(&slide_xml(shapes))
        .build();

    let doc = parse(&data);
    assert_eq!(doc.plain_text(), "First\nSecond");
}

// ---------------------------------------------------------------------------
// 20. Outline levels in markdown
// ---------------------------------------------------------------------------

#[test]
fn outline_levels_markdown() {
    let shapes = r#"<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Body"/>
    <p:cNvSpPr/>
    <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="2000"/><a:ext cx="9000" cy="4000"/></a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:p><a:pPr lvl="0"/><a:r><a:t>Top level</a:t></a:r></a:p>
    <a:p><a:pPr lvl="1"/><a:r><a:t>Sub item</a:t></a:r></a:p>
    <a:p><a:pPr lvl="2"/><a:r><a:t>Sub sub item</a:t></a:r></a:p>
  </p:txBody>
</p:sp>"#;

    let data = PptxBuilder::new()
        .with_presentation(&pres_xml(&[(256, "rId1")]))
        .with_slide(&slide_xml(shapes))
        .build();

    let doc = parse(&data);
    let md = doc.to_markdown();
    assert!(md.contains("Top level"));
    assert!(md.contains("  - Sub item"));
    assert!(md.contains("    - Sub sub item"));
}
