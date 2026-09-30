//! DOCX reader paths no other test executed: document-default paragraph
//! properties (`w:docDefaults/w:pPrDefault/w:pPr`), and the direct markdown
//! renderer's superscript/subscript wrappers and drawing output.
//!
//! Fixtures are minimal synthetic packages built in code (AGENTS.md
//! rule 4).

use std::io::Cursor;

use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::rel_types;
use office_oxide::docx::{DocxDocument, Justification, StyleSheet};

const W: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart""#;

fn docx(
    body: &str,
    extra: impl FnOnce(&mut OpcWriter<Cursor<Vec<u8>>>, &PartName),
) -> DocxDocument {
    let mut w = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let doc = PartName::new("/word/document.xml").unwrap();
    w.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    w.add_part(
        &doc,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        format!(r#"<?xml version="1.0"?><w:document {W}><w:body>{body}</w:body></w:document>"#)
            .as_bytes(),
    )
    .unwrap();
    extra(&mut w, &doc);
    let bytes = w.finish().unwrap().into_inner();
    DocxDocument::from_reader(Cursor::new(bytes)).unwrap()
}

/// `w:pPrDefault` is parsed into the document defaults, skipping any
/// sibling element ahead of its `w:pPr`.
#[test]
fn test_doc_defaults_paragraph_properties_are_parsed() {
    let sheet = StyleSheet::parse(
        format!(
            r#"<?xml version="1.0"?><w:styles {W}><w:docDefaults>
                 <w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>
                 <w:pPrDefault><w:extLst><w:ext/></w:extLst><w:pPr><w:jc w:val="center"/><w:spacing w:after="160"/></w:pPr></w:pPrDefault>
               </w:docDefaults></w:styles>"#
        )
        .as_bytes(),
    )
    .unwrap();
    let ppr = sheet
        .doc_defaults
        .as_ref()
        .and_then(|d| d.paragraph_properties.as_ref())
        .expect("pPrDefault/pPr parsed");
    assert!(matches!(ppr.justification, Some(Justification::Center)));
    let after = ppr.spacing.as_ref().and_then(|s| s.after).map(|t| t.0);
    assert_eq!(after, Some(160));
}

/// Superscript and subscript have no markdown syntax; the direct
/// renderer wraps them in HTML `<sup>`/`<sub>`.
#[test]
fn test_markdown_wraps_superscript_and_subscript_runs() {
    let doc = docx(
        r#"<w:p><w:r><w:t xml:space="preserve">E=mc</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r><w:r><w:t xml:space="preserve"> and H</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>2</w:t></w:r><w:r><w:t>O</w:t></w:r></w:p>"#,
        |_, _| {},
    );
    let md = doc.to_markdown();
    assert!(md.contains("E=mc<sup>2</sup>"), "{md}");
    assert!(md.contains("H<sub>2</sub>O"), "{md}");
}

/// A drawing renders in the direct markdown as its description (the
/// package has no address a markdown reader could resolve, so it is not
/// an image link, as in the IR renderer), and a chart drawing as the text
/// recovered from its part.
#[test]
fn test_markdown_renders_drawings_and_chart_text() {
    let picture = r#"<w:p><w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="Picture 1" descr="A red square"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#;
    let chart = r#"<w:p><w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="2" name="Chart 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId2"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#;
    let doc = docx(&format!("{picture}{chart}"), |w, doc| {
        let img = PartName::new("/word/media/image1.png").unwrap();
        w.add_part(&img, "image/png", &[0x89, b'P', b'N', b'G', 0, 0, 0, 0])
            .unwrap();
        w.add_part_rel(doc, rel_types::IMAGE, "media/image1.png");
        let chart = PartName::new("/word/charts/chart1.xml").unwrap();
        w.add_part(
            &chart,
            "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Sales by region</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>"#,
        )
        .unwrap();
        w.add_part_rel(doc, rel_types::CHART, "charts/chart1.xml");
    });
    let md = doc.to_markdown();
    assert!(md.contains("*A red square*") && !md.contains("rId1"), "{md}");
    assert!(md.contains("Sales by region"), "{md}");
}
