//! A slide whose relationships part cannot be read keeps its text but loses
//! its images and links, and the loss is reported, naming the relationships
//! part, as every other unreadable part is.

use std::io::{Cursor, Write};

use office_oxide::pptx::PptxDocument;
use office_oxide::{Document, DocumentFormat};

const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

fn rels(entries: &str) -> String {
    format!(r#"<Relationships xmlns="{PKG_REL}">{entries}</Relationships>"#)
}

/// A one-slide `.pptx` whose slide relationships part is `slide_rels`
/// (`None`: no such part).
fn pptx(slide_rels: Option<String>) -> Vec<u8> {
    let mut parts: Vec<(String, String)> = vec![
        ("[Content_Types].xml".into(), r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#.into()),
        ("_rels/.rels".into(), rels(&format!(r#"<Relationship Id="rId1" Type="{REL}/officeDocument" Target="ppt/presentation.xml"/>"#))),
        ("ppt/_rels/presentation.xml.rels".into(), rels(&format!(r#"<Relationship Id="rId1" Type="{REL}/slide" Target="slides/slide1.xml"/>"#))),
        ("ppt/presentation.xml".into(), format!(r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>"#)),
        ("ppt/slides/slide1.xml".into(), format!(r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:rPr><a:hlinkClick r:id="rId7"/></a:rPr><a:t>Slide text survives</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#)),
    ];
    if let Some(r) = slide_rels {
        parts.push(("ppt/slides/_rels/slide1.xml.rels".into(), r));
    }
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = zip::write::SimpleFileOptions::default();
    for (name, body) in parts {
        zip.start_file(name, opts).unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn oversized_link() -> String {
    rels(&format!(
        r#"<Relationship Id="rId7" Type="{REL}/hyperlink" Target="https://example.com/{}" TargetMode="External"/>"#,
        "a".repeat(70_000)
    ))
}

fn check_reported(bytes: Vec<u8>) {
    const PART: &str = "ppt/slides/_rels/slide1.xml.rels";
    let doc = PptxDocument::from_reader(Cursor::new(bytes.clone()))
        .expect("an unreadable relationships part does not fail the deck");
    assert!(doc.plain_text().contains("Slide text survives"));
    assert!(
        doc.unreadable_parts.iter().any(|(p, _)| p.contains(PART)),
        "{PART} not reported: {:?}",
        doc.unreadable_parts
    );
    let ir = Document::from_reader(Cursor::new(bytes), DocumentFormat::Pptx)
        .unwrap()
        .to_ir();
    assert!(ir.metadata.text_truncated);
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("skipped unreadable part") && w.contains(PART)),
        "{PART} not in the warnings: {:?}",
        ir.metadata.warnings
    );
}

#[test]
fn test_slide_relationships_past_the_target_limit_are_reported() {
    check_reported(pptx(Some(oversized_link())));
}

#[test]
fn test_malformed_slide_relationships_are_reported() {
    check_reported(pptx(Some("<Relationships><Relationship Id=\"rId7\"".into())));
}

#[test]
fn test_readable_or_absent_slide_relationships_report_nothing() {
    let ok = rels(&format!(
        r#"<Relationship Id="rId7" Type="{REL}/hyperlink" Target="https://example.com/" TargetMode="External"/>"#
    ));
    for bytes in [pptx(Some(ok)), pptx(None)] {
        let doc = PptxDocument::from_reader(Cursor::new(bytes)).unwrap();
        assert!(doc.unreadable_parts.is_empty(), "{:?}", doc.unreadable_parts);
    }
}
