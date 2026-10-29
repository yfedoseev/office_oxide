//! A `.docx` body or a `.pptx` slide whose XML is malformed (or cut off)
//! part-way through.
//!
//! The blocks before the damage parsed fine. They are kept, the loss is
//! reported as a warning and in the text, and the rest of the package is read
//! as usual. Damage before the first complete block is still an error.

use std::io::Cursor;

use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::rel_types;
use office_oxide::{Document, DocumentFormat};

const WML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn docx(body_xml: &str) -> Vec<u8> {
    let xml = format!(r#"<w:document xmlns:w="{WML}"><w:body>{body_xml}"#);
    let mut writer = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let part = PartName::new("/word/document.xml").unwrap();
    writer
        .add_part(
            &part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            xml.as_bytes(),
        )
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    writer.finish().unwrap().into_inner()
}

fn para(text: &str) -> String {
    format!("<w:p><w:r><w:t>{text}</w:t></w:r></w:p>")
}

fn open(bytes: Vec<u8>, format: DocumentFormat) -> office_oxide::Result<Document> {
    Document::from_reader(Cursor::new(bytes), format)
}

#[test]
fn test_a_docx_cut_off_mid_paragraph_keeps_the_paragraphs_before_it() {
    let body = format!("{}{}<w:p><w:r><w:t>cut", para("first"), para("second"));
    let doc = open(docx(&body), DocumentFormat::Docx).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("first") && text.contains("second"), "{text}");
    assert!(!text.contains("cut"), "{text}");
    let ir = doc.to_ir();
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("word/document.xml")),
        "{:?}",
        ir.metadata.warnings
    );
    assert!(text.contains("malformed"), "the loss is said in the text: {text}");
}

#[test]
fn test_a_docx_with_a_syntax_error_keeps_the_paragraphs_before_it() {
    let body = format!(
        "{}<w:p><w:r><w:t>x</w:t></w:r><</w:p>{}</w:body></w:document>",
        para("kept"),
        para("lost")
    );
    let doc = open(docx(&body), DocumentFormat::Docx).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("kept") && !text.contains("lost"), "{text}");
}

#[test]
fn test_a_docx_damaged_before_its_first_paragraph_is_still_an_error() {
    assert!(open(docx("<w:p><w:r><w:t>cut"), DocumentFormat::Docx).is_err());
}

#[test]
fn test_a_sound_docx_reports_nothing() {
    let body = format!("{}</w:body></w:document>", para("whole"));
    let doc = open(docx(&body), DocumentFormat::Docx).unwrap();
    assert!(doc.to_ir().metadata.warnings.is_empty());
    assert!(!doc.plain_text().contains("malformed"));
}

fn pptx(slide: &str) -> Vec<u8> {
    let mut writer = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let pres = PartName::new("/ppt/presentation.xml").unwrap();
    let slide_part = PartName::new("/ppt/slides/slide1.xml").unwrap();
    let presentation = format!(
        r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#
    );
    writer
        .add_part(
            &pres,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            presentation.as_bytes(),
        )
        .unwrap();
    writer
        .add_part(
            &slide_part,
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
            slide.as_bytes(),
        )
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "ppt/presentation.xml");
    writer.add_part_rel(&pres, rel_types::SLIDE, "slides/slide1.xml");
    writer.finish().unwrap().into_inner()
}

fn shape(id: u32, text: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="S{id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
    )
}

fn slide(shapes: &str) -> String {
    format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shapes}"#
    )
}

#[test]
fn test_a_pptx_slide_cut_off_mid_shape_keeps_the_shapes_before_it() {
    let cut =
        format!("{}{}<p:sp><p:txBody><a:p><a:r><a:t>lost", shape(2, "alpha"), shape(3, "beta"));
    let doc = open(pptx(&slide(&cut)), DocumentFormat::Pptx).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("alpha") && text.contains("beta"), "{text}");
    assert!(!text.contains("lost"), "{text}");
    let warnings = doc.to_ir().metadata.warnings;
    assert!(warnings.iter().any(|w| w.contains("slides/slide1.xml")), "{warnings:?}");
    assert!(text.contains("malformed"), "the loss is said in the text: {text}");
}

#[test]
fn test_a_pptx_slide_damaged_before_any_shape_is_still_unreadable() {
    let doc = open(
        pptx(&format!(r#"<p:sld xmlns:p="{PML}"><p:cSld><p:spTree><p:sp"#)),
        DocumentFormat::Pptx,
    );
    // Either the deck refuses to open or the slide is reported unreadable; it
    // is never read as an empty, complete slide.
    if let Ok(doc) = doc {
        assert!(!doc.to_ir().metadata.warnings.is_empty());
    }
}
