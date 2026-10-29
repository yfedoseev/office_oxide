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

/// A `.docx` whose body references one footnote and one header; `footnotes`
/// and `header` are the parts' raw XML (which a test may damage).
fn docx_with_ancillary(footnotes: &str, header: &str) -> Vec<u8> {
    let mut writer = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let main = PartName::new("/word/document.xml").unwrap();
    let body = format!(
        r#"<w:document xmlns:w="{WML}" xmlns:r="{REL}"><w:body><w:p><w:r><w:t>body text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr></w:body></w:document>"#
    );
    writer
        .add_part(
            &main,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            body.as_bytes(),
        )
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    let notes = PartName::new("/word/footnotes.xml").unwrap();
    writer
        .add_part(
            &notes,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
            footnotes.as_bytes(),
        )
        .unwrap();
    writer.add_part_rel(&main, rel_types::FOOTNOTES, "footnotes.xml");
    let hdr = PartName::new("/word/header1.xml").unwrap();
    writer
        .add_part(
            &hdr,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
            header.as_bytes(),
        )
        .unwrap();
    writer.add_part_rel(&main, rel_types::HEADER, "header1.xml");
    writer.finish().unwrap().into_inner()
}

#[test]
fn test_a_damaged_footnotes_or_header_part_keeps_its_first_blocks() {
    let notes = format!(
        r#"<w:footnotes xmlns:w="{WML}"><w:footnote w:id="1"><w:p><w:r><w:t>first note</w:t></w:r></w:p><w:p><w:r><w:t>second para</w:t></w:r></w:p><w:p><w:r a=""#
    );
    let header = format!(
        r#"<w:hdr xmlns:w="{WML}"><w:p><w:r><w:t>kept header</w:t></w:r></w:p><w:p><w:r a=""#
    );
    // The ids are the document's own (rIdH is not in a relationship part), so
    // the header is reached through its relationship by type.
    let doc = open(docx_with_ancillary(&notes, &header), DocumentFormat::Docx).unwrap();
    let warnings = doc.to_ir().metadata.warnings;
    assert!(warnings.iter().any(|w| w.contains("word/footnotes.xml")), "{warnings:?}");
    let text = doc.plain_text();
    assert!(text.contains("first note") && text.contains("second para"), "{text}");
    assert!(!text.contains('"'), "{text}");
}

#[test]
fn test_sound_ancillary_parts_report_nothing() {
    let notes = format!(
        r#"<w:footnotes xmlns:w="{WML}"><w:footnote w:id="1"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:footnote></w:footnotes>"#
    );
    let header = format!(r#"<w:hdr xmlns:w="{WML}"><w:p><w:r><w:t>head</w:t></w:r></w:p></w:hdr>"#);
    let doc = open(docx_with_ancillary(&notes, &header), DocumentFormat::Docx).unwrap();
    assert!(doc.to_ir().metadata.warnings.is_empty(), "{:?}", doc.to_ir().metadata.warnings);
}

fn pptx_with_notes(notes: &str) -> Vec<u8> {
    let mut writer = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let pres = PartName::new("/ppt/presentation.xml").unwrap();
    let slide_part = PartName::new("/ppt/slides/slide1.xml").unwrap();
    let notes_part = PartName::new("/ppt/notesSlides/notesSlide1.xml").unwrap();
    let presentation = format!(
        r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#
    );
    let ct = "application/vnd.openxmlformats-officedocument.presentationml";
    writer
        .add_part(&pres, &format!("{ct}.presentation.main+xml"), presentation.as_bytes())
        .unwrap();
    writer
        .add_part(
            &slide_part,
            &format!("{ct}.slide+xml"),
            slide(&format!("{}</p:spTree></p:cSld></p:sld>", shape(2, "slide text"))).as_bytes(),
        )
        .unwrap();
    writer
        .add_part(&notes_part, &format!("{ct}.notesSlide+xml"), notes.as_bytes())
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "ppt/presentation.xml");
    writer.add_part_rel(&pres, rel_types::SLIDE, "slides/slide1.xml");
    writer.add_part_rel(&slide_part, rel_types::NOTES_SLIDE, "../notesSlides/notesSlide1.xml");
    writer.finish().unwrap().into_inner()
}

#[test]
fn test_pptx_notes_cut_off_mid_shape_keep_the_shapes_before_it() {
    let head = format!(
        r#"<p:notes xmlns:p="{PML}" xmlns:a="{DML}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="N"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>speaker note</a:t></a:r></a:p></p:txBody></p:sp>"#
    );
    let sound = open(
        pptx_with_notes(&format!("{head}</p:spTree></p:cSld></p:notes>")),
        DocumentFormat::Pptx,
    )
    .unwrap();
    assert!(sound.plain_text().contains("speaker note"), "{}", sound.plain_text());
    assert!(sound.to_ir().metadata.warnings.is_empty());
    let cut = open(
        pptx_with_notes(&format!(r#"{head}<p:sp><p:txBody><a:p><a:r a=""#)),
        DocumentFormat::Pptx,
    )
    .unwrap();
    assert!(cut.plain_text().contains("speaker note"), "{}", cut.plain_text());
    let warnings = cut.to_ir().metadata.warnings;
    assert!(warnings.iter().any(|w| w.contains("notesSlide1.xml")), "{warnings:?}");
}

fn pptx_with_layout(layout: &str) -> Vec<u8> {
    let mut writer = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let pres = PartName::new("/ppt/presentation.xml").unwrap();
    let slide_part = PartName::new("/ppt/slides/slide1.xml").unwrap();
    let layout_part = PartName::new("/ppt/slideLayouts/slideLayout1.xml").unwrap();
    let presentation = format!(
        r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#
    );
    let ct = "application/vnd.openxmlformats-officedocument.presentationml";
    writer
        .add_part(&pres, &format!("{ct}.presentation.main+xml"), presentation.as_bytes())
        .unwrap();
    let slide_xml = slide(&format!("{}</p:spTree></p:cSld></p:sld>", shape(2, "slide text")));
    writer
        .add_part(&slide_part, &format!("{ct}.slide+xml"), slide_xml.as_bytes())
        .unwrap();
    writer
        .add_part(&layout_part, &format!("{ct}.slideLayout+xml"), layout.as_bytes())
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "ppt/presentation.xml");
    writer.add_part_rel(&pres, rel_types::SLIDE, "slides/slide1.xml");
    writer.add_part_rel(&slide_part, rel_types::SLIDE_LAYOUT, "../slideLayouts/slideLayout1.xml");
    writer.finish().unwrap().into_inner()
}

#[test]
fn test_a_pptx_layout_cut_off_mid_shape_keeps_the_shapes_before_it() {
    let head = format!(
        r#"<p:sldLayout xmlns:p="{PML}" xmlns:a="{DML}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{}"#,
        shape(3, "static layout text")
    );
    let sound = open(
        pptx_with_layout(&format!("{head}</p:spTree></p:cSld></p:sldLayout>")),
        DocumentFormat::Pptx,
    )
    .unwrap();
    assert!(sound.plain_text().contains("static layout text"), "{}", sound.plain_text());
    assert!(sound.to_ir().metadata.warnings.is_empty());
    let cut = open(
        pptx_with_layout(&format!(r#"{head}<p:sp><p:txBody><a:p><a:r a=""#)),
        DocumentFormat::Pptx,
    )
    .unwrap();
    assert!(cut.plain_text().contains("static layout text"), "{}", cut.plain_text());
    let warnings = cut.to_ir().metadata.warnings;
    assert!(warnings.iter().any(|w| w.contains("slideLayout1.xml")), "{warnings:?}");
}

#[test]
fn test_a_pptx_slide_cut_inside_its_first_shape_is_not_an_empty_slide() {
    // The shape tree opens with its own header elements; they are not shapes.
    // Damage in the first real shape leaves nothing to keep, which is an
    // unreadable slide, never a complete slide with no shapes.
    let cut = "<p:sp><p:txBody><a:p><a:r><a:t>lost";
    let doc = open(pptx(&slide(cut)), DocumentFormat::Pptx);
    if let Ok(doc) = doc {
        let warnings = doc.to_ir().metadata.warnings;
        assert!(warnings.iter().any(|w| w.contains("slides/slide1.xml")), "{warnings:?}");
        assert!(
            !warnings.iter().any(|w| w.contains("malformed after")),
            "reported as kept shapes: {warnings:?}"
        );
    }
}

#[test]
fn test_a_docx_body_damaged_in_its_first_block_is_an_error_whatever_precedes_it() {
    // `w:background` sits beside `w:body` and has children of its own.
    let xml = format!(
        r#"<w:document xmlns:w="{WML}" xmlns:v="urn:schemas-microsoft-com:vml"><w:background w:color="FFFFFF"><v:fill/><v:fill/></w:background><w:body><w:p><w:r><w:t>cut"#
    );
    let mut writer = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    writer
        .add_part(
            &PartName::new("/word/document.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            xml.as_bytes(),
        )
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    let bytes = writer.finish().unwrap().into_inner();
    assert!(
        open(bytes, DocumentFormat::Docx).is_err(),
        "children of w:background were counted as body blocks"
    );
}
