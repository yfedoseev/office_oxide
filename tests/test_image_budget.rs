//! Image bytes copied into an IR: many pictures can share one image part,
//! and every picture's `Image` owns a copy. The copies are charged to a
//! per-document budget. These tests lower the process-wide budgets, so they
//! have a test binary to themselves, and each sets the same values.

use std::io::{Cursor, Write};

use office_oxide::ir::{Element, Section};
use office_oxide::limits::{set_max_image_bytes, set_max_text_chars};
use office_oxide::{Document, DocumentFormat};

const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const WML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

const IMAGE_LEN: usize = 4_000;
/// Two copies of the image fit.
const IMAGE_BUDGET: u64 = 9_000;
/// One serialised copy of a header holding the image fits (about 5.4 KB of
/// base64 plus structure), a second does not.
const TEXT_BUDGET: usize = 8_000;

fn limits() {
    set_max_image_bytes(IMAGE_BUDGET);
    set_max_text_chars(TEXT_BUDGET);
}

fn zip_of(parts: Vec<(String, Vec<u8>)>) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = zip::write::SimpleFileOptions::default();
    for (name, data) in parts {
        zip.start_file(name, opts).unwrap();
        zip.write_all(&data).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn rels(entries: &[(&str, &str, &str)]) -> String {
    let mut s = format!(r#"<Relationships xmlns="{PKG_REL}">"#);
    for (id, ty, target) in entries {
        s.push_str(&format!(r#"<Relationship Id="{id}" Type="{REL}/{ty}" Target="{target}"/>"#));
    }
    s.push_str("</Relationships>");
    s
}

fn image() -> Vec<u8> {
    vec![b'P'; IMAGE_LEN]
}

fn drawing(rid: &str) -> String {
    format!(
        r#"<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="p"/><a:graphic xmlns:a="{DML}"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="{rid}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#
    )
}

const DOCX_CT: &str = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#;

fn docx(body: &str, header: Option<&str>) -> Vec<u8> {
    let mut doc_rels = vec![("rId5", "image", "media/img.png")];
    if header.is_some() {
        doc_rels.push(("rId6", "header", "header1.xml"));
    }
    let mut parts = vec![
        ("[Content_Types].xml".to_string(), DOCX_CT.as_bytes().to_vec()),
        (
            "_rels/.rels".to_string(),
            rels(&[("rId1", "officeDocument", "word/document.xml")]).into_bytes(),
        ),
        (
            "word/document.xml".to_string(),
            format!(r#"<w:document xmlns:w="{WML}" xmlns:r="{REL}"><w:body>{body}</w:body></w:document>"#)
                .into_bytes(),
        ),
        ("word/_rels/document.xml.rels".to_string(), rels(&doc_rels).into_bytes()),
        ("word/media/img.png".to_string(), image()),
    ];
    if let Some(h) = header {
        parts.push((
            "word/header1.xml".to_string(),
            format!(r#"<w:hdr xmlns:w="{WML}" xmlns:r="{REL}">{h}</w:hdr>"#).into_bytes(),
        ));
        parts.push((
            "word/_rels/header1.xml.rels".to_string(),
            rels(&[("rId1", "image", "media/img.png")]).into_bytes(),
        ));
    }
    zip_of(parts)
}

fn image_copies(sections: &[Section]) -> usize {
    fn walk(elements: &[Element]) -> usize {
        elements
            .iter()
            .map(|e| match e {
                Element::Image(i) => usize::from(i.data.is_some()),
                _ => 0,
            })
            .sum()
    }
    sections.iter().map(|s| walk(&s.elements)).sum()
}

#[test]
fn test_docx_pictures_sharing_one_image_stop_copying_at_the_budget() {
    limits();
    let body: String = (0..20)
        .map(|_| format!("<w:p>{}</w:p>", drawing("rId5")))
        .collect();
    let bytes = docx(&body, None);
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx).unwrap();
    let ir = doc.to_ir();
    assert_eq!(image_copies(&ir.sections), 2, "two copies fit the budget");
    // The pictures keep their place without their bytes, and the document says so.
    let pictures: usize = ir
        .sections
        .iter()
        .flat_map(|s| &s.elements)
        .filter(|e| matches!(e, Element::Image(_)))
        .count();
    assert_eq!(pictures, 20);
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("image bytes")),
        "{:?}",
        ir.metadata.warnings
    );
}

#[test]
fn test_docx_header_with_a_large_image_shared_by_many_sections_stays_bounded() {
    limits();
    let sections: String = (0..30)
        .map(|i| {
            format!(
                r#"<w:p><w:r><w:t>s{i}</w:t></w:r><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId6"/></w:sectPr></w:pPr></w:p>"#
            )
        })
        .collect();
    let bytes = docx(&sections, Some(&format!("<w:p>{}</w:p>", drawing("rId1"))));
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx).unwrap();
    let ir = doc.to_ir();
    let with_header = ir.sections.iter().filter(|s| s.header.is_some()).count();
    assert_eq!(with_header, 1, "only the copy that fits the budget is made");
    assert!(ir.metadata.text_truncated);
}

#[test]
fn test_pptx_pictures_sharing_one_image_stop_copying_at_the_budget() {
    limits();
    let pic = r#"<p:pic><p:nvPicPr><p:cNvPr id="3" name="P"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId1"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:pic>"#;
    let pics = pic.repeat(20);
    let slide = format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{pics}</p:spTree></p:cSld></p:sld>"#
    );
    let ct = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#;
    let bytes = zip_of(vec![
        ("[Content_Types].xml".to_string(), ct.as_bytes().to_vec()),
        (
            "_rels/.rels".to_string(),
            rels(&[("rId1", "officeDocument", "ppt/presentation.xml")]).into_bytes(),
        ),
        (
            "ppt/presentation.xml".to_string(),
            format!(r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#)
                .into_bytes(),
        ),
        (
            "ppt/_rels/presentation.xml.rels".to_string(),
            rels(&[("rId2", "slide", "slides/slide1.xml")]).into_bytes(),
        ),
        ("ppt/slides/slide1.xml".to_string(), slide.into_bytes()),
        (
            "ppt/slides/_rels/slide1.xml.rels".to_string(),
            rels(&[("rId1", "image", "../media/img.png")]).into_bytes(),
        ),
        ("ppt/media/img.png".to_string(), image()),
    ]);
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Pptx).unwrap();
    // The reader keeps the bytes once: every picture holds the same allocation.
    let p = doc.as_pptx().expect("a pptx");
    let office_oxide::pptx::Shape::Picture(first) = &p.slides[0].shapes[0] else {
        panic!("a picture")
    };
    let first = first.data.as_ref().unwrap();
    for shape in &p.slides[0].shapes {
        let office_oxide::pptx::Shape::Picture(pic) = shape else {
            panic!("a picture")
        };
        assert!(std::sync::Arc::ptr_eq(pic.data.as_ref().unwrap(), first));
    }
    let ir = doc.to_ir();
    fn count(elements: &[Element]) -> usize {
        elements
            .iter()
            .map(|e| match e {
                Element::Image(i) => usize::from(i.data.is_some()),
                Element::TextBox(t) => count(&t.content),
                _ => 0,
            })
            .sum()
    }
    let copies: usize = ir.sections.iter().map(|s| count(&s.elements)).sum();
    assert_eq!(copies, 2, "two copies fit the budget");
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("image bytes"))
    );
}
