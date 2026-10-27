//! Shared parts that many references point at: the reader must keep one
//! copy of the part, not one per reference.
//!
//! Every fixture is a minimal synthetic package built in code. None of
//! these tests touches the process-wide limits, so they may run in parallel.

use std::io::{Cursor, Write};

use office_oxide::pptx::PptxDocument;
use office_oxide::{Document, DocumentFormat};

const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

fn zip_of(parts: Vec<(String, String)>) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = zip::write::SimpleFileOptions::default();
    for (name, data) in parts {
        zip.start_file(name, opts).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
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

fn sp(text: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
    )
}

fn slide_xml(shapes: &str) -> String {
    format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shapes}</p:spTree></p:cSld></p:sld>"#
    )
}

/// A deck whose slide list is `sld_ids` (each entry `(id, rId)`), with the
/// `slides` slide parts `slide1.xml..` wired to `rId{n+1}`, the optional
/// per-slide relationships, and any extra parts.
fn pptx(
    sld_ids: &[(u32, String)],
    slides: &[String],
    slide_rels: &[Option<String>],
    extra: Vec<(String, String)>,
) -> Vec<u8> {
    let mut ct = String::from(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>"#,
    );
    let mut pres_rels: Vec<(String, String)> = Vec::new();
    let mut parts: Vec<(String, String)> = Vec::new();
    for (i, s) in slides.iter().enumerate() {
        let n = i + 1;
        ct.push_str(&format!(r#"<Override PartName="/ppt/slides/slide{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#));
        pres_rels.push((format!("rId{}", n + 1), format!("slides/slide{n}.xml")));
        parts.push((format!("ppt/slides/slide{n}.xml"), s.clone()));
        if let Some(Some(r)) = slide_rels.get(i) {
            parts.push((format!("ppt/slides/_rels/slide{n}.xml.rels"), r.clone()));
        }
    }
    ct.push_str(r#"<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/></Types>"#);
    let mut list = String::new();
    for (id, r) in sld_ids {
        list.push_str(&format!(r#"<p:sldId id="{id}" r:id="{r}"/>"#));
    }
    let pr: Vec<(&str, &str, &str)> = pres_rels
        .iter()
        .map(|(a, c)| (a.as_str(), "slide", c.as_str()))
        .collect();
    let mut all = vec![
        ("[Content_Types].xml".to_string(), ct),
        (
            "_rels/.rels".to_string(),
            rels(&[("rId1", "officeDocument", "ppt/presentation.xml")]),
        ),
        (
            "ppt/presentation.xml".to_string(),
            format!(
                r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst>{list}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#
            ),
        ),
        ("ppt/_rels/presentation.xml.rels".to_string(), rels(&pr)),
    ];
    all.extend(parts);
    all.extend(extra);
    zip_of(all)
}

// ---------------------------------------------------------------------------
// A slide list that names one slide part many times
// ---------------------------------------------------------------------------

#[test]
fn test_repeated_slide_list_entry_keeps_one_slide() {
    let big = "w".repeat(2_000);
    let ids: Vec<(u32, String)> = (0..200).map(|i| (256 + i, "rId2".to_string())).collect();
    let bytes = pptx(&ids, &[slide_xml(&sp(&big))], &[], vec![]);
    let doc = PptxDocument::from_reader(Cursor::new(bytes.clone())).unwrap();
    assert_eq!(doc.slides.len(), 1, "one part is one slide, however often it is listed");
    // The repetition is reported, not silently swallowed.
    assert!(
        doc.unreadable_parts
            .iter()
            .any(|(p, e)| p.ends_with("slide1.xml") && e.contains("more than once")),
        "{:?}",
        doc.unreadable_parts
    );
    let md = Document::from_reader(Cursor::new(bytes), DocumentFormat::Pptx)
        .unwrap()
        .to_markdown();
    assert_eq!(md.matches(&big).count(), 1);
}

#[test]
fn test_distinct_slides_listed_in_order_are_all_kept() {
    let ids = vec![(256, "rId2".to_string()), (257, "rId3".to_string())];
    let bytes = pptx(&ids, &[slide_xml(&sp("first")), slide_xml(&sp("second"))], &[], vec![]);
    let doc = PptxDocument::from_reader(Cursor::new(bytes)).unwrap();
    assert_eq!(doc.slides.len(), 2);
    assert!(doc.unreadable_parts.is_empty());
}

// ---------------------------------------------------------------------------
// DOCX: one header part referenced from every section
// ---------------------------------------------------------------------------

const WML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// A document of `sections` sections, each `w:sectPr` referencing the
/// default header `rId1` and the default footer `rId2` (one part each).
fn docx_with_shared_header(sections: usize, header_text: &str) -> Vec<u8> {
    let mut body = String::new();
    for i in 0..sections {
        body.push_str(&format!(
            r#"<w:p><w:r><w:t>s{i}</w:t></w:r><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:footerReference w:type="default" r:id="rId2"/></w:sectPr></w:pPr></w:p>"#
        ));
    }
    let parts = vec![
        ("[Content_Types].xml".to_string(), r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>"#.to_string()),
        ("_rels/.rels".to_string(), rels(&[("rId1", "officeDocument", "word/document.xml")])),
        (
            "word/document.xml".to_string(),
            format!(r#"<w:document xmlns:w="{WML}" xmlns:r="{REL}"><w:body>{body}</w:body></w:document>"#),
        ),
        (
            "word/_rels/document.xml.rels".to_string(),
            rels(&[("rId1", "header", "header1.xml"), ("rId2", "footer", "footer1.xml")]),
        ),
        (
            "word/header1.xml".to_string(),
            format!(r#"<w:hdr xmlns:w="{WML}"><w:p><w:r><w:t>{header_text}</w:t></w:r></w:p></w:hdr>"#),
        ),
        (
            "word/footer1.xml".to_string(),
            format!(r#"<w:ftr xmlns:w="{WML}"><w:p><w:r><w:t>foot</w:t></w:r></w:p></w:ftr>"#),
        ),
    ];
    zip_of(parts)
}

#[test]
fn test_header_shared_by_many_sections_is_stored_once() {
    let hdr = "h".repeat(1_500);
    let bytes = docx_with_shared_header(100, &hdr);
    let doc = office_oxide::docx::DocxDocument::from_reader(Cursor::new(bytes.clone())).unwrap();
    assert_eq!(doc.sections.len(), 100);
    assert_eq!(
        doc.headers_footers.len(),
        2,
        "one header part and one footer part, however many sections reference them"
    );
    assert_eq!(doc.plain_text().matches(&hdr).count(), 1);
    // Every section still shows the shared header and footer in the IR.
    let ir = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    assert_eq!(ir.sections.len(), 100);
    assert!(
        ir.sections
            .iter()
            .all(|s| s.header.is_some() && s.footer.is_some())
    );
}
