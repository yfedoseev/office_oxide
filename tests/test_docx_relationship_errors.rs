//! A relationships part that cannot be read leaves the part it belongs to
//! without its images and links. The part is still used, but the loss is
//! reported, naming the relationships part and the reason, as every other
//! unreadable part is.

use std::io::{Cursor, Write};

use office_oxide::docx::DocxDocument;
use office_oxide::{Document, DocumentFormat};

const WML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

/// A target past the limit on a relationship's target (64 KiB).
fn oversized_target() -> String {
    format!("https://example.com/{}", "a".repeat(70_000))
}

fn rels(entries: &str) -> String {
    format!(r#"<Relationships xmlns="{PKG_REL}">{entries}</Relationships>"#)
}

/// A `.docx` whose header, footnotes part and main relationships are as
/// given. `header_rels` and `notes_rels` are the entries of their parts'
/// relationships (`None`: no such part).
fn docx(main_rels: &str, header_rels: Option<String>, notes_rels: Option<String>) -> Vec<u8> {
    let mut parts: Vec<(String, String)> = vec![
        ("[Content_Types].xml".into(), r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>"#.into()),
        ("_rels/.rels".into(), rels(&format!(r#"<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/>"#))),
        ("word/_rels/document.xml.rels".into(), rels(main_rels)),
        ("word/document.xml".into(), format!(r#"<w:document xmlns:w="{WML}" xmlns:r="{REL}"><w:body><w:p><w:r><w:t>Body survives</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId2"/></w:sectPr></w:body></w:document>"#)),
        ("word/header1.xml".into(), format!(r#"<w:hdr xmlns:w="{WML}" xmlns:r="{REL}"><w:p><w:hyperlink r:id="rId1"><w:r><w:t>HEADERLINK</w:t></w:r></w:hyperlink></w:p></w:hdr>"#)),
        ("word/footnotes.xml".into(), format!(r#"<w:footnotes xmlns:w="{WML}" xmlns:r="{REL}"><w:footnote w:id="1"><w:p><w:hyperlink r:id="rId1"><w:r><w:t>NOTELINK</w:t></w:r></w:hyperlink></w:p></w:footnote></w:footnotes>"#)),
    ];
    if let Some(h) = header_rels {
        parts.push(("word/_rels/header1.xml.rels".into(), rels(&h)));
    }
    if let Some(n) = notes_rels {
        parts.push(("word/_rels/footnotes.xml.rels".into(), rels(&n)));
    }
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = zip::write::SimpleFileOptions::default();
    for (name, body) in parts {
        zip.start_file(name, opts).unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn main_rels() -> String {
    format!(
        r#"<Relationship Id="rId2" Type="{REL}/header" Target="header1.xml"/><Relationship Id="rId3" Type="{REL}/footnotes" Target="footnotes.xml"/>"#
    )
}

fn bad_link() -> String {
    format!(
        r#"<Relationship Id="rId1" Type="{REL}/hyperlink" Target="{}" TargetMode="External"/>"#,
        oversized_target()
    )
}

#[test]
fn test_unreadable_header_and_note_relationships_are_reported() {
    let bytes = docx(&main_rels(), Some(bad_link()), Some(bad_link()));
    let doc = DocxDocument::from_reader(Cursor::new(bytes.clone()))
        .expect("an unreadable relationships part does not fail the document");
    assert!(doc.plain_text().contains("Body survives"));
    for part in [
        "word/_rels/header1.xml.rels",
        "word/_rels/footnotes.xml.rels",
    ] {
        assert!(
            doc.unreadable_parts
                .iter()
                .any(|p| p.contains(part) && p.contains("limit")),
            "{part} not reported: {:?}",
            doc.unreadable_parts
        );
    }
    let ir = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    assert!(ir.metadata.text_truncated);
    for part in [
        "word/_rels/header1.xml.rels",
        "word/_rels/footnotes.xml.rels",
    ] {
        assert!(
            ir.metadata
                .warnings
                .iter()
                .any(|w| w.contains("skipped unreadable part") && w.contains(part)),
            "{part} not in the warnings: {:?}",
            ir.metadata.warnings
        );
    }
}

#[test]
fn test_readable_header_and_note_relationships_report_nothing() {
    let ok = format!(
        r#"<Relationship Id="rId1" Type="{REL}/hyperlink" Target="https://example.com/" TargetMode="External"/>"#
    );
    let bytes = docx(&main_rels(), Some(ok.clone()), Some(ok));
    let doc = DocxDocument::from_reader(Cursor::new(bytes)).unwrap();
    assert!(doc.unreadable_parts.is_empty(), "{:?}", doc.unreadable_parts);
}

/// The main document's relationships name every other part, so a package
/// whose own cannot be read has no document: refused, as for a presentation.
#[test]
fn test_unreadable_main_relationships_refuse_the_package() {
    let bytes = docx(&bad_link(), None, None);
    assert!(DocxDocument::from_reader(Cursor::new(bytes)).is_err());
}
