//! A value stored once in the package and copied into every list, run or
//! paragraph that uses it: a numbering level's `w:lvlText`, a style's font
//! name, a style's tab stops. The reader keeps one copy; every copy that
//! lands in an IR is charged to the text budget first, and the conversion
//! says so when the budget refuses one.
//!
//! These tests lower the process-wide text budget, so they have a test binary
//! to themselves, and each sets the same value.

use std::io::{Cursor, Write};

use office_oxide::ir::{DocumentIR, Element};
use office_oxide::limits::set_max_text_chars;
use office_oxide::{Document, DocumentFormat};

const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;
const PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// The length of the value stored once: a large but storable string.
const BIG: usize = 50_000;
/// About six copies of `BIG` fit.
const TEXT_BUDGET: usize = 300_000;
/// How many users of the shared value the document has.
const USERS: usize = 40;

fn limits() {
    set_max_text_chars(TEXT_BUDGET);
}

fn zip_of(parts: Vec<(String, String)>) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = zip::write::SimpleFileOptions::default();
    for (name, data) in parts {
        zip.start_file(name, opts).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// A `.docx` with `body`, plus the optional `styles.xml` and `numbering.xml`.
fn docx(body: &str, styles: Option<&str>, numbering: Option<&str>) -> Vec<u8> {
    let mut ct = String::new();
    let mut rels = String::new();
    let mut parts = vec![];
    if let Some(s) = styles {
        ct.push_str(r#"<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>"#);
        rels.push_str(&format!(
            r#"<Relationship Id="rId3" Type="{REL}/styles" Target="styles.xml"/>"#
        ));
        parts.push(("word/styles.xml".to_string(), s.to_string()));
    }
    if let Some(n) = numbering {
        ct.push_str(r#"<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>"#);
        rels.push_str(&format!(
            r#"<Relationship Id="rId4" Type="{REL}/numbering" Target="numbering.xml"/>"#
        ));
        parts.push(("word/numbering.xml".to_string(), n.to_string()));
    }
    let mut all = vec![
        (
            "[Content_Types].xml".to_string(),
            format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>{ct}</Types>"#
            ),
        ),
        (
            "_rels/.rels".to_string(),
            format!(
                r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>"#
            ),
        ),
        (
            "word/_rels/document.xml.rels".to_string(),
            format!(r#"<Relationships xmlns="{PKG_REL}">{rels}</Relationships>"#),
        ),
        (
            "word/document.xml".to_string(),
            format!(r#"<w:document {NS}><w:body>{body}</w:body></w:document>"#),
        ),
    ];
    all.extend(parts);
    zip_of(all)
}

fn to_ir(bytes: Vec<u8>) -> DocumentIR {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .unwrap()
        .to_ir()
}

fn assert_reported(ir: &DocumentIR) {
    assert!(ir.metadata.text_truncated, "a refused copy must set the truncation flag");
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("set_max_text_chars")),
        "a refused copy must be reported: {:?}",
        ir.metadata.warnings
    );
}

fn lists(ir: &DocumentIR) -> Vec<&office_oxide::ir::List> {
    ir.sections
        .iter()
        .flat_map(|s| s.elements.iter())
        .filter_map(|e| match e {
            Element::List(l) => Some(l),
            _ => None,
        })
        .collect()
}

fn numbering(level_text: &str) -> String {
    format!(
        r#"<w:numbering {NS}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="{level_text}"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
    )
}

/// `lists` one-item lists, each ended by an ordinary paragraph.
fn list_body(lists: usize) -> String {
    r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p><w:p><w:r><w:t>y</w:t></w:r></w:p>"#
        .repeat(lists)
}

#[test]
fn test_marker_pattern_copies_per_list_are_charged() {
    limits();
    let num = numbering(&format!("{}%1.", "x".repeat(BIG)));
    let ir = to_ir(docx(&list_body(USERS), None, Some(&num)));
    let copied: usize = lists(&ir)
        .iter()
        .map(|l| l.marker_pattern.as_deref().map_or(0, str::len))
        .sum();
    assert!(copied > 0, "the budget holds a few copies");
    assert!(
        copied <= TEXT_BUDGET,
        "{copied} characters of pattern were copied against a budget of {TEXT_BUDGET}"
    );
    assert_eq!(lists(&ir).len(), USERS, "the lists themselves stay");
    assert_reported(&ir);
}

#[test]
fn test_small_marker_pattern_is_kept_on_every_list() {
    limits();
    let ir = to_ir(docx(&list_body(USERS), None, Some(&numbering("(%1)"))));
    assert!(
        lists(&ir)
            .iter()
            .all(|l| l.marker_pattern.as_deref() == Some("(%1)"))
    );
    assert!(!ir.metadata.text_truncated);
}
