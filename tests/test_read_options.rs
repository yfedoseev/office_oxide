//! Hidden text and tracked deletions are left out by default and come back
//! when the reader is asked for them, in `.docx` and `.doc` alike.

mod common;

use std::io::Cursor;

use common::{FibTweaks, LimitScope, Para, Subdocs, build_doc_full, default_limits, open_doc};
use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::rel_types;
use office_oxide::read_options::{
    include_deleted_text, include_hidden_text, set_include_deleted_text, set_include_hidden_text,
};
use office_oxide::{Document, DocumentFormat};

fn docx(body: &str) -> Document {
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#
    );
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
    let bytes = writer.finish().unwrap().into_inner();
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx).unwrap()
}

const MIXED: &str = r#"<w:p>
<w:r><w:t>kept </w:t></w:r>
<w:r><w:rPr><w:vanish/></w:rPr><w:t>secret </w:t></w:r>
<w:del w:id="1" w:author="a" w:date="2020-01-01T00:00:00Z"><w:r><w:delText>removed </w:delText></w:r></w:del>
<w:ins w:id="2" w:author="a" w:date="2020-01-01T00:00:00Z"><w:r><w:t>added </w:t></w:r></w:ins>
<w:moveFrom w:id="3" w:author="a" w:date="2020-01-01T00:00:00Z"><w:r><w:t>moved </w:t></w:r></w:moveFrom>
<w:r><w:t>end</w:t></w:r>
</w:p>"#;

fn words(doc: &Document) -> Vec<String> {
    doc.plain_text()
        .split_whitespace()
        .map(str::to_string)
        .collect()
}

#[test]
fn test_docx_hidden_and_deleted_text_are_left_out_by_default() {
    let _defaults = default_limits();
    assert!(!include_hidden_text() && !include_deleted_text());
    assert_eq!(words(&docx(MIXED)), ["kept", "added", "end"]);
}

#[test]
fn test_docx_hidden_text_comes_back_when_asked_for() {
    let mut scope = LimitScope::new();
    scope.set(include_hidden_text, set_include_hidden_text, true);
    let doc = docx(MIXED);
    assert_eq!(words(&doc), ["kept", "secret", "added", "end"]);
    assert!(doc.to_markdown().contains("secret"));
    assert!(doc.to_html().contains("secret"));
}

#[test]
fn test_docx_deleted_text_comes_back_when_asked_for() {
    let mut scope = LimitScope::new();
    scope.set(include_deleted_text, set_include_deleted_text, true);
    let doc = docx(MIXED);
    // `moveFrom` stays out: the move's destination carries that text.
    assert_eq!(words(&doc), ["kept", "removed", "added", "end"]);
    assert!(doc.to_markdown().contains("removed"));
    assert!(doc.to_html().contains("removed"));
}

#[test]
fn test_docx_both_options_together() {
    let mut scope = LimitScope::new();
    scope.set(include_hidden_text, set_include_hidden_text, true);
    scope.set(include_deleted_text, set_include_deleted_text, true);
    assert_eq!(words(&docx(MIXED)), ["kept", "secret", "removed", "added", "end"]);
}

#[test]
fn test_docx_a_deleted_table_cell_follows_the_option() {
    let table = r#"<w:tbl><w:tr>
<w:tc><w:p><w:r><w:t>one</w:t></w:r></w:p></w:tc>
<w:tc><w:tcPr><w:cellDel w:id="1" w:author="a" w:date="2020-01-01T00:00:00Z"/></w:tcPr><w:p><w:r><w:t>gone</w:t></w:r></w:p></w:tc>
</w:tr></w:tbl>"#;
    {
        let _defaults = default_limits();
        assert!(!docx(table).to_markdown().contains("gone"));
    }
    let mut scope = LimitScope::new();
    scope.set(include_deleted_text, set_include_deleted_text, true);
    let doc = docx(table);
    assert!(doc.to_markdown().contains("gone"));
    assert!(doc.plain_text().contains("gone"));
}

/// Four paragraphs of 99 `x`; 300 one-character runs over the start of the
/// text carry the character property `grpprl`.
fn doc_with_marked_runs(grpprl: Vec<u8>) -> Document {
    let paras: Vec<Para> = (0..4)
        .map(|_| Para {
            text: Box::leak("x".repeat(99).into_boxed_str()),
            terminator: '\r',
            grpprl: Vec::new(),
        })
        .collect();
    let tweaks = FibTweaks {
        chpx_runs: 300,
        chpx_grpprl: grpprl,
        ..Default::default()
    };
    open_doc(&build_doc_full(&paras, &Subdocs::default(), tweaks))
}

fn xs(doc: &Document) -> usize {
    doc.plain_text().matches('x').count()
}

#[test]
fn test_doc_hidden_text_follows_the_option() {
    let hidden = vec![0x3C, 0x08, 0x01]; // sprmCFVanish = 1
    {
        let _defaults = default_limits();
        assert!(xs(&doc_with_marked_runs(hidden.clone())) < 4 * 99);
    }
    let mut scope = LimitScope::new();
    scope.set(include_hidden_text, set_include_hidden_text, true);
    let doc = doc_with_marked_runs(hidden);
    assert_eq!(xs(&doc), 4 * 99);
    assert_eq!(doc.to_markdown().matches('x').count(), 4 * 99);
}

#[test]
fn test_doc_deleted_text_follows_the_option() {
    let deleted = vec![0x00, 0x08, 0x01]; // sprmCFRMarkDel = 1
    {
        let _defaults = default_limits();
        assert!(xs(&doc_with_marked_runs(deleted.clone())) < 4 * 99);
    }
    let mut scope = LimitScope::new();
    scope.set(include_deleted_text, set_include_deleted_text, true);
    assert_eq!(xs(&doc_with_marked_runs(deleted.clone())), 4 * 99);
    // Hidden text is a separate choice: this document only deletes.
    scope.set(include_hidden_text, set_include_hidden_text, false);
    assert_eq!(xs(&doc_with_marked_runs(deleted)), 4 * 99);
}
