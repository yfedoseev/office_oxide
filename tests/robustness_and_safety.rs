//! Malformed and adversarial input must fail loudly, not quietly.
//!
//! Every fixture here is built in code. The property under test is the same
//! throughout: a file we cannot read correctly must produce an error naming
//! what is wrong, never an `Ok` holding partial or empty content — and no
//! input may abort the process.

use std::io::{Cursor, Write};

use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::rel_types;
use office_oxide::{Document, DocumentFormat};

const CT_DOC: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_WB: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

fn docx_with(body: &str) -> Vec<u8> {
    let mut w = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let part = PartName::new("/word/document.xml").unwrap();
    w.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    let xml = format!(
        r#"<?xml version="1.0"?><w:document
             xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
           <w:body>{body}</w:body></w:document>"#
    );
    w.add_part(&part, CT_DOC, xml.as_bytes()).unwrap();
    w.finish().unwrap().into_inner()
}

fn open_docx(bytes: Vec<u8>) -> office_oxide::Result<Document> {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
}

/// Rewrite every occurrence of a ZIP entry name in-place. `from` and `to`
/// must be the same length so the surrounding headers stay valid.
fn rename_entry(bytes: &mut [u8], from: &[u8], to: &[u8]) {
    assert_eq!(from.len(), to.len());
    let mut i = 0;
    while i + from.len() <= bytes.len() {
        if &bytes[i..i + from.len()] == from {
            bytes[i..i + from.len()].copy_from_slice(to);
            i += from.len();
        } else {
            i += 1;
        }
    }
}

/// `Document` has no `Debug`, so `expect_err` is unavailable.
fn expect_err(r: office_oxide::Result<Document>, why: &str) -> office_oxide::OfficeError {
    match r {
        Ok(_) => panic!("{why}"),
        Err(e) => e,
    }
}

// ---------------------------------------------------------------------------
// #145 — truncated XML
// ---------------------------------------------------------------------------

#[test]
fn a_truncated_document_part_is_an_error_not_a_short_document() {
    // A file cut short by a failed download used to parse to whatever came
    // before the cut and report success, so a partial document looked
    // complete.
    let mut w = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let part = PartName::new("/word/document.xml").unwrap();
    w.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    let truncated = br#"<?xml version="1.0"?><w:document
         xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
       <w:body><w:p><w:r><w:t>TRUNC"#;
    w.add_part(&part, CT_DOC, truncated).unwrap();
    let bytes = w.finish().unwrap().into_inner();

    let err = expect_err(open_docx(bytes), "a truncated part must not parse to Ok");
    let msg = err.to_string();
    assert!(msg.contains("truncated"), "expected a truncation error, got {msg}");
}

#[test]
fn a_complete_document_still_parses() {
    let doc = open_docx(docx_with(r#"<w:p><w:r><w:t>OK</w:t></w:r></w:p>"#)).expect("parse");
    assert_eq!(doc.plain_text(), "OK");
}

// ---------------------------------------------------------------------------
// #145 — format mismatch
// ---------------------------------------------------------------------------

#[test]
fn opening_a_workbook_as_a_document_names_what_it_found() {
    let mut w = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let part = PartName::new("/xl/workbook.xml").unwrap();
    w.add_package_rel(rel_types::OFFICE_DOCUMENT, "xl/workbook.xml");
    w.add_part(
        &part,
        CT_WB,
        br#"<?xml version="1.0"?><workbook
             xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
           <sheets/></workbook>"#,
    )
    .unwrap();
    let bytes = w.finish().unwrap().into_inner();

    let err = expect_err(open_docx(bytes), "an XLSX opened as a DOCX must not report success");
    let msg = err.to_string();
    assert!(
        msg.contains("format mismatch") && msg.contains("spreadsheetml"),
        "error should name what was found, got {msg}"
    );
}

// ---------------------------------------------------------------------------
// #145 — duplicate part names
// ---------------------------------------------------------------------------

#[test]
fn duplicate_part_names_are_refused() {
    // Two entries with the same name mean two readers can see two different
    // documents from the same bytes. Whichever copy we picked would be
    // accidental, so the package is refused.
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts: zip::write::FileOptions<'_, ()> = zip::write::FileOptions::default();

    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(
        format!(
            r#"<?xml version="1.0"?><Types
                 xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
               <Override PartName="/word/document.xml" ContentType="{CT_DOC}"/>
             </Types>"#
        )
        .as_bytes(),
    )
    .unwrap();

    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
        format!(
            r#"<?xml version="1.0"?><Relationships
                 xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
               <Relationship Id="rId1" Type="{}" Target="word/document.xml"/>
             </Relationships>"#,
            rel_types::OFFICE_DOCUMENT
        )
        .as_bytes(),
    )
    .unwrap();

    // The zip crate refuses to write two entries with the same name, so the
    // second is written under a same-length placeholder and renamed
    // byte-for-byte afterwards. Entry-name lengths are unchanged, so every
    // header stays valid — which is exactly how such an archive is crafted
    // in the wild.
    for (name, text) in [
        ("word/document.xml", "FIRST"),
        ("word/dokument.xml", "SECOND"),
    ] {
        zip.start_file(name, opts).unwrap();
        zip.write_all(
            format!(
                r#"<?xml version="1.0"?><w:document
                     xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                   <w:body><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:body></w:document>"#
            )
            .as_bytes(),
        )
        .unwrap();
    }
    let mut bytes = zip.finish().unwrap().into_inner();
    rename_entry(&mut bytes, b"word/dokument.xml", b"word/document.xml");

    let err = expect_err(open_docx(bytes), "a package with duplicate parts must be refused");
    assert!(err.to_string().contains("duplicate part"), "got {err}");
}

// ---------------------------------------------------------------------------
// #151 — unbounded recursion
// ---------------------------------------------------------------------------

#[test]
fn deeply_nested_tables_do_not_abort_the_process() {
    // Nested `<w:tbl>` elements used to drive the recursive-descent parser
    // into a stack overflow, which aborts the process — an uncatchable
    // crash no caller in any binding can defend against.
    let depth = 5_000;
    let mut body = String::new();
    for _ in 0..depth {
        body.push_str("<w:tbl><w:tr><w:tc>");
    }
    body.push_str("<w:p><w:r><w:t>DEEP</w:t></w:r></w:p>");
    for _ in 0..depth {
        body.push_str("</w:tc></w:tr></w:tbl>");
    }
    // Either outcome is acceptable — a clean parse or a clean error. What
    // must not happen is a crash, which this test would fail to reach.
    let _ = open_docx(docx_with(&body));
}

#[test]
fn deeply_nested_tables_stay_within_the_depth_limit() {
    let depth = 200;
    let mut body = String::new();
    for _ in 0..depth {
        body.push_str("<w:tbl><w:tr><w:tc>");
    }
    body.push_str("<w:p><w:r><w:t>DEEP</w:t></w:r></w:p>");
    for _ in 0..depth {
        body.push_str("</w:tc></w:tr></w:tbl>");
    }
    let doc = open_docx(docx_with(&body)).expect("parse");
    // Rendering must also terminate; the IR tree is bounded by the same cap.
    let _ = doc.to_ir().to_markdown();
}

// ---------------------------------------------------------------------------
// #138 — malformed CFB header
// ---------------------------------------------------------------------------

#[test]
fn a_cfb_header_with_an_absurd_sector_shift_is_rejected_cleanly() {
    // `1usize << shift` overflows for any shift a malformed file cares to
    // write. The spec permits only 9 and 12.
    let mut data = vec![0u8; 1536];
    data[..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
    data[0x1A] = 3; // major version 3
    data[0x1B] = 0;
    data[0x1C] = 0xFE; // byte order
    data[0x1D] = 0xFF;
    data[0x1E] = 0xFF; // sector shift = 65535
    data[0x1F] = 0xFF;
    data[0x20] = 6; // mini sector shift
    data[0x21] = 0;

    // Any of the CFB-backed formats exercises the header parser.
    let err = expect_err(
        Document::from_reader(Cursor::new(data), DocumentFormat::Doc),
        "an invalid sector shift must be rejected",
    );
    // The point is that it is an error rather than a panic; the message is
    // whatever the reader chose.
    let _ = err.to_string();
}

// ---------------------------------------------------------------------------
// #158 — writers must not emit XML-illegal characters
// ---------------------------------------------------------------------------

#[test]
fn control_characters_do_not_reach_the_generated_xml() {
    use office_oxide::create;
    use office_oxide::ir::*;

    // U+0000..U+0008 and friends have no XML 1.0 representation at all —
    // not even as a numeric character reference. Emitting one produces a
    // file Word, Excel and LibreOffice all reject.
    let dirty = "before\u{0}\u{1}\u{8}\u{B}\u{C}\u{1F}after";
    let ir = DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Docx,
            title: Some(dirty.to_string()),
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(dirty))],
                ..Default::default()
            })],
            ..Default::default()
        }],
    };

    for fmt in [
        DocumentFormat::Docx,
        DocumentFormat::Xlsx,
        DocumentFormat::Pptx,
    ] {
        let mut buf = Cursor::new(Vec::new());
        create::create_from_ir_to_writer(&ir, fmt, &mut buf).expect("write");
        buf.set_position(0);
        // The generated package must re-open, and the round-tripped text
        // must keep the printable characters and none of the illegal ones.
        let doc = Document::from_reader(buf, fmt).expect("re-open generated file");
        let text = doc.plain_text();
        assert!(
            text.contains("before") && text.contains("after"),
            "{fmt:?}: lost the printable text: {text:?}"
        );
        assert!(
            !text.chars().any(|c| matches!(c,
                '\u{0}'..='\u{8}' | '\u{B}' | '\u{C}' | '\u{E}'..='\u{1F}')),
            "{fmt:?}: an XML-illegal character survived: {text:?}"
        );
    }
}

// ---------------------------------------------------------------------------
// #159 — replace_text must not corrupt the document
// ---------------------------------------------------------------------------

#[test]
fn replace_text_matches_and_writes_escaped_characters_correctly() {
    use office_oxide::edit::EditableDocument;

    // The stored XML holds `AT&amp;T`, so matching the raw bytes never
    // found `AT&T`; and writing `a & b` back injected a raw `&`.
    let bytes = docx_with(r#"<w:p><w:r><w:t>AT&amp;T is here</w:t></w:r></w:p>"#);
    let mut doc = EditableDocument::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .expect("open for editing");
    let n = doc.replace_text("AT&T", "M & S <Ltd>");
    assert_eq!(n, 1, "the decoded text must match");

    let mut out = Cursor::new(Vec::new());
    doc.write_to(&mut out).expect("save");
    out.set_position(0);
    let reopened = Document::from_reader(out, DocumentFormat::Docx)
        .expect("the edited document must still be readable");
    assert_eq!(reopened.plain_text(), "M & S <Ltd> is here");
}

#[test]
fn replace_text_does_not_rewrite_table_elements() {
    use office_oxide::edit::EditableDocument;

    // `<w:t` is a prefix of `<w:tbl>`, `<w:tc>`, `<w:tr>` and `<w:tab/>`.
    let bytes = docx_with(
        r#"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    let mut doc = EditableDocument::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .expect("open for editing");
    assert_eq!(doc.replace_text("cell", "CELL"), 1);

    let mut out = Cursor::new(Vec::new());
    doc.write_to(&mut out).expect("save");
    out.set_position(0);
    let reopened = Document::from_reader(out, DocumentFormat::Docx).expect("still readable");
    assert!(reopened.plain_text().contains("CELL"));
}

// ---------------------------------------------------------------------------
// #176 — a malformed theme must not make the document unreadable
// ---------------------------------------------------------------------------

#[test]
fn a_malformed_theme_part_does_not_fail_the_open() {
    // A *missing* theme was always fine; a malformed one failed the whole
    // open, which is the inconsistency that gave the bug away.
    let mut w = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let part = PartName::new("/word/document.xml").unwrap();
    w.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    w.add_part(
        &part,
        CT_DOC,
        br#"<?xml version="1.0"?><w:document
             xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
           <w:body><w:p><w:r><w:t>BODY</w:t></w:r></w:p></w:body></w:document>"#,
    )
    .unwrap();
    let theme = PartName::new("/word/theme/theme1.xml").unwrap();
    w.add_part(
        &theme,
        "application/vnd.openxmlformats-officedocument.theme+xml",
        b"<<< not xml at all >>>",
    )
    .unwrap();
    w.add_part_rel(&part, rel_types::THEME, "theme/theme1.xml");
    let bytes = w.finish().unwrap().into_inner();

    let doc = open_docx(bytes).expect("a bad theme must not fail the open");
    assert_eq!(doc.plain_text(), "BODY");
}

// ---------------------------------------------------------------------------
// #135 / #136 — heading level normalisation
// ---------------------------------------------------------------------------

#[test]
fn outline_level_nine_is_body_text_not_a_heading() {
    use office_oxide::ir::Element;

    // ECMA-376 §17.3.1.20: 9 means "no outline level applied", and is the
    // value assumed when the element is absent. Word writes it for
    // "Outline level: Body Text" and for the built-in TOCHeading style.
    let doc = open_docx(docx_with(
        r#"<w:p><w:pPr><w:outlineLvl w:val="9"/></w:pPr>
             <w:r><w:t>body text</w:t></w:r></w:p>"#,
    ))
    .expect("parse");
    assert!(
        matches!(doc.to_ir().sections[0].elements[0], Element::Paragraph(_)),
        "outlineLvl=9 must not become a heading"
    );
    assert!(
        !doc.to_markdown().starts_with('#'),
        "markdown emitted a heading: {:?}",
        doc.to_markdown()
    );
}

#[test]
fn every_renderer_agrees_on_an_out_of_range_heading_level() {
    use office_oxide::ir::*;

    // `Heading` derives Default and serde accepts an explicit 0, so level 0
    // is reachable. The markdown renderer used to clamp with `min(6)` — no
    // lower bound — and produced a body line while HTML produced `<h1>`.
    let ir = DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![Element::Heading(Heading {
                level: 0,
                content: vec![InlineContent::Text(TextSpan::plain("Zero"))],
                ..Default::default()
            })],
            ..Default::default()
        }],
    };
    assert!(ir.to_markdown().starts_with("# Zero"), "markdown: {:?}", ir.to_markdown());
    assert!(ir.to_html().contains("<h1>Zero</h1>"), "html: {}", ir.to_html());
}

// ---------------------------------------------------------------------------
// #137 — the section title must not be rendered twice
// ---------------------------------------------------------------------------

#[test]
fn a_sections_first_heading_is_not_rendered_twice() {
    let doc = open_docx(docx_with(
        r#"<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr>
             <w:r><w:t>Introduction</w:t></w:r></w:p>
           <w:p><w:r><w:t>Body.</w:t></w:r></w:p>"#,
    ))
    .expect("parse");
    let md = doc.to_ir().to_markdown();
    assert_eq!(
        md.matches("Introduction").count(),
        1,
        "the section title duplicated the heading: {md}"
    );
    assert!(md.contains("# Introduction"));
}
