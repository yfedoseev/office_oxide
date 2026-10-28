//! A relationship target is stored once but named by every hyperlink that
//! refers to it. The reader keeps one copy; the copies that land in an IR or in
//! Markdown are charged to the text budget; and a target too long to be a
//! genuine address is refused outright.
//!
//! These tests lower the process-wide text budget, so they have a test binary
//! to themselves, and each sets the same value.

use std::io::{Cursor, Write};

use office_oxide::docx::{DocxDocument, HyperlinkTarget, ParagraphContent};
use office_oxide::ir::{DocumentIR, Element, InlineContent};
use office_oxide::limits::set_max_text_chars;
use office_oxide::{Document, DocumentFormat};

const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const WML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// The length of the shared target: a long but admissible address.
const TARGET_LEN: usize = 50_000;
/// About six copies of the target fit.
const TEXT_BUDGET: usize = 300_000;
const LINKS: usize = 40;

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

/// A `.rels` part: `(id, type, target, external)` entries.
fn rels(entries: &[(&str, &str, &str, bool)]) -> String {
    let mut s = format!(r#"<Relationships xmlns="{PKG_REL}">"#);
    for (id, ty, target, external) in entries {
        let mode = if *external {
            r#" TargetMode="External""#
        } else {
            ""
        };
        s.push_str(&format!(
            r#"<Relationship Id="{id}" Type="{REL}/{ty}" Target="{target}"{mode}/>"#
        ));
    }
    s.push_str("</Relationships>");
    s
}

fn url(len: usize) -> String {
    format!("https://example.com/{}", "a".repeat(len - 20))
}

fn docx(target: &str, links: usize, fragments: bool) -> Vec<u8> {
    let mut body = String::new();
    for i in 0..links {
        let anchor = if fragments {
            format!(r#" w:anchor="f{i}""#)
        } else {
            String::new()
        };
        body.push_str(&format!(
            r#"<w:p><w:hyperlink r:id="rId1"{anchor}><w:r><w:t>link{i}</w:t></w:r></w:hyperlink></w:p>"#
        ));
    }
    zip_of(vec![
        ("[Content_Types].xml".to_string(), r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_string()),
        ("_rels/.rels".to_string(), rels(&[("rId1", "officeDocument", "word/document.xml", false)])),
        (
            "word/document.xml".to_string(),
            format!(r#"<w:document xmlns:w="{WML}" xmlns:r="{REL}"><w:body>{body}</w:body></w:document>"#),
        ),
        (
            "word/_rels/document.xml.rels".to_string(),
            rels(&[("rId1", "hyperlink", target, true)]),
        ),
    ])
}

/// Every hyperlink target in the IR's text spans, and how many bytes they hold.
fn linked_spans(ir: &DocumentIR) -> (usize, usize) {
    fn walk(elements: &[Element], count: &mut usize, bytes: &mut usize) {
        for e in elements {
            let content = match e {
                Element::Paragraph(p) => &p.content,
                Element::Heading(h) => &h.content,
                _ => continue,
            };
            for c in content {
                if let InlineContent::Text(t) = c {
                    if let Some(u) = &t.hyperlink {
                        *count += 1;
                        *bytes += u.len();
                    }
                }
            }
        }
    }
    let (mut count, mut bytes) = (0, 0);
    for s in &ir.sections {
        walk(&s.elements, &mut count, &mut bytes);
    }
    (count, bytes)
}

fn assert_budget_notice(ir: &DocumentIR) {
    assert!(ir.metadata.text_truncated, "a cut-off must set the truncation flag");
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("character budget")),
        "a cut-off must be reported: {:?}",
        ir.metadata.warnings
    );
}

// ---------------------------------------------------------------------------
// The relationship target itself
// ---------------------------------------------------------------------------

#[test]
fn test_relationship_target_beyond_the_length_cap_is_refused() {
    limits();
    let bytes = docx(&url(1 << 20), 400, false);
    let err = DocxDocument::from_reader(Cursor::new(bytes)).unwrap_err();
    assert!(err.to_string().contains("Target"), "{err}");
}

// ---------------------------------------------------------------------------
// DOCX
// ---------------------------------------------------------------------------

#[test]
fn test_docx_links_naming_one_relationship_share_its_text() {
    limits();
    let doc = DocxDocument::from_reader(Cursor::new(docx(&url(TARGET_LEN), LINKS, false))).unwrap();
    let mut targets = Vec::new();
    for e in &doc.body.elements {
        if let office_oxide::docx::BlockElement::Paragraph(p) = e {
            for c in &p.content {
                if let ParagraphContent::Hyperlink(h) = c {
                    if let HyperlinkTarget::External(t) = &h.target {
                        targets.push(t.clone());
                    }
                }
            }
        }
    }
    assert_eq!(targets.len(), LINKS);
    assert!(
        targets
            .iter()
            .all(|t| std::sync::Arc::ptr_eq(t, &targets[0]))
    );
}

#[test]
fn test_docx_ir_copies_of_a_link_target_are_charged_and_cut_off_loudly() {
    limits();
    let bytes = docx(&url(TARGET_LEN), LINKS, false);
    let ir = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    let (count, bytes) = linked_spans(&ir);
    assert!(count > 0 && bytes <= TEXT_BUDGET, "{count} links, {bytes} bytes");
    assert!(count < LINKS, "the budget must cut the copies off");
    assert_budget_notice(&ir);
    // The cut loses link targets only: every link keeps its text.
    let text = ir.plain_text();
    for i in 0..LINKS {
        assert!(text.contains(&format!("link{i}")), "link{i} lost");
    }
}

#[test]
fn test_docx_markdown_copies_of_a_link_target_are_charged_and_cut_off_loudly() {
    limits();
    let bytes = docx(&url(TARGET_LEN), LINKS, false);
    let md = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .unwrap()
        .to_markdown();
    let copies = md.matches("https://example.com/").count();
    assert!(copies > 0 && copies * TARGET_LEN <= TEXT_BUDGET, "{copies} copies");
    assert!(md.contains("character budget"), "the cut-off is reported");
    for i in 0..LINKS {
        assert!(md.contains(&format!("link{i}")), "link{i} lost");
    }
}

#[test]
fn test_docx_fragment_copies_made_while_reading_are_charged() {
    limits();
    let bytes = docx(&url(TARGET_LEN), LINKS, true);
    let doc = DocxDocument::from_reader(Cursor::new(bytes.clone())).unwrap();
    let mut owned = 0usize;
    for e in &doc.body.elements {
        if let office_oxide::docx::BlockElement::Paragraph(p) = e {
            for c in &p.content {
                if let ParagraphContent::Hyperlink(h) = c {
                    if let HyperlinkTarget::External(t) = &h.target {
                        if t.contains('#') {
                            owned += t.len();
                        }
                    }
                }
            }
        }
    }
    assert!(owned > 0, "links within the budget keep their fragment");
    assert!(owned <= TEXT_BUDGET, "{owned} bytes of composed targets");
    let ir = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("hyperlink targets")),
        "{:?}",
        ir.metadata.warnings
    );
    assert!(ir.metadata.text_truncated);
}
