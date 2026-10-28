//! A relationship target is stored once but named by every hyperlink that
//! refers to it. The reader keeps one copy; the copies that land in an IR or in
//! Markdown are charged to the text budget; and a target too long to be a
//! genuine address is refused outright.
//!
//! These tests lower the process-wide text budget, so each holds the scope
//! that serialises them and puts the budget back.

use std::io::{Cursor, Write};

use office_oxide::docx::{DocxDocument, HyperlinkTarget, ParagraphContent};
use office_oxide::ir::{DocumentIR, Element, InlineContent};
use office_oxide::limits::{max_text_chars, set_max_text_chars};
use office_oxide::{Document, DocumentFormat};

mod common;
use common::LimitScope;

const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const WML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

/// The length of the shared target: a long but admissible address.
const TARGET_LEN: usize = 50_000;
/// About six copies of the target fit.
const TEXT_BUDGET: usize = 300_000;
const LINKS: usize = 40;

fn limits() -> LimitScope {
    let mut scope = LimitScope::new();
    scope.set(max_text_chars, set_max_text_chars, TEXT_BUDGET);
    scope
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
    docx_with_body(&body, target)
}

fn docx_with_body(body: &str, target: &str) -> Vec<u8> {
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

fn pptx(target: &str, links: usize) -> Vec<u8> {
    let mut paras = String::new();
    for i in 0..links {
        paras.push_str(&format!(
            r#"<a:p><a:r><a:rPr><a:hlinkClick r:id="rId1"/></a:rPr><a:t>link{i}</a:t></a:r></a:p>"#
        ));
    }
    let slide = format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/>{paras}</p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#
    );
    pptx_with_slide(slide, target)
}

fn pptx_with_slide(slide: String, target: &str) -> Vec<u8> {
    zip_of(vec![
        ("[Content_Types].xml".to_string(), r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#.to_string()),
        ("_rels/.rels".to_string(), rels(&[("rId1", "officeDocument", "ppt/presentation.xml", false)])),
        (
            "ppt/presentation.xml".to_string(),
            format!(r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#),
        ),
        (
            "ppt/_rels/presentation.xml.rels".to_string(),
            rels(&[("rId2", "slide", "slides/slide1.xml", false)]),
        ),
        ("ppt/slides/slide1.xml".to_string(), slide),
        (
            "ppt/slides/_rels/slide1.xml.rels".to_string(),
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
    let _limits = limits();
    let bytes = docx(&url(1 << 20), 400, false);
    let err = DocxDocument::from_reader(Cursor::new(bytes)).unwrap_err();
    assert!(err.to_string().contains("Target"), "{err}");
    let bytes = pptx(&url(1 << 20), 300);
    // A slide whose relationships are refused is reported, not shown without
    // them in silence.
    let doc = office_oxide::pptx::PptxDocument::from_reader(Cursor::new(bytes)).unwrap();
    let (part, reason) = &doc.unreadable_parts[0];
    assert!(part.ends_with("slide1.xml.rels"), "{part}");
    assert!(reason.contains("Target"), "{reason}");
}

// ---------------------------------------------------------------------------
// DOCX
// ---------------------------------------------------------------------------

#[test]
fn test_docx_links_naming_one_relationship_share_its_text() {
    let _limits = limits();
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
    let _limits = limits();
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
    let _limits = limits();
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
    let _limits = limits();
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

// ---------------------------------------------------------------------------
// PPTX
// ---------------------------------------------------------------------------

#[test]
fn test_pptx_runs_naming_one_relationship_share_its_text_and_ir_copies_are_charged() {
    let _limits = limits();
    let bytes = pptx(&url(TARGET_LEN), LINKS);
    let doc = office_oxide::pptx::PptxDocument::from_reader(Cursor::new(bytes.clone())).unwrap();
    let mut targets = Vec::new();
    for shape in &doc.slides[0].shapes {
        if let office_oxide::pptx::Shape::AutoShape(a) = shape {
            for p in &a.text_body.as_ref().unwrap().paragraphs {
                for c in &p.content {
                    if let office_oxide::pptx::TextContent::Run(r) = c {
                        if let Some(h) = &r.hyperlink {
                            if let office_oxide::pptx::HyperlinkTarget::External(t) = &h.target {
                                targets.push(t.clone());
                            }
                        }
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

    let ir = Document::from_reader(Cursor::new(bytes), DocumentFormat::Pptx)
        .unwrap()
        .to_ir();
    let (count, bytes) = linked_spans(&ir);
    assert!(
        count > 0 && count < LINKS && bytes <= TEXT_BUDGET,
        "{count} links, {bytes} bytes"
    );
    assert_budget_notice(&ir);
    let text = ir.plain_text();
    for i in 0..LINKS {
        assert!(text.contains(&format!("link{i}")), "link{i} lost");
    }
}

/// Pictures linked to one long target: the Markdown names it once per picture.
#[test]
fn test_pptx_markdown_copies_of_a_picture_link_target_are_charged_and_cut_off_loudly() {
    let _limits = limits();
    let pic = |i: usize| {
        format!(
            r#"<p:pic><p:nvPicPr><p:cNvPr id="{id}" name="P{i}" descr="pic{i}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:link="rId1"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:pic>"#,
            id = i + 2
        )
    };
    let pics: String = (0..LINKS).map(pic).collect();
    let slide = format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{pics}</p:spTree></p:cSld></p:sld>"#
    );
    // Short of the whole budget, so the slide's own output budget also holds.
    let len = 40_000;
    let bytes = pptx_with_slide(slide, &url(len).replace("example.com", "img.example.com"));
    let md = Document::from_reader(Cursor::new(bytes), DocumentFormat::Pptx)
        .unwrap()
        .to_markdown();
    let copies = md.matches("https://img.example.com/").count();
    assert!(copies > 0 && copies * len <= TEXT_BUDGET, "{copies} copies");
    assert!(md.contains("character budget"), "the cut-off is reported");
    for i in 0..LINKS {
        assert!(md.contains(&format!("pic{i}")), "pic{i} lost");
    }
}

/// One hyperlink with a long tooltip and many runs: each span owns a copy of
/// the tooltip, which is charged along with the link.
#[test]
fn test_docx_tooltip_copies_are_charged_with_the_link() {
    let _limits = limits();
    let tip = "t".repeat(20_000);
    let runs: String = (0..LINKS)
        .map(|i| format!(r#"<w:r><w:t>run{i} </w:t></w:r>"#))
        .collect();
    let body =
        format!(r#"<w:p><w:hyperlink r:id="rId1" w:tooltip="{tip}">{runs}</w:hyperlink></w:p>"#);
    let bytes = docx_with_body(&body, "https://example.com/");
    let ir = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    let mut tip_bytes = 0;
    for s in &ir.sections {
        for e in &s.elements {
            if let Element::Paragraph(p) = e {
                for c in &p.content {
                    if let InlineContent::Text(t) = c {
                        tip_bytes += t.hyperlink_tooltip.as_ref().map_or(0, String::len);
                    }
                }
            }
        }
    }
    assert!(tip_bytes > 0 && tip_bytes <= TEXT_BUDGET, "{tip_bytes} tooltip bytes");
    assert_budget_notice(&ir);
    let text = ir.plain_text();
    for i in 0..LINKS {
        assert!(text.contains(&format!("run{i}")), "run{i} lost");
    }
}
