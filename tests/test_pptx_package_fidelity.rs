//! PPTX package-level fidelity: parts reached through relationships
//! (notes, comments, comment authors, SmartArt data, media, charts, fonts,
//! layouts and masters) and the degradation policy when one is broken.
//!
//! Every fixture is a minimal synthetic package built in code
//! (AGENTS.md rule 4). The builder writes the zip entries directly, so a
//! test can produce packages the crate's own writer never would — dangling
//! relationships, malformed parts, missing `r:id`s.

use std::collections::BTreeMap;
use std::io::{Cursor, Write};

use office_oxide::core::relationships::rel_types;
use office_oxide::ir::*;
use office_oxide::pptx::PptxDocument;
use office_oxide::{Document, DocumentFormat};

// ---------------------------------------------------------------------------
// Package builder
// ---------------------------------------------------------------------------

const NS: &str = r#"xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;
const CT_PML: &str = "application/vnd.openxmlformats-officedocument.presentationml.";

/// A raw OPC package: part bytes, content-type overrides and relationships,
/// written verbatim.
#[derive(Default)]
struct Pkg {
    parts: BTreeMap<String, Vec<u8>>,
    overrides: Vec<(String, String)>,
    /// Source part (`""` for the package) → `<Relationship …/>` elements.
    rels: BTreeMap<String, Vec<String>>,
}

impl Pkg {
    /// Add a part at `name` (no leading slash) with a content-type override.
    fn part(&mut self, name: &str, ct: &str, data: impl Into<Vec<u8>>) -> &mut Self {
        self.parts.insert(name.to_string(), data.into());
        if !ct.is_empty() {
            self.overrides.push((format!("/{name}"), ct.to_string()));
        }
        self
    }

    /// Add an internal relationship from `source` (no leading slash; `""`
    /// for the package) with an explicit id.
    fn rel(&mut self, source: &str, id: &str, ty: &str, target: &str) -> &mut Self {
        self.rels
            .entry(source.to_string())
            .or_default()
            .push(format!(r#"<Relationship Id="{id}" Type="{ty}" Target="{target}"/>"#));
        self
    }

    /// Add an external relationship.
    fn ext_rel(&mut self, source: &str, id: &str, ty: &str, target: &str) -> &mut Self {
        self.rels
            .entry(source.to_string())
            .or_default()
            .push(format!(
                r#"<Relationship Id="{id}" Type="{ty}" Target="{target}" TargetMode="External"/>"#
            ));
        self
    }

    fn build(&self) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = zip::write::SimpleFileOptions::default();
        let mut ct = String::from(
            r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="ttf" ContentType="application/x-font-ttf"/>"#,
        );
        for (name, t) in &self.overrides {
            ct.push_str(&format!(r#"<Override PartName="{name}" ContentType="{t}"/>"#));
        }
        ct.push_str("</Types>");
        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(ct.as_bytes()).unwrap();
        for (source, rels) in &self.rels {
            let path = match source.rsplit_once('/') {
                _ if source.is_empty() => "_rels/.rels".to_string(),
                Some((dir, file)) => format!("{dir}/_rels/{file}.rels"),
                None => format!("_rels/{source}.rels"),
            };
            zip.start_file(path, opts).unwrap();
            zip.write_all(
                format!(
                    r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}</Relationships>"#,
                    rels.concat()
                )
                .as_bytes(),
            )
            .unwrap();
        }
        for (name, data) in &self.parts {
            zip.start_file(name.as_str(), opts).unwrap();
            zip.write_all(data).unwrap();
        }
        zip.finish().unwrap().into_inner()
    }

    fn open(&self) -> PptxDocument {
        PptxDocument::from_reader(Cursor::new(self.build())).expect("open pptx")
    }

    fn document(&self) -> Document {
        Document::from_reader(Cursor::new(self.build()), DocumentFormat::Pptx).expect("open pptx")
    }
}

/// A slide part's XML holding `tree` inside its `<p:spTree>`.
fn slide_xml(tree: &str) -> String {
    format!(
        r#"<?xml version="1.0"?><p:sld {NS}><p:cSld><p:spTree>{tree}</p:spTree></p:cSld></p:sld>"#
    )
}

/// A text box shape (no placeholder) holding one run of `text`.
fn text_sp(text: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
    )
}

/// A deck of slides `ppt/slides/slide{N}.xml`, each holding `trees[N-1]`,
/// related from the presentation as `rId{N}`.
fn deck(trees: &[&str]) -> Pkg {
    let mut pkg = Pkg::default();
    pkg.rel("", "rId1", rel_types::OFFICE_DOCUMENT, "ppt/presentation.xml");
    let mut ids = String::new();
    for (i, tree) in trees.iter().enumerate() {
        let n = i + 1;
        ids.push_str(&format!(r#"<p:sldId id="{}" r:id="rId{n}"/>"#, 255 + n));
        pkg.rel(
            "ppt/presentation.xml",
            &format!("rId{n}"),
            rel_types::SLIDE,
            &format!("slides/slide{n}.xml"),
        );
        pkg.part(
            &format!("ppt/slides/slide{n}.xml"),
            &format!("{CT_PML}slide+xml"),
            slide_xml(tree),
        );
    }
    pkg.part(
        "ppt/presentation.xml",
        &format!("{CT_PML}presentation.main+xml"),
        format!(r#"<?xml version="1.0"?><p:presentation {NS}><p:sldIdLst>{ids}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#),
    );
    pkg
}

/// A notes-slide part whose body placeholder holds `paragraphs`.
fn notes_xml(paragraphs: &str) -> String {
    format!(
        r#"<?xml version="1.0"?><p:notes {NS}><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/>{paragraphs}</p:txBody></p:sp></p:spTree></p:cSld></p:notes>"#
    )
}

fn add_notes(pkg: &mut Pkg, slide: usize, paragraphs: &str) {
    let name = format!("ppt/notesSlides/notesSlide{slide}.xml");
    pkg.part(&name, &format!("{CT_PML}notesSlide+xml"), notes_xml(paragraphs));
    pkg.rel(
        &format!("ppt/slides/slide{slide}.xml"),
        "rIdN",
        rel_types::NOTES_SLIDE,
        &format!("../notesSlides/notesSlide{slide}.xml"),
    );
}

fn spans(elements: &[Element]) -> Vec<&TextSpan> {
    fn walk<'a>(els: &'a [Element], out: &mut Vec<&'a TextSpan>) {
        for e in els {
            match e {
                Element::Paragraph(p) => out.extend(p.content.iter().filter_map(|c| match c {
                    InlineContent::Text(t) => Some(t),
                    _ => None,
                })),
                Element::List(l) => {
                    for item in &l.items {
                        walk(&item.content, out);
                    }
                },
                Element::TextBox(tb) => walk(&tb.content, out),
                _ => {},
            }
        }
    }
    let mut out = Vec::new();
    walk(elements, &mut out);
    out
}

// ---------------------------------------------------------------------------
// Degradation policy: one broken part does not cost the rest of the deck
// ---------------------------------------------------------------------------

/// A notes slide is review content attached to a slide. A malformed notes
/// part was dropped silently, and one whose relationship target could not
/// be resolved failed the whole deck — while an unresolvable *slide*
/// relationship was skipped. Every unreadable part is now skipped, logged
/// and recorded, and the rest of the deck is extracted.
#[test]
fn test_an_unreadable_notes_part_is_recorded_and_the_deck_still_opens() {
    let mut pkg = deck(&[
        &text_sp("SLIDE ONE"),
        &text_sp("SLIDE TWO"),
        &text_sp("SLIDE THREE"),
    ]);
    // Slide 1: notes part is not well-formed XML.
    pkg.part(
        "ppt/notesSlides/notesSlide1.xml",
        &format!("{CT_PML}notesSlide+xml"),
        "<p:notes><p:cSld><unclosed",
    );
    pkg.rel(
        "ppt/slides/slide1.xml",
        "rIdN",
        rel_types::NOTES_SLIDE,
        "../notesSlides/notesSlide1.xml",
    );
    // Slide 2: notes relationship target is not a legal part name.
    pkg.rel(
        "ppt/slides/slide2.xml",
        "rIdN",
        rel_types::NOTES_SLIDE,
        "../notesSlides/n.xml?v=1",
    );
    // Slide 3: intact notes.
    add_notes(&mut pkg, 3, "<a:p><a:r><a:t>INTACT NOTES</a:t></a:r></a:p>");

    let doc = pkg.open();
    let text = doc.plain_text();
    for want in ["SLIDE ONE", "SLIDE TWO", "SLIDE THREE", "INTACT NOTES"] {
        assert!(text.contains(want), "{want} missing from {text:?}");
    }
    assert_eq!(doc.unreadable_parts.len(), 2, "{:?}", doc.unreadable_parts);
    assert!(
        doc.unreadable_parts
            .iter()
            .any(|(p, _)| p.contains("notesSlide1"))
    );
    assert!(
        doc.unreadable_parts
            .iter()
            .any(|(p, _)| p.contains("n.xml"))
    );
    let ir = pkg.document().to_ir();
    assert!(ir.metadata.text_truncated, "the loss must be on record");
}

/// The same policy for slides: an unreadable slide part (or a slide
/// relationship that cannot be resolved) is a notice in its own position,
/// and the other slides are extracted. It used to fail the whole deck (or,
/// for the relationship, vanish without a trace).
#[test]
fn test_an_unreadable_slide_is_a_notice_in_place_and_the_others_are_kept() {
    let mut pkg = deck(&[&text_sp("FIRST"), &text_sp("SECOND"), &text_sp("THIRD")]);
    pkg.part(
        "ppt/slides/slide2.xml",
        &format!("{CT_PML}slide+xml"),
        "<p:sld><p:cSld><unclosed",
    );
    let doc = pkg.open();
    assert_eq!(doc.slides.len(), 3);
    assert!(doc.slides[1].parse_error.is_some());
    assert_eq!(doc.unreadable_parts.len(), 1);

    let ir = pkg.document().to_ir();
    assert_eq!(ir.sections.len(), 3, "the unreadable slide keeps its place");
    let plain = ir.plain_text();
    let (first, second, third) = (
        plain.find("FIRST").unwrap(),
        plain.find("[unreadable slide").expect("notice"),
        plain.find("THIRD").unwrap(),
    );
    assert!(first < second && second < third, "{plain:?}");
    assert!(ir.metadata.text_truncated);
    let direct = pkg.document().plain_text();
    assert!(direct.contains("[unreadable slide"), "direct renderer: {direct:?}");

    // A dangling slide relationship is recorded the same way.
    let mut pkg = deck(&[&text_sp("ONLY")]);
    pkg.rels.get_mut("ppt/presentation.xml").unwrap()[0] = format!(
        r#"<Relationship Id="rId1" Type="{}" Target="slides/missing.xml"/>"#,
        rel_types::SLIDE
    );
    pkg.part(
        "ppt/slides/slide9.xml",
        &format!("{CT_PML}slide+xml"),
        slide_xml(&text_sp("NINE")),
    );
    pkg.rel("ppt/presentation.xml", "rId9", rel_types::SLIDE, "slides/slide9.xml");
    let pres = String::from_utf8(pkg.parts["ppt/presentation.xml"].clone()).unwrap();
    pkg.parts.insert(
        "ppt/presentation.xml".into(),
        pres.replace("</p:sldIdLst>", r#"<p:sldId id="300" r:id="rId9"/></p:sldIdLst>"#)
            .into_bytes(),
    );
    let doc = pkg.open();
    assert_eq!(doc.slides.len(), 2);
    assert!(doc.slides[0].parse_error.is_some());
    assert!(doc.plain_text().contains("NINE"));
}

/// When no slide at all can be read there is nothing to degrade to.
#[test]
fn test_a_deck_with_no_readable_slide_is_an_error() {
    let mut pkg = deck(&[&text_sp("X")]);
    pkg.part("ppt/slides/slide1.xml", &format!("{CT_PML}slide+xml"), "<p:sld><unclosed");
    assert!(PptxDocument::from_reader(Cursor::new(pkg.build())).is_err());
}

/// Hyperlinks in speaker notes resolve through the notes part's own
/// relationships. The notes body was parsed with an empty relationship
/// set, so every notes hyperlink was dropped on read.
#[test]
fn test_notes_hyperlinks_resolve_through_the_notes_part_rels() {
    let mut pkg = deck(&[&text_sp("BODY")]);
    add_notes(
        &mut pkg,
        1,
        r#"<a:p><a:r><a:rPr><a:hlinkClick r:id="rIdL"/></a:rPr><a:t>see docs</a:t></a:r></a:p>"#,
    );
    pkg.ext_rel(
        "ppt/notesSlides/notesSlide1.xml",
        "rIdL",
        rel_types::HYPERLINK,
        "https://example.com/notes",
    );
    let ir = pkg.document().to_ir();
    let notes = ir.sections[0].speaker_notes.as_ref().expect("notes");
    let link = spans(notes)
        .into_iter()
        .find(|s| s.text == "see docs")
        .and_then(|s| s.hyperlink.clone());
    assert_eq!(link.as_deref(), Some("https://example.com/notes"));
}

// ---------------------------------------------------------------------------
// Comment authors
// ---------------------------------------------------------------------------

const REL_LEGACY_COMMENTS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const REL_LEGACY_AUTHORS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors";
const REL_MODERN_COMMENTS: &str =
    "http://schemas.microsoft.com/office/2018/10/relationships/comments";
const REL_MODERN_AUTHORS: &str =
    "http://schemas.microsoft.com/office/2018/10/relationships/authors";

/// Comment authors live in a presentation-level part —
/// `ppt/commentAuthors.xml` for legacy comments, `ppt/authors.xml` for
/// modern ones — and nothing read either, so every comment's author was
/// `None`. Modern (threaded) comment parts are read too, replies included.
#[test]
fn test_comment_authors_resolve_for_legacy_and_modern_comments() {
    let mut pkg = deck(&[&text_sp("ONE"), &text_sp("TWO")]);
    // Legacy comment on slide 1.
    pkg.part(
        "ppt/commentAuthors.xml",
        &format!("{CT_PML}commentAuthors+xml"),
        format!(
            r#"<?xml version="1.0"?><p:cmAuthorLst {NS}><p:cmAuthor id="0" name="Ada Lovelace" initials="AL" lastIdx="1" clrIdx="0"/><p:cmAuthor id="1" name="Alan Turing" initials="AT" lastIdx="1" clrIdx="1"/></p:cmAuthorLst>"#
        ),
    );
    pkg.rel("ppt/presentation.xml", "rIdCA", REL_LEGACY_AUTHORS, "commentAuthors.xml");
    pkg.part(
        "ppt/comments/comment1.xml",
        &format!("{CT_PML}comments+xml"),
        format!(
            r#"<?xml version="1.0"?><p:cmLst {NS}><p:cm authorId="1" idx="1"><p:pos x="10" y="10"/><p:text>LEGACY NOTE</p:text></p:cm></p:cmLst>"#
        ),
    );
    pkg.rel("ppt/slides/slide1.xml", "rIdC", REL_LEGACY_COMMENTS, "../comments/comment1.xml");
    // Modern comment with a reply on slide 2.
    pkg.part(
        "ppt/authors.xml",
        "application/vnd.ms-powerpoint.authors+xml",
        r#"<?xml version="1.0"?><p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><p188:author id="{A1}" name="Grace Hopper" initials="GH" userId="g" providerId="None"/><p188:author id="{A2}" name="Edsger Dijkstra" initials="ED" userId="e" providerId="None"/></p188:authorLst>"#,
    );
    pkg.rel("ppt/presentation.xml", "rIdAU", REL_MODERN_AUTHORS, "authors.xml");
    pkg.part(
        "ppt/comments/modernComment_101_0.xml",
        "application/vnd.ms-powerpoint.comments+xml",
        r#"<?xml version="1.0"?><p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p188:cm id="{C1}" authorId="{A1}" created="2026-01-01T00:00:00Z"><p188:replyLst><p188:reply id="{R1}" authorId="{A2}" created="2026-01-02T00:00:00Z"><p188:txBody><a:bodyPr/><a:p><a:r><a:t>REPLY TEXT</a:t></a:r></a:p></p188:txBody></p188:reply></p188:replyLst><p188:txBody><a:bodyPr/><a:p><a:r><a:t>MODERN NOTE</a:t></a:r></a:p></p188:txBody></p188:cm></p188:cmLst>"#,
    );
    pkg.rel(
        "ppt/slides/slide2.xml",
        "rIdM",
        REL_MODERN_COMMENTS,
        "../comments/modernComment_101_0.xml",
    );

    let doc = pkg.open();
    let authored = |i: usize| -> Vec<(Option<String>, String)> {
        doc.slides[i]
            .comments
            .iter()
            .map(|c| (c.author.clone(), c.text.clone()))
            .collect()
    };
    assert_eq!(authored(0), vec![(Some("Alan Turing".into()), "LEGACY NOTE".into())]);
    assert_eq!(
        authored(1),
        vec![
            (Some("Grace Hopper".into()), "MODERN NOTE".into()),
            (Some("Edsger Dijkstra".into()), "REPLY TEXT".into()),
        ]
    );
    let ir = pkg.document().to_ir();
    let authors: Vec<Option<String>> = ir
        .sections
        .iter()
        .flat_map(|s| &s.elements)
        .filter_map(|e| match e {
            Element::Endnote(n) => Some(n.author.clone()),
            _ => None,
        })
        .collect();
    assert_eq!(
        authors,
        [
            Some("Alan Turing"),
            Some("Grace Hopper"),
            Some("Edsger Dijkstra")
        ]
        .map(|a| a.map(String::from))
    );
}

// ---------------------------------------------------------------------------
// SmartArt
// ---------------------------------------------------------------------------

/// A real SmartArt `graphicFrame` holds only `<dgm:relIds r:dm=…/>`; the
/// node text lives in the separate `ppt/diagrams/dataN.xml` part, which
/// was never resolved — so real SmartArt extracted as nothing.
#[test]
fn test_smartart_text_is_read_from_the_diagram_data_part() {
    let frame = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Diagram 3"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="100" y="100"/><a:ext cx="5000" cy="3000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId5" r:lo="rId6" r:qs="rId7" r:cs="rId8"/></a:graphicData></a:graphic></p:graphicFrame>"#;
    let mut pkg = deck(&[frame]);
    pkg.part(
        "ppt/diagrams/data1.xml",
        "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
        r#"<?xml version="1.0"?><dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst>
            <dgm:pt modelId="{0}" type="doc"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:p><a:endParaRPr/></a:p></dgm:t></dgm:pt>
            <dgm:pt modelId="{1}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:p><a:r><a:t>Plan</a:t></a:r><a:r><a:t>ning</a:t></a:r></a:p></dgm:t></dgm:pt>
            <dgm:pt modelId="{2}" type="parTrans"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:p><a:endParaRPr/></a:p></dgm:t></dgm:pt>
            <dgm:pt modelId="{3}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:p><a:r><a:t>Delivery</a:t></a:r></a:p></dgm:t></dgm:pt>
          </dgm:ptLst><dgm:cxnLst/></dgm:dataModel>"#,
    );
    pkg.rel(
        "ppt/slides/slide1.xml",
        "rId5",
        rel_types::DIAGRAM_DATA,
        "../diagrams/data1.xml",
    );

    let doc = pkg.document();
    let ir = doc.to_ir();
    let text = ir.plain_text();
    assert!(text.contains("Planning"), "runs of one node are one line: {text:?}");
    assert!(text.contains("Delivery"), "{text:?}");
    assert!(text.find("Planning") < text.find("Delivery"));
    let direct = doc.plain_text();
    assert!(direct.contains("Planning") && direct.contains("Delivery"), "{direct:?}");
}

// ---------------------------------------------------------------------------
// Linked pictures, media and OLE objects
// ---------------------------------------------------------------------------

/// A minimal valid PNG header — enough for format sniffing.
const PNG: &[u8] = &[0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 0];

fn picture(nv_pr: &str, blip: &str, descr: &str) -> String {
    format!(
        r#"<p:pic><p:nvPicPr><p:cNvPr id="5" name="Pic" descr="{descr}"/><p:cNvPicPr/><p:nvPr>{nv_pr}</p:nvPr></p:nvPicPr><p:blipFill>{blip}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="10" y="10"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:pic>"#
    )
}

fn images(ir: &DocumentIR) -> Vec<&Image> {
    fn walk<'a>(els: &'a [Element], out: &mut Vec<&'a Image>) {
        for e in els {
            match e {
                Element::Image(i) => out.push(i),
                Element::TextBox(tb) => walk(&tb.content, out),
                _ => {},
            }
        }
    }
    let mut out = Vec::new();
    for s in &ir.sections {
        walk(&s.elements, &mut out);
    }
    out
}

/// `<a:blip r:link>` references an image outside the package. Only
/// `r:embed` was read, so a linked picture came through with no trace of
/// where its image lives.
#[test]
fn test_a_linked_picture_keeps_its_external_target() {
    let mut pkg = deck(&[&picture("", r#"<a:blip r:link="rIdL"/>"#, "Company logo")]);
    pkg.ext_rel(
        "ppt/slides/slide1.xml",
        "rIdL",
        rel_types::IMAGE,
        "https://example.com/logo.png",
    );
    let doc = pkg.open();
    let office_oxide::pptx::Shape::Picture(ref pic) = doc.slides[0].shapes[0] else {
        panic!("expected a picture");
    };
    assert_eq!(pic.link_target.as_deref(), Some("https://example.com/logo.png"));
    let ir = pkg.document().to_ir();
    assert_eq!(images(&ir)[0].source_url.as_deref(), Some("https://example.com/logo.png"));
    let md = ir.to_markdown();
    assert!(md.contains("![Company logo](https://example.com/logo.png)"), "{md}");
    let html = ir.to_html();
    assert!(
        html.contains(r#"<img src="https://example.com/logo.png" alt="Company logo" />"#),
        "{html}"
    );
    let direct = pkg.document().to_markdown();
    assert!(direct.contains("](https://example.com/logo.png)"), "{direct}");
}

/// A dangerous scheme on a linked picture is not rendered as a source.
#[test]
fn test_a_linked_picture_with_a_script_scheme_is_not_rendered_as_a_source() {
    let mut pkg = deck(&[&picture("", r#"<a:blip r:link="rIdL"/>"#, "x")]);
    pkg.ext_rel("ppt/slides/slide1.xml", "rIdL", rel_types::IMAGE, "javascript:alert(1)");
    let doc = pkg.document();
    for out in [
        doc.to_ir().to_markdown(),
        doc.to_ir().to_html(),
        doc.to_markdown(),
    ] {
        assert!(!out.contains("javascript"), "{out}");
    }
}

/// A video shape is a picture (its poster frame) whose `p:nvPr` names the
/// clip; the clip reference was skipped.
#[test]
fn test_a_video_shape_surfaces_its_clip() {
    let nv_pr = r#"<a:videoFile r:link="rIdV"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DA1B2A7D8E3}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rIdM"/></p:ext></p:extLst>"#;
    let mut pkg = deck(&[&picture(nv_pr, r#"<a:blip r:embed="rIdP"/>"#, "")]);
    pkg.part("ppt/media/image1.png", "", PNG);
    pkg.part("ppt/media/media1.mp4", "", b"fake mp4".to_vec());
    pkg.rel("ppt/slides/slide1.xml", "rIdP", rel_types::IMAGE, "../media/image1.png");
    pkg.rel(
        "ppt/slides/slide1.xml",
        "rIdV",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video",
        "../media/media1.mp4",
    );
    pkg.rel(
        "ppt/slides/slide1.xml",
        "rIdM",
        "http://schemas.microsoft.com/office/2007/relationships/media",
        "../media/media1.mp4",
    );
    let doc = pkg.open();
    let office_oxide::pptx::Shape::Picture(ref pic) = doc.slides[0].shapes[0] else {
        panic!("expected a picture");
    };
    let media = pic.media.as_ref().expect("the clip reference");
    assert_eq!(media.kind, office_oxide::pptx::MediaKind::Video);
    assert_eq!(media.target, "../media/media1.mp4");
    assert!(!media.external);
    assert_eq!(pic.data.as_deref(), Some(PNG), "the poster frame is still the picture");
}

/// An OLE object frame fell through as an unknown graphic, losing its
/// preview picture. It is now an `OleObject` and its preview reaches the
/// IR as an image.
#[test]
fn test_an_ole_object_keeps_its_preview_picture() {
    let frame = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="Object 5"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="100" y="100"/><a:ext cx="2000" cy="1000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:v="urn:schemas-microsoft-com:vml" Requires="v"><p:oleObj spid="_x0000_s1026" name="Worksheet" r:id="rIdO" imgW="100" imgH="50" progId="Excel.Sheet.12"><p:embed/></p:oleObj></mc:Choice><mc:Fallback><p:oleObj name="Worksheet" r:id="rIdO" imgW="100" imgH="50" progId="Excel.Sheet.12"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdI"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr/></p:pic></p:oleObj></mc:Fallback></mc:AlternateContent></a:graphicData></a:graphic></p:graphicFrame>"#;
    let mut pkg = deck(&[frame, &text_sp("AFTER")]);
    pkg.part("ppt/media/image1.png", "", PNG);
    pkg.part("ppt/embeddings/Microsoft_Excel_Worksheet.xlsx", "", b"PK".to_vec());
    pkg.rel("ppt/slides/slide1.xml", "rIdI", rel_types::IMAGE, "../media/image1.png");
    pkg.rel(
        "ppt/slides/slide1.xml",
        "rIdO",
        rel_types::PACKAGE,
        "../embeddings/Microsoft_Excel_Worksheet.xlsx",
    );

    let doc = pkg.open();
    let office_oxide::pptx::Shape::GraphicFrame(ref gf) = doc.slides[0].shapes[0] else {
        panic!("expected a graphic frame");
    };
    let office_oxide::pptx::GraphicContent::OleObject(ref ole) = gf.content else {
        panic!("expected an OLE object, got {:?}", gf.content);
    };
    assert_eq!(ole.prog_id.as_deref(), Some("Excel.Sheet.12"));
    assert_eq!(ole.name.as_deref(), Some("Worksheet"));
    assert_eq!(ole.rel_id.as_deref(), Some("rIdO"));
    assert_eq!(ole.preview_data.as_deref(), Some(PNG));
    let ir = pkg.document().to_ir();
    let imgs = images(&ir);
    assert_eq!(imgs.len(), 1);
    assert_eq!(imgs[0].data.as_deref(), Some(PNG));
    assert!(ir.plain_text().contains("AFTER"), "the reader position stays right");
}

// ---------------------------------------------------------------------------
// Text-style inheritance and layout/master static text
// ---------------------------------------------------------------------------

fn ph_sp(ph: &str, lst_style: &str, paragraphs: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="S"/><p:cNvSpPr/><p:nvPr>{ph}</p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle>{lst_style}</a:lstStyle>{paragraphs}</p:txBody></p:sp>"#
    )
}

fn run(text: &str) -> String {
    format!("<a:r><a:rPr/><a:t>{text}</a:t></a:r>")
}

/// A deck whose one slide uses a layout (with a static text box) on a
/// master with `txStyles`, in a presentation with a `defaultTextStyle`.
fn styled_deck(slide_attrs: &str) -> Pkg {
    let slide = [
        ph_sp(r#"<p:ph type="title"/>"#, "", &format!("<a:p>{}</a:p>", run("TITLE"))),
        ph_sp(
            r#"<p:ph idx="1"/>"#,
            "",
            &format!(
                r#"<a:p>{}</a:p><a:p><a:pPr lvl="1"/>{}</a:p>"#,
                run("LEVEL ONE"),
                run("LEVEL TWO")
            ),
        ),
        ph_sp(r#"<p:ph type="dt" idx="10"/>"#, "", &format!("<a:p>{}</a:p>", run("DATE"))),
        ph_sp(
            "",
            r#"<a:lvl1pPr><a:defRPr b="1"/></a:lvl1pPr>"#,
            &format!("<a:p>{}</a:p>", run("TEXT BOX")),
        ),
    ]
    .concat();
    let mut pkg = deck(&[&slide]);
    pkg.parts.insert(
        "ppt/slides/slide1.xml".into(),
        format!(
            r#"<?xml version="1.0"?><p:sld {NS} {slide_attrs}><p:cSld><p:spTree>{slide}</p:spTree></p:cSld></p:sld>"#
        )
        .into_bytes(),
    );
    let pres = String::from_utf8(pkg.parts["ppt/presentation.xml"].clone()).unwrap();
    pkg.parts.insert(
        "ppt/presentation.xml".into(),
        pres.replace(
            "</p:presentation>",
            r#"<p:defaultTextStyle><a:defPPr><a:defRPr sz="1800"/></a:defPPr></p:defaultTextStyle></p:presentation>"#,
        )
        .into_bytes(),
    );
    // Layout: the body placeholder centres and underlines level 1; a
    // plain text box carries boilerplate.
    let layout_tree = [
        ph_sp(
            r#"<p:ph idx="1"/>"#,
            r#"<a:lvl1pPr algn="ctr"><a:defRPr u="sng"/></a:lvl1pPr>"#,
            &format!("<a:p>{}</a:p>", run("Click to edit")),
        ),
        text_sp("ACME CONFIDENTIAL"),
    ]
    .concat();
    pkg.part(
        "ppt/slideLayouts/slideLayout1.xml",
        &format!("{CT_PML}slideLayout+xml"),
        format!(
            r#"<?xml version="1.0"?><p:sldLayout {NS}><p:cSld name="Title and Content"><p:spTree>{layout_tree}</p:spTree></p:cSld></p:sldLayout>"#
        ),
    );
    pkg.rel(
        "ppt/slides/slide1.xml",
        "rIdLay",
        rel_types::SLIDE_LAYOUT,
        "../slideLayouts/slideLayout1.xml",
    );
    // Master: its body placeholder colours level 1 red; txStyles give
    // sizes per level and category.
    let master_tree = [
        ph_sp(r#"<p:ph type="body" idx="1"/>"#, r#"<a:lvl1pPr><a:defRPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:defRPr></a:lvl1pPr>"#, ""),
        text_sp("MASTER FOOTER"),
    ]
    .concat();
    pkg.part(
        "ppt/slideMasters/slideMaster1.xml",
        &format!("{CT_PML}slideMaster+xml"),
        format!(
            r#"<?xml version="1.0"?><p:sldMaster {NS}><p:cSld><p:spTree>{master_tree}</p:spTree></p:cSld><p:txStyles>
              <p:titleStyle><a:lvl1pPr><a:defRPr sz="4400" b="1"/></a:lvl1pPr></p:titleStyle>
              <p:bodyStyle><a:lvl1pPr><a:defRPr sz="3200"/></a:lvl1pPr><a:lvl2pPr><a:defRPr sz="2800" i="1"/></a:lvl2pPr></p:bodyStyle>
              <p:otherStyle><a:lvl1pPr><a:defRPr sz="1200"/></a:lvl1pPr></p:otherStyle>
            </p:txStyles></p:sldMaster>"#
        ),
    );
    pkg.rel(
        "ppt/slideLayouts/slideLayout1.xml",
        "rIdM",
        rel_types::SLIDE_MASTER,
        "../slideMasters/slideMaster1.xml",
    );
    pkg
}

fn span_named<'a>(ir: &'a DocumentIR, text: &str) -> &'a TextSpan {
    let mut all = Vec::new();
    for s in &ir.sections {
        all.extend(spans(&s.elements));
        for e in &s.elements {
            if let Element::Heading(h) = e {
                all.extend(h.content.iter().filter_map(|c| match c {
                    InlineContent::Text(t) => Some(t),
                    _ => None,
                }));
            }
        }
    }
    all.into_iter()
        .find(|s| s.text == text)
        .unwrap_or_else(|| panic!("no span {text:?}"))
}

/// Formatting is inherited through the whole ECMA-376 chain: the layout's
/// placeholder, the master's placeholder, the master's per-level
/// `txStyles` (title/body/other) and the presentation default. Only the
/// master's level-1 title/body style used to be read.
#[test]
fn test_text_formatting_inherits_through_layout_master_and_presentation() {
    let doc = styled_deck("").open();
    let runs = |i: usize| -> Vec<(String, office_oxide::pptx::TextRun)> {
        let office_oxide::pptx::Shape::AutoShape(ref a) = doc.slides[0].shapes[i] else {
            panic!()
        };
        a.text_body
            .as_ref()
            .unwrap()
            .paragraphs
            .iter()
            .flat_map(|p| &p.content)
            .filter_map(|c| match c {
                office_oxide::pptx::TextContent::Run(r) => Some((r.text.clone(), r.clone())),
                _ => None,
            })
            .collect()
    };
    // Title: master titleStyle.
    let (_, title) = &runs(0)[0];
    assert_eq!((title.font_size_hundredths_pt, title.bold), (Some(4400), Some(true)));
    // Body level 1: layout placeholder (underline), master placeholder
    // (colour), master bodyStyle level 1 (size).
    let body = runs(1);
    let one = &body[0].1;
    assert_eq!(one.underline.as_deref(), Some("sng"));
    assert_eq!(one.color_rgb, Some([0xFF, 0, 0]));
    assert_eq!(one.font_size_hundredths_pt, Some(3200));
    // Body level 2: master bodyStyle level 2; the level-1-only layout and
    // master placeholder styles do not apply.
    let two = &body[1].1;
    assert_eq!((two.font_size_hundredths_pt, two.italic), (Some(2800), Some(true)));
    assert_eq!((two.underline.as_deref(), two.color_rgb), (None, None));
    // Date placeholder: otherStyle.
    assert_eq!(runs(2)[0].1.font_size_hundredths_pt, Some(1200));
    // A plain text box: its own list style, then otherStyle.
    let tb = &runs(3)[0].1;
    assert_eq!((tb.bold, tb.font_size_hundredths_pt), (Some(true), Some(1200)));
    let office_oxide::pptx::Shape::AutoShape(ref a) = doc.slides[0].shapes[1] else {
        panic!()
    };
    assert_eq!(
        a.text_body.as_ref().unwrap().paragraphs[0].alignment,
        Some(ParagraphAlignment::Center),
        "layout placeholder alignment"
    );

    let ir = styled_deck("").document().to_ir();
    let one = span_named(&ir, "LEVEL ONE");
    assert_eq!(one.font_size_half_pt, Some(64));
    assert_eq!(one.color, Some([0xFF, 0, 0]));
}

/// A body placeholder's paragraphs are bulleted because the master's
/// `bodyStyle` says so, not their own `<a:pPr>`. The bullet was not part of
/// the inherited properties, so such a list read as plain paragraphs — or,
/// once any paragraph was indented, the whole body became one list.
#[test]
fn test_body_placeholder_bullets_inherit_from_the_master_body_style() {
    let mut pkg = styled_deck("");
    let master = String::from_utf8(pkg.parts["ppt/slideMasters/slideMaster1.xml"].clone()).unwrap();
    let master = master
        .replace(
            r#"<a:lvl1pPr><a:defRPr sz="3200"/>"#,
            r#"<a:lvl1pPr><a:buChar char="&#8226;"/><a:defRPr sz="3200"/>"#,
        )
        .replace(
            r#"<a:lvl2pPr><a:defRPr sz="2800" i="1"/>"#,
            r#"<a:lvl2pPr><a:buChar char="&#8211;"/><a:defRPr sz="2800" i="1"/>"#,
        );
    pkg.parts
        .insert("ppt/slideMasters/slideMaster1.xml".into(), master.into_bytes());
    let doc = pkg.document();
    for md in [doc.to_markdown(), doc.to_ir().to_markdown()] {
        assert!(md.contains("- LEVEL ONE\n  - *LEVEL TWO*"), "{md}");
        // The text box and the date placeholder use `otherStyle`: no bullets.
        assert!(!md.contains("- DATE") && !md.contains("- TEXT BOX"), "{md}");
    }
    assert_eq!(doc.to_markdown().trim_end(), doc.to_ir().to_markdown().trim_end());
}

/// With no master `txStyles` for a level, the presentation's
/// `defaultTextStyle` is the last layer.
#[test]
fn test_presentation_default_text_style_is_the_last_layer() {
    let mut pkg = styled_deck("");
    let master = String::from_utf8(pkg.parts["ppt/slideMasters/slideMaster1.xml"].clone()).unwrap();
    pkg.parts.insert(
        "ppt/slideMasters/slideMaster1.xml".into(),
        master
            .replace(r#"<a:lvl1pPr><a:defRPr sz="1200"/></a:lvl1pPr>"#, "")
            .into_bytes(),
    );
    let doc = pkg.open();
    let office_oxide::pptx::Shape::AutoShape(ref a) = doc.slides[0].shapes[2] else {
        panic!()
    };
    let office_oxide::pptx::TextContent::Run(ref r) =
        a.text_body.as_ref().unwrap().paragraphs[0].content[0]
    else {
        panic!()
    };
    assert_eq!(r.font_size_hundredths_pt, Some(1800));
}

/// Text placed directly on a layout or master is available per slide,
/// honouring `showMasterSp`. It is not any one slide's text: the slide's
/// own text excludes it, and the document carries it once, after the
/// slides. (This used to assert it was absent from `plain_text()`
/// altogether, the python-pptx default; that dropped text Apache POI/Tika
/// extract and PowerPoint shows on every slide.)
#[test]
fn test_layout_and_master_static_text_is_available_but_not_slide_text() {
    let doc = styled_deck("").open();
    assert_eq!(doc.layouts.len(), 1);
    assert_eq!(doc.masters.len(), 1);
    assert_eq!(doc.layouts[0].name, "Title and Content");
    assert_eq!(doc.slides[0].layout_index, Some(0));
    assert_eq!(doc.layouts[0].master_index, Some(0));
    assert_eq!(doc.static_text_for_slide(0), vec!["ACME CONFIDENTIAL", "MASTER FOOTER"]);
    assert!(!doc.slide_plain_text(0).unwrap().contains("ACME"), "not slide text");
    assert_eq!(doc.plain_text().matches("ACME").count(), 1, "once per document");
    assert_eq!(doc.master_static_text(), vec!["ACME CONFIDENTIAL", "MASTER FOOTER"]);
    // The layout's prompt text is not static text.
    assert!(
        !doc.static_text_for_slide(0)
            .iter()
            .any(|t| t.contains("Click"))
    );

    let hidden = styled_deck(r#"showMasterSp="0""#).open();
    assert!(hidden.static_text_for_slide(0).is_empty());
}

// ---------------------------------------------------------------------------
// Package paths the test suite never executed
// ---------------------------------------------------------------------------

/// An embedded picture and an embedded chart are pre-loaded from the
/// slide's relationships before the (parallel) slide parse; neither loop
/// was exercised by any test. An image target without an extension is
/// typed by sniffing its bytes.
#[test]
fn test_embedded_picture_and_chart_parts_are_loaded_through_the_slide_rels() {
    let chart_frame = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="3000"/><a:ext cx="100" cy="100"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdC"/></a:graphicData></a:graphic></p:graphicFrame>"#;
    let tree = format!("{}{chart_frame}", picture("", r#"<a:blip r:embed="rIdI"/>"#, "logo"));
    let mut pkg = deck(&[&tree]);
    pkg.part("ppt/media/image1", "", PNG);
    pkg.rel("ppt/slides/slide1.xml", "rIdI", rel_types::IMAGE, "../media/image1");
    pkg.part(
        "ppt/charts/chart1.xml",
        "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Quarterly revenue</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>"#,
    );
    pkg.rel("ppt/slides/slide1.xml", "rIdC", rel_types::CHART, "../charts/chart1.xml");

    let doc = pkg.open();
    let office_oxide::pptx::Shape::Picture(ref pic) = doc.slides[0].shapes[0] else {
        panic!("expected a picture");
    };
    assert_eq!(pic.data.as_deref(), Some(PNG));
    assert_eq!(pic.format.as_deref(), Some("png"), "sniffed from the bytes");
    let ir = pkg.document().to_ir();
    assert_eq!(images(&ir)[0].format, Some(ImageFormat::Png));
    assert!(ir.plain_text().contains("Quarterly revenue"), "{}", ir.plain_text());
}

/// Font programs under `ppt/fonts/` are collected by name.
#[test]
fn test_embedded_fonts_are_collected() {
    let mut pkg = deck(&[&text_sp("X")]);
    pkg.part("ppt/fonts/font_1_Garamond.ttf", "", b"\0\x01\0\0ttf".to_vec());
    pkg.part("ppt/fonts/Futura.OTF", "", b"OTTOotf".to_vec());
    pkg.part("ppt/fonts/font1.fntdata", "", b"obfuscated".to_vec());
    let doc = pkg.open();
    let mut fonts: Vec<(String, Vec<u8>)> = doc.embedded_fonts.clone();
    fonts.sort();
    assert_eq!(
        fonts,
        vec![
            ("Futura".to_string(), b"OTTOotf".to_vec()),
            ("Garamond".to_string(), b"\0\x01\0\0ttf".to_vec()),
        ]
    );
}

/// A `<p:sldId>` without an `r:id` (seen in some producers' output) falls
/// back to the `ppt/slides/slideN.xml` naming convention.
#[test]
fn test_a_slide_id_without_r_id_resolves_by_convention() {
    let mut pkg = deck(&[&text_sp("BY CONVENTION")]);
    let pres = String::from_utf8(pkg.parts["ppt/presentation.xml"].clone()).unwrap();
    pkg.parts
        .insert("ppt/presentation.xml".into(), pres.replace(r#" r:id="rId1""#, "").into_bytes());
    let doc = pkg.open();
    assert_eq!(doc.slides.len(), 1);
    assert!(doc.plain_text().contains("BY CONVENTION"));
}

// ---------------------------------------------------------------------------
// Master/layout static text in the extracted output
// ---------------------------------------------------------------------------

const PROMPT: &str = "Click to edit Master title style";

/// A deck whose slides (one per entry of `slide_attrs`, the `<p:sld>`
/// attributes) all use one layout (`<p:sldLayout>` attributes
/// `layout_attrs`) on one master. The master and the layout each hold a
/// placeholder prompt and a static text box, and share one identical line.
fn master_text_deck(slide_attrs: &[&str], layout_attrs: &str) -> Pkg {
    let own: Vec<String> = (1..=slide_attrs.len())
        .map(|n| text_sp(&format!("OWN TEXT {n}")))
        .collect();
    let trees: Vec<&str> = own.iter().map(String::as_str).collect();
    let mut pkg = deck(&trees);
    for (i, attrs) in slide_attrs.iter().enumerate() {
        let n = i + 1;
        pkg.parts.insert(
            format!("ppt/slides/slide{n}.xml"),
            format!(
                r#"<?xml version="1.0"?><p:sld {NS} {attrs}><p:cSld><p:spTree>{}</p:spTree></p:cSld></p:sld>"#,
                own[i]
            )
            .into_bytes(),
        );
        pkg.rel(
            &format!("ppt/slides/slide{n}.xml"),
            "rIdLay",
            rel_types::SLIDE_LAYOUT,
            "../slideLayouts/slideLayout1.xml",
        );
    }
    let layout_tree = [
        ph_sp(r#"<p:ph type="title"/>"#, "", &format!("<a:p>{}</a:p>", run(PROMPT))),
        text_sp("LAYOUT STATIC"),
        text_sp("SHARED LINE"),
    ]
    .concat();
    pkg.part(
        "ppt/slideLayouts/slideLayout1.xml",
        &format!("{CT_PML}slideLayout+xml"),
        format!(
            r#"<?xml version="1.0"?><p:sldLayout {NS} {layout_attrs}><p:cSld><p:spTree>{layout_tree}</p:spTree></p:cSld></p:sldLayout>"#
        ),
    );
    pkg.rel(
        "ppt/slideLayouts/slideLayout1.xml",
        "rIdM",
        rel_types::SLIDE_MASTER,
        "../slideMasters/slideMaster1.xml",
    );
    let master_tree = [
        ph_sp(r#"<p:ph type="title"/>"#, "", &format!("<a:p>{}</a:p>", run(PROMPT))),
        ph_sp(
            r#"<p:ph type="body" idx="1"/>"#,
            "",
            &format!("<a:p>{}</a:p>", run("Second level")),
        ),
        text_sp("MASTER STATIC"),
        text_sp("SHARED LINE"),
    ]
    .concat();
    pkg.part(
        "ppt/slideMasters/slideMaster1.xml",
        &format!("{CT_PML}slideMaster+xml"),
        format!(
            r#"<?xml version="1.0"?><p:sldMaster {NS}><p:cSld><p:spTree>{master_tree}</p:spTree></p:cSld></p:sldMaster>"#
        ),
    );
    pkg
}

/// Every text surface of a document.
fn text_surfaces(doc: &Document) -> [(&'static str, String); 5] {
    let ir = doc.to_ir();
    [
        ("plain_text", doc.plain_text()),
        ("to_markdown", doc.to_markdown()),
        ("ir.plain_text", ir.plain_text()),
        ("ir.to_markdown", ir.to_markdown()),
        ("to_html", doc.to_html()),
    ]
}

/// Static text on the layout and master is drawn on every slide using
/// them, and reference extractors (Apache POI/Tika) include it. It is
/// surfaced once per deck — never once per slide — in a trailing
/// "Slide Master" section, identically on every surface; an identical
/// line on both the layout and the master appears once. Placeholder
/// prompts on the layout and master never appear.
#[test]
fn test_pptx_master_and_layout_static_text_appear_once_and_prompts_never() {
    let doc = master_text_deck(&["", ""], "").document();
    for (name, text) in text_surfaces(&doc) {
        for line in ["LAYOUT STATIC", "MASTER STATIC", "SHARED LINE"] {
            assert_eq!(text.matches(line).count(), 1, "{name}: {line}\n{text}");
        }
        assert!(!text.contains(PROMPT), "{name}: {text}");
        assert!(!text.contains("Second level"), "{name}: {text}");
        assert!(text.contains("OWN TEXT 1") && text.contains("OWN TEXT 2"), "{name}: {text}");
    }
    let ir = doc.to_ir();
    assert_eq!(ir.sections.len(), 3, "two slides plus the master section");
    assert_eq!(ir.sections[2].title.as_deref(), Some("Slide Master"));
    // The direct and IR markdown render the master section identically.
    let tail = |md: String| {
        md[md.find("## Slide Master").unwrap()..]
            .trim_end()
            .to_string()
    };
    assert_eq!(tail(doc.to_markdown()), tail(ir.to_markdown()));
    assert_eq!(
        tail(doc.to_markdown()),
        "## Slide Master\n\nLAYOUT STATIC\n\nSHARED LINE\n\nMASTER STATIC"
    );
}

/// Slides that hide the master's shapes (`<p:sld showMasterSp="0">`,
/// ECMA-376 Part 1 §19.3.1.38) show neither layout nor master static text;
/// when every slide does, the deck has no master section.
#[test]
fn test_pptx_static_text_is_absent_when_every_slide_hides_master_shapes() {
    let doc = master_text_deck(&[r#"showMasterSp="0""#, r#"showMasterSp="0""#], "").document();
    for (name, text) in text_surfaces(&doc) {
        for line in [
            "LAYOUT STATIC",
            "MASTER STATIC",
            "SHARED LINE",
            "Slide Master",
        ] {
            assert!(!text.contains(line), "{name}: {line}\n{text}");
        }
    }
    assert_eq!(doc.to_ir().sections.len(), 2);
}

/// One slide showing the master's shapes is enough.
#[test]
fn test_pptx_static_text_is_kept_when_one_slide_shows_master_shapes() {
    let doc = master_text_deck(&[r#"showMasterSp="0""#, ""], "").document();
    for (name, text) in text_surfaces(&doc) {
        assert_eq!(text.matches("MASTER STATIC").count(), 1, "{name}: {text}");
    }
}

/// A layout with `showMasterSp="0"` keeps its own static text but hides
/// the master's.
#[test]
fn test_pptx_layout_hiding_master_shapes_keeps_only_layout_static_text() {
    let doc = master_text_deck(&["", ""], r#"showMasterSp="0""#).document();
    for (name, text) in text_surfaces(&doc) {
        assert_eq!(text.matches("LAYOUT STATIC").count(), 1, "{name}: {text}");
        assert_eq!(text.matches("SHARED LINE").count(), 1, "{name}: {text}");
        assert!(!text.contains("MASTER STATIC"), "{name}: {text}");
    }
}
