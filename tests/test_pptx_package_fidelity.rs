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
