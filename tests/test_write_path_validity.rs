//! Would the packages the write/create paths produce open cleanly? Each
//! test builds a document through the public writer or `create_from_*`
//! API and inspects the written parts: dangling relationship ids,
//! XML-illegal characters, elements an Office application requires.

use std::io::{Cursor, Read};

use office_oxide::DocumentFormat;
use office_oxide::create::create_from_ir_to_writer;
use office_oxide::ir::*;

fn write_ir(ir: &DocumentIR, format: DocumentFormat) -> Vec<u8> {
    let mut buf = Cursor::new(Vec::new());
    create_from_ir_to_writer(ir, format, &mut buf).expect("write");
    buf.into_inner()
}

fn part(bytes: &[u8], name: &str) -> Option<String> {
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut entry = zip.by_name(name).ok()?;
    let mut s = String::new();
    entry.read_to_string(&mut s).unwrap();
    Some(s)
}

fn part_names(bytes: &[u8]) -> Vec<String> {
    let zip = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    zip.file_names().map(str::to_string).collect()
}

/// Every `r:id="…"` / `r:embed="…"` value in `xml`.
fn rel_ids(xml: &str) -> Vec<String> {
    let mut out = Vec::new();
    for key in ["r:id=\"", "r:embed=\"", "r:link=\""] {
        let mut rest = xml;
        while let Some(at) = rest.find(key) {
            rest = &rest[at + key.len()..];
            let end = rest.find('"').unwrap();
            out.push(rest[..end].to_string());
            rest = &rest[end..];
        }
    }
    out
}

/// Assert every relationship id used in each XML part resolves in that
/// part's own `.rels` — the defect class PowerPoint "repairs" on open.
fn assert_no_dangling_rel_ids(bytes: &[u8]) {
    for name in part_names(bytes) {
        if !name.ends_with(".xml") || name.contains("_rels/") {
            continue;
        }
        let xml = part(bytes, &name).unwrap();
        let ids = rel_ids(&xml);
        if ids.is_empty() {
            continue;
        }
        let (dir, file) = name.rsplit_once('/').unwrap_or(("", &name));
        let rels_name = if dir.is_empty() {
            format!("_rels/{file}.rels")
        } else {
            format!("{dir}/_rels/{file}.rels")
        };
        let rels = part(bytes, &rels_name).unwrap_or_default();
        for id in ids {
            assert!(
                rels.contains(&format!("Id=\"{id}\"")),
                "{name} uses {id}, which {rels_name} does not define:\n{rels}"
            );
        }
    }
}

fn linked(text: &str, url: &str) -> InlineContent {
    InlineContent::Text(TextSpan {
        text: text.to_string(),
        hyperlink: Some(url.to_string()),
        ..Default::default()
    })
}

fn para(content: Vec<InlineContent>) -> Element {
    Element::Paragraph(Paragraph {
        content,
        ..Default::default()
    })
}

fn one_section(elements: Vec<Element>, format: DocumentFormat) -> DocumentIR {
    DocumentIR {
        metadata: Metadata {
            format,
            ..Default::default()
        },
        sections: vec![Section {
            elements,
            ..Default::default()
        }],
        defined_names: Vec::new(),
    }
}

// ---------------------------------------------------------------------------
// PPTX speaker-note hyperlinks
// ---------------------------------------------------------------------------

/// A notes hyperlink was registered only on the *slide* part: a URL used
/// only in the notes was silently dropped, and one shared with the slide
/// body was written as an `r:id` the notes part's `.rels` does not define.
#[test]
fn test_pptx_notes_hyperlinks_resolve_in_the_notes_part() {
    let mut ir = one_section(
        vec![para(vec![linked(
            "body link",
            "https://example.com/shared",
        )])],
        DocumentFormat::Pptx,
    );
    ir.sections[0].speaker_notes = Some(vec![
        para(vec![linked("shared", "https://example.com/shared")]),
        para(vec![linked("notes only", "https://example.com/notes-only")]),
        Element::List(List {
            items: vec![ListItem {
                content: vec![para(vec![linked("listed", "https://example.com/listed")])],
                nested: None,
            }],
            ..Default::default()
        }),
    ]);
    let bytes = write_ir(&ir, DocumentFormat::Pptx);
    assert_no_dangling_rel_ids(&bytes);
    let rels = part(&bytes, "ppt/notesSlides/_rels/notesSlide1.xml.rels").unwrap();
    for url in [
        "https://example.com/shared",
        "https://example.com/notes-only",
        "https://example.com/listed",
    ] {
        assert!(rels.contains(url), "{url} missing from the notes rels: {rels}");
    }
    // And the crate's own reader resolves them back.
    let doc =
        office_oxide::Document::from_reader(Cursor::new(bytes), DocumentFormat::Pptx).unwrap();
    let back = doc.to_ir();
    let notes = format!("{:?}", back.sections[0].speaker_notes);
    assert!(notes.contains("https://example.com/notes-only"), "{notes}");
}

// ---------------------------------------------------------------------------
// PPTX ordered lists
// ---------------------------------------------------------------------------

fn find_list(elements: &[Element]) -> Option<&List> {
    elements.iter().find_map(|e| match e {
        Element::List(l) => Some(l),
        Element::TextBox(tb) => find_list(&tb.content),
        _ => None,
    })
}

/// Every IR list — ordered or not — was written with
/// `<a:buChar char="•"/>`: a numbered list came out of PPTX as bullets.
#[test]
fn test_pptx_ordered_list_is_written_with_auto_numbering() {
    let item = |t: &str, nested: Option<List>| ListItem {
        content: vec![para(vec![InlineContent::Text(TextSpan::plain(t))])],
        nested,
    };
    let list = List {
        ordered: true,
        start_number: Some(3),
        style: Some(ListStyle::LowerAlpha),
        items: vec![
            item(
                "third",
                Some(List {
                    items: vec![item("a bullet", None)],
                    ..Default::default()
                }),
            ),
            item("fourth", None),
        ],
        ..Default::default()
    };
    let mut ir = one_section(vec![Element::List(list.clone())], DocumentFormat::Pptx);
    ir.sections[0].speaker_notes = Some(vec![Element::List(list)]);
    let bytes = write_ir(&ir, DocumentFormat::Pptx);
    for name in ["ppt/slides/slide1.xml", "ppt/notesSlides/notesSlide1.xml"] {
        let xml = part(&bytes, name).unwrap();
        assert_eq!(
            xml.matches(r#"<a:buAutoNum type="alphaLcPeriod" startAt="3"/>"#)
                .count(),
            2,
            "{name}: {xml}"
        );
        assert_eq!(xml.matches("<a:buChar").count(), 1, "the nested list stays bulleted: {xml}");
    }
    let doc =
        office_oxide::Document::from_reader(Cursor::new(bytes), DocumentFormat::Pptx).unwrap();
    let back = doc.to_ir();
    let l = find_list(&back.sections[0].elements).expect("a list");
    assert!(l.ordered);
    assert_eq!(l.start_number, Some(3));
    assert_eq!(l.style, Some(ListStyle::LowerAlpha));
}

// ---------------------------------------------------------------------------
// XML-illegal characters
// ---------------------------------------------------------------------------

/// Assert no XML part contains a character XML 1.0 cannot represent
/// (C0 controls other than tab, LF and CR) — raw or as a character
/// reference.
fn assert_all_parts_are_legal_xml(bytes: &[u8]) {
    for name in part_names(bytes) {
        if !(name.ends_with(".xml") || name.ends_with(".rels") || name.ends_with(".vml")) {
            continue;
        }
        let xml = part(bytes, &name).unwrap();
        let bad: Vec<char> = xml
            .chars()
            .filter(|&c| (c as u32) < 0x20 && !matches!(c, '\t' | '\n' | '\r'))
            .collect();
        assert!(bad.is_empty(), "{name} holds XML-illegal characters {bad:?}: {xml}");
        assert!(!xml.contains("&#x1;") && !xml.contains("&#1;"), "{name}: {xml}");
    }
}

/// Comment author/text went through an escaper that handled `&<>` but not
/// control characters, and relationship targets were written verbatim:
/// `save()` returned Ok for a package with non-well-formed parts.
#[test]
fn test_xlsx_comments_and_hyperlink_targets_carry_no_illegal_characters() {
    let mut wb = office_oxide::xlsx::write::XlsxWriter::new();
    {
        let mut sheet = wb.add_sheet("S");
        sheet.set_cell(0, 0, office_oxide::xlsx::write::CellData::String("cell".into()));
        sheet.set_cell_comment(0, 0, Some("Ann\u{0B}Lee".into()), "see\u{01} A & <B>");
        sheet.set_cell_hyperlink(0, 0, "https://example.com/a\u{01}b");
    }
    let mut buf = Cursor::new(Vec::new());
    wb.write_to(&mut buf).unwrap();
    let bytes = buf.into_inner();
    assert_all_parts_are_legal_xml(&bytes);

    let doc = office_oxide::Document::from_reader(Cursor::new(bytes.clone()), DocumentFormat::Xlsx)
        .unwrap();
    let text = doc.to_ir().plain_text();
    assert!(text.contains("see A") && text.contains("<B>"), "{text}");
    let comments = part(&bytes, "xl/comments1.xml").unwrap();
    assert!(comments.contains("<author>AnnLee</author>"), "{comments}");
    assert!(comments.contains("see A &amp; &lt;B&gt;"), "{comments}");
    // The control character is percent-encoded in the target, as RFC 3986
    // requires for a URI, rather than silently deleted.
    let rels = part(&bytes, "xl/worksheets/_rels/sheet1.xml.rels").unwrap();
    assert!(rels.contains("https://example.com/a%01b"), "{rels}");
}

/// The same relationship-target rule holds for every format's hyperlinks.
#[test]
fn test_hyperlink_targets_with_control_characters_are_percent_encoded_in_every_format() {
    for format in [
        DocumentFormat::Docx,
        DocumentFormat::Pptx,
        DocumentFormat::Xlsx,
    ] {
        let ir = one_section(
            vec![Element::Table(Table {
                rows: vec![TableRow {
                    cells: vec![TableCell {
                        content: vec![para(vec![linked("x", "https://example.com/\u{02}q")])],
                        col_span: 1,
                        row_span: 1,
                        ..Default::default()
                    }],
                    ..Default::default()
                }],
                ..Default::default()
            })],
            format,
        );
        let bytes = write_ir(&ir, format);
        assert_all_parts_are_legal_xml(&bytes);
    }
}

// ---------------------------------------------------------------------------
// PPTX notes master declaration
// ---------------------------------------------------------------------------

/// A deck with notes relates a notes master from the presentation part;
/// PowerPoint-authored decks also declare it in `<p:notesMasterIdLst>`,
/// which the writer never emitted.
#[test]
fn test_pptx_with_notes_declares_its_notes_master() {
    let mut w = office_oxide::pptx::write::PptxWriter::new();
    w.add_slide().add_text("body").set_notes("a note");
    let mut buf = Cursor::new(Vec::new());
    w.write_to(&mut buf).unwrap();
    let bytes = buf.into_inner();
    let pres = part(&bytes, "ppt/presentation.xml").unwrap();
    let rels = part(&bytes, "ppt/_rels/presentation.xml.rels").unwrap();
    let key = r#"<p:notesMasterId r:id=""#;
    let at = pres
        .find(key)
        .unwrap_or_else(|| panic!("no notesMasterIdLst: {pres}"));
    let rid = &pres[at + key.len()..];
    let rid = &rid[..rid.find('"').unwrap()];
    let rel = rels
        .split("<Relationship ")
        .find(|r| r.contains(&format!(r#"Id="{rid}""#)))
        .unwrap_or_else(|| panic!("{rid} not in {rels}"));
    assert!(rel.contains("/notesMaster\""), "{rel}");
    // Schema order: after the slide masters, before the slides.
    let (masters, notes_master, slides) = (
        pres.find("<p:sldMasterIdLst>").unwrap(),
        pres.find("<p:notesMasterIdLst>").unwrap(),
        pres.find("<p:sldIdLst>").unwrap(),
    );
    assert!(masters < notes_master && notes_master < slides, "{pres}");
    assert_no_dangling_rel_ids(&bytes);

    // Without notes there is no notes master to declare.
    let mut w = office_oxide::pptx::write::PptxWriter::new();
    w.add_slide().add_text("body");
    let mut buf = Cursor::new(Vec::new());
    w.write_to(&mut buf).unwrap();
    let pres = part(&buf.into_inner(), "ppt/presentation.xml").unwrap();
    assert!(!pres.contains("notesMaster"), "{pres}");
}
