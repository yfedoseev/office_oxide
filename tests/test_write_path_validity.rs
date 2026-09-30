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
