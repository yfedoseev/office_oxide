#![no_main]

// A `.docx` or `.pptx` part that is malformed or cut off part-way keeps the
// blocks before the damage (`core::xml::salvage_prefix`): the body, a header,
// the footnotes, a slide, its notes and a layout. Whole-file fuzzing rarely
// reaches those paths, because it needs a valid package around the damaged
// part. This target builds a well-formed package in code and hands the
// fuzzer's bytes in as one part, after a valid opening, so every run tries to
// read, fail and salvage. None may panic, overflow or hang.
use std::io::Cursor;

use libfuzzer_sys::fuzz_target;
use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::rel_types;
use office_oxide::{Document, DocumentFormat};

const WML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT_DOCX: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_PPTX: &str = "application/vnd.openxmlformats-officedocument.presentationml";

fn part(opc: &mut OpcWriter<Cursor<Vec<u8>>>, name: &str, ct: &str, data: &[u8]) -> PartName {
    let name = PartName::new(name).unwrap();
    opc.add_part(&name, ct, data).unwrap();
    name
}

fn docx(which: u8, fuzz: &[u8]) -> Vec<u8> {
    let mut opc = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let ok_body = format!(
        r#"<w:document xmlns:w="{WML}" xmlns:r="{REL}"><w:body><w:p><w:r><w:t>kept</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr></w:body></w:document>"#
    );
    let mut damaged = format!(r#"<w:document xmlns:w="{WML}"><w:body><w:p><w:r><w:t>kept</w:t></w:r></w:p>"#)
        .into_bytes();
    damaged.extend_from_slice(fuzz);
    let main = part(
        &mut opc,
        "/word/document.xml",
        CT_DOCX,
        if which == 0 { &damaged } else { ok_body.as_bytes() },
    );
    opc.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    let mut with_head = |root: &str, open: &str, ct: &str, name: &str, rel: &str, target: &str| {
        let mut bytes = format!(r#"<w:{root} xmlns:w="{WML}">{open}"#).into_bytes();
        bytes.extend_from_slice(fuzz);
        part(&mut opc, name, ct, &bytes);
        opc.add_part_rel(&main, rel, target);
    };
    match which {
        1 => with_head(
            "hdr",
            "<w:p><w:r><w:t>head</w:t></w:r></w:p>",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
            "/word/header1.xml",
            rel_types::HEADER,
            "header1.xml",
        ),
        2 => with_head(
            "footnotes",
            r#"<w:footnote w:id="1"><w:p><w:r><w:t>note</w:t></w:r></w:p>"#,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
            "/word/footnotes.xml",
            rel_types::FOOTNOTES,
            "footnotes.xml",
        ),
        _ => {},
    }
    opc.finish().unwrap().into_inner()
}

fn pptx(which: u8, fuzz: &[u8]) -> Vec<u8> {
    let mut opc = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let pres = part(
        &mut opc,
        "/ppt/presentation.xml",
        &format!("{CT_PPTX}.presentation.main+xml"),
        format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#
        )
        .as_bytes(),
    );
    opc.add_package_rel(rel_types::OFFICE_DOCUMENT, "ppt/presentation.xml");
    let head = |root: &str| {
        format!(
            r#"<p:{root} xmlns:p="{PML}" xmlns:a="{DML}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="S"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>kept</a:t></a:r></a:p></p:txBody></p:sp>"#
        )
        .into_bytes()
    };
    let tail = |root: &str| format!("</p:spTree></p:cSld></p:{root}>").into_bytes();
    let broken = |root: &str| {
        let mut b = head(root);
        b.extend_from_slice(fuzz);
        b
    };
    let slide_bytes = if which == 3 { broken("sld") } else { [head("sld"), tail("sld")].concat() };
    let slide = part(&mut opc, "/ppt/slides/slide1.xml", &format!("{CT_PPTX}.slide+xml"), &slide_bytes);
    opc.add_part_rel(&pres, rel_types::SLIDE, "slides/slide1.xml");
    if which == 4 {
        part(&mut opc, "/ppt/notesSlides/notesSlide1.xml", &format!("{CT_PPTX}.notesSlide+xml"), &broken("notes"));
        opc.add_part_rel(&slide, rel_types::NOTES_SLIDE, "../notesSlides/notesSlide1.xml");
    }
    if which == 5 {
        part(&mut opc, "/ppt/slideLayouts/slideLayout1.xml", &format!("{CT_PPTX}.slideLayout+xml"), &broken("sldLayout"));
        opc.add_part_rel(&slide, rel_types::SLIDE_LAYOUT, "../slideLayouts/slideLayout1.xml");
    }
    opc.finish().unwrap().into_inner()
}

fuzz_target!(|data: &[u8]| {
    let Some((&which, fuzz)) = data.split_first() else {
        return;
    };
    let which = which % 6;
    let (bytes, format) = if which < 3 {
        (docx(which, fuzz), DocumentFormat::Docx)
    } else {
        (pptx(which, fuzz), DocumentFormat::Pptx)
    };
    if let Ok(doc) = Document::from_reader(Cursor::new(bytes), format) {
        let _ = doc.plain_text();
        let _ = doc.to_markdown();
        let _ = doc.to_ir();
    }
});
