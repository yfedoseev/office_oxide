//! Conversion fidelity for DOCX → `DocumentIR`.
//!
//! Every test here builds a minimal DOCX **in code** (AGENTS.md rule #4 —
//! no committed third-party fixtures) and asserts that a construct the
//! parser already understood actually reaches the IR. The defects these
//! guard against were all of the "parsed then discarded" shape: the parser
//! read the value correctly and the converter dropped it, so a consumer saw
//! a plausible-looking document with the formatting silently missing.

use std::io::Cursor;

use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::rel_types;
use office_oxide::ir::*;
use office_oxide::{Document, DocumentFormat};

// ---------------------------------------------------------------------------
// Fixture builder
// ---------------------------------------------------------------------------

const CT_DOC: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_STYLES: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const CT_NUMBERING: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const CT_CORE: &str = "application/vnd.openxmlformats-package.core-properties+xml";
const CT_HF: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const CT_FOOTNOTES: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
const CT_COMMENTS: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const CT_HTML: &str = "text/html";
const REL_ALT_CHUNK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";

struct Docx {
    w: OpcWriter<Cursor<Vec<u8>>>,
    doc_part: PartName,
    body: String,
}

impl Docx {
    fn new(body: &str) -> Self {
        let mut w = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
        let doc_part = PartName::new("/word/document.xml").unwrap();
        w.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
        Self {
            w,
            doc_part,
            body: body.to_string(),
        }
    }

    fn styles(mut self, inner: &str) -> Self {
        let part = PartName::new("/word/styles.xml").unwrap();
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{inner}</w:styles>"#
        );
        self.w.add_part(&part, CT_STYLES, xml.as_bytes()).unwrap();
        self.w
            .add_part_rel(&self.doc_part, rel_types::STYLES, "styles.xml");
        self
    }

    fn numbering(mut self, inner: &str) -> Self {
        let part = PartName::new("/word/numbering.xml").unwrap();
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{inner}</w:numbering>"#
        );
        self.w
            .add_part(&part, CT_NUMBERING, xml.as_bytes())
            .unwrap();
        self.w
            .add_part_rel(&self.doc_part, rel_types::NUMBERING, "numbering.xml");
        self
    }

    /// Add a header or footer part. The relationship id is assigned by the
    /// OPC writer, so the body XML uses the `placeholder` token and this
    /// substitutes the real id into it.
    fn hf(mut self, file: &str, placeholder: &str, text: &str) -> Self {
        let tag = if file.starts_with("header") {
            "hdr"
        } else {
            "ftr"
        };
        let part = PartName::new(&format!("/word/{file}")).unwrap();
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<w:{tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>{text}</w:t></w:r></w:p>
</w:{tag}>"#
        );
        self.w.add_part(&part, CT_HF, xml.as_bytes()).unwrap();
        let rel_type = if tag == "hdr" {
            rel_types::HEADER
        } else {
            rel_types::FOOTER
        };
        let rid = self.w.add_part_rel(&self.doc_part, rel_type, file);
        self.body = self.body.replace(placeholder, &rid);
        self
    }

    /// Add a header part with the given raw bytes — for a part that is
    /// not well-formed XML.
    fn header_raw(mut self, placeholder: &str, xml: &[u8]) -> Self {
        let part = PartName::new("/word/header1.xml").unwrap();
        self.w.add_part(&part, CT_HF, xml).unwrap();
        let rid = self
            .w
            .add_part_rel(&self.doc_part, rel_types::HEADER, "header1.xml");
        self.body = self.body.replace(placeholder, &rid);
        self
    }

    /// Add a header part carrying a picture: the image part and its
    /// relationship live on the *header*, under an id assigned by the
    /// header's own rels file.
    fn header_with_picture(mut self, placeholder: &str, png: &[u8]) -> Self {
        let img_part = PartName::new("/word/media/logo.png").unwrap();
        self.w.add_part(&img_part, "image/png", png).unwrap();
        let part = PartName::new("/word/header1.xml").unwrap();
        let img_rid = self
            .w
            .add_part_rel(&part, rel_types::IMAGE, "media/logo.png");
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p><w:r><w:t>HEADER TEXT</w:t></w:r><w:r><w:drawing>
    <wp:inline xmlns:wp="x">
      <wp:extent cx="100" cy="100"/>
      <wp:docPr id="1" name="Pic" descr="HEADER LOGO"/>
      <a:graphic xmlns:a="y"><a:graphicData>
        <pic:pic xmlns:pic="p"><pic:blipFill><a:blip r:embed="{img_rid}"/></pic:blipFill></pic:pic>
      </a:graphicData></a:graphic>
    </wp:inline>
  </w:drawing></w:r></w:p>
</w:hdr>"#
        );
        self.w.add_part(&part, CT_HF, xml.as_bytes()).unwrap();
        let rid = self
            .w
            .add_part_rel(&self.doc_part, rel_types::HEADER, "header1.xml");
        self.body = self.body.replace(placeholder, &rid);
        self
    }

    /// Add `word/footnotes.xml` / `endnotes.xml` / `comments.xml` with the
    /// given item elements already written out.
    fn notes(mut self, file: &str, rel_type: &str, ct: &str, inner: &str) -> Self {
        let root = file.trim_end_matches(".xml");
        let part = PartName::new(&format!("/word/{file}")).unwrap();
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<w:{root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{inner}</w:{root}>"#
        );
        self.w.add_part(&part, ct, xml.as_bytes()).unwrap();
        self.w.add_part_rel(&self.doc_part, rel_type, file);
        self
    }

    /// Add an `altChunk` target part and substitute its relationship id
    /// into the body's `placeholder` token.
    fn alt_chunk(mut self, file: &str, ct: &str, placeholder: &str, body: &str) -> Self {
        let part = PartName::new(&format!("/word/{file}")).unwrap();
        self.w.add_part(&part, ct, body.as_bytes()).unwrap();
        let rid = self.w.add_part_rel(&self.doc_part, REL_ALT_CHUNK, file);
        self.body = self.body.replace(placeholder, &rid);
        self
    }

    /// Add an external relationship from the document part and substitute
    /// its id into the body's `placeholder` token.
    fn external_rel(mut self, placeholder: &str, rel_type: &str, target: &str) -> Self {
        let rid = self.w.add_part_rel_with_mode(
            &self.doc_part,
            rel_type,
            target,
            office_oxide::core::relationships::TargetMode::External,
        );
        self.body = self.body.replace(placeholder, &rid);
        self
    }

    /// A theme part whose font scheme has `major` / `minor` Latin faces
    /// and a `minor` East Asian face `minor_ea`.
    fn theme_fonts(mut self, major: &str, minor: &str, minor_ea: &str) -> Self {
        let part = PartName::new("/word/theme/theme1.xml").unwrap();
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements>
<a:clrScheme name="C"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1></a:clrScheme>
<a:fontScheme name="F">
<a:majorFont><a:latin typeface="{major}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="{minor}"/><a:ea typeface="{minor_ea}"/><a:cs typeface=""/></a:minorFont>
</a:fontScheme></a:themeElements></a:theme>"#
        );
        self.w
            .add_part(
                &part,
                "application/vnd.openxmlformats-officedocument.theme+xml",
                xml.as_bytes(),
            )
            .unwrap();
        self.w
            .add_part_rel(&self.doc_part, rel_types::THEME, "theme/theme1.xml");
        self
    }

    fn core_props(mut self, inner: &str) -> Self {
        let part = PartName::new("/docProps/core.xml").unwrap();
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/">{inner}</cp:coreProperties>"#
        );
        self.w.add_part(&part, CT_CORE, xml.as_bytes()).unwrap();
        self.w
            .add_package_rel(rel_types::CORE_PROPERTIES, "docProps/core.xml");
        self
    }

    fn ir(self) -> DocumentIR {
        Document::from_reader(Cursor::new(self.bytes()), DocumentFormat::Docx)
            .expect("parse")
            .to_ir()
    }

    fn bytes(mut self) -> Vec<u8> {
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>{}</w:body>
</w:document>"#,
            self.body
        );
        self.w
            .add_part(&self.doc_part, CT_DOC, xml.as_bytes())
            .unwrap();
        self.w.finish().unwrap().into_inner()
    }
}

/// First element of the first section.
fn first(ir: &DocumentIR) -> &Element {
    &ir.sections[0].elements[0]
}

fn para(ir: &DocumentIR, idx: usize) -> &Paragraph {
    match &ir.sections[0].elements[idx] {
        Element::Paragraph(p) => p,
        other => panic!("element {idx} is not a paragraph: {other:?}"),
    }
}

fn first_span(p: &Paragraph) -> &TextSpan {
    match &p.content[0] {
        InlineContent::Text(s) => s,
        other => panic!("not a text span: {other:?}"),
    }
}

fn table(ir: &DocumentIR) -> &Table {
    match first(ir) {
        Element::Table(t) => t,
        other => panic!("not a table: {other:?}"),
    }
}

// ---------------------------------------------------------------------------
// Run properties beyond bold/italic
// ---------------------------------------------------------------------------

#[test]
fn test_run_underline_reaches_the_ir() {
    let ir =
        Docx::new(r#"<w:p><w:r><w:rPr><w:u w:val="double"/></w:rPr><w:t>x</w:t></w:r></w:p>"#).ir();
    assert_eq!(first_span(para(&ir, 0)).underline, Some(UnderlineStyle::Double));
}

/// Write `ir` back out as a `.docx`.
fn docx_bytes(ir: &DocumentIR) -> Vec<u8> {
    let mut out = Cursor::new(Vec::new());
    office_oxide::create::create_from_ir_to_writer(ir, DocumentFormat::Docx, &mut out)
        .expect("write");
    out.into_inner()
}

/// `w:hyperlink/@w:tooltip` (ECMA-376 §17.16) and a `HYPERLINK`
/// field's `\o` switch (§17.16.5) are the link's hover text. Both were
/// parsed (or skipped) and never reached the IR, any renderer or the
/// writer.
#[test]
fn test_hyperlink_tooltips_reach_the_ir_html_and_the_writer() {
    let ir = Docx::new(
        r#"<w:p><w:hyperlink w:anchor="intro" w:tooltip="Jump to the intro"><w:r><w:t>Intro</w:t></w:r></w:hyperlink></w:p>
           <w:p><w:fldSimple w:instr=" HYPERLINK \l &quot;end&quot; \o &quot;Go to the end&quot; "><w:r><w:t>End</w:t></w:r></w:fldSimple></w:p>"#,
    )
    .ir();
    let tip = |ir: &DocumentIR, i: usize| first_span(para(ir, i)).hyperlink_tooltip.clone();
    assert_eq!(tip(&ir, 0).as_deref(), Some("Jump to the intro"));
    assert_eq!(tip(&ir, 1).as_deref(), Some("Go to the end"));
    let html = ir.to_html();
    assert!(html.contains(r##"<a href="#intro" title="Jump to the intro">"##), "{html}");

    let again = Document::from_reader(Cursor::new(docx_bytes(&ir)), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    assert_eq!(tip(&again, 0).as_deref(), Some("Jump to the intro"));
    assert_eq!(tip(&again, 1).as_deref(), Some("Go to the end"));
}

fn spans(p: &Paragraph) -> Vec<&TextSpan> {
    p.content
        .iter()
        .filter_map(|c| match c {
            InlineContent::Text(s) => Some(s),
            _ => None,
        })
        .collect()
}

/// `w:rFonts` names four faces (ECMA-376 §17.3.2.26) and Word picks one
/// per character: `w:ascii` for Basic Latin, `w:eastAsia` for CJK,
/// `w:cs` for complex scripts (and for every character of a `w:rtl` or
/// `w:cs` run), `w:hAnsi` for the rest. Complex-script text also takes
/// its size from `w:szCs` (§17.3.2) and bold/italic from `w:bCs`/`w:iCs`
/// (§17.3.2). Only `w:ascii`/`w:sz`/`w:b` were read, so
/// CJK and Arabic runs reported the Latin face and size.
#[test]
fn test_run_fonts_and_sizes_follow_the_script_of_the_text() {
    let rpr = r#"<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Cambria" w:eastAsia="MS Mincho" w:cs="Arial"/>
                 <w:b/><w:sz w:val="22"/><w:szCs w:val="28"/></w:rPr>"#;
    let ir = Docx::new(&format!(
        r#"<w:p><w:r>{rpr}<w:t>Tokyo</w:t></w:r></w:p>
           <w:p><w:r>{rpr}<w:t>東京</w:t></w:r></w:p>
           <w:p><w:r>{rpr}<w:t>مرحبا</w:t></w:r></w:p>
           <w:p><w:r>{rpr}<w:t>café</w:t></w:r></w:p>
           <w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:cs="Arial"/><w:bCs/><w:rtl/>
               <w:sz w:val="22"/><w:szCs w:val="28"/></w:rPr><w:t>abc</w:t></w:r></w:p>
           <w:p><w:r>{rpr}<w:t>東京 2024</w:t></w:r></w:p>"#
    ))
    .ir();
    let latin = first_span(para(&ir, 0));
    assert_eq!(latin.font_name.as_deref(), Some("Calibri"));
    assert_eq!(latin.font_size_half_pt, Some(22));
    assert!(latin.bold);

    let cjk = first_span(para(&ir, 1));
    assert_eq!(cjk.font_name.as_deref(), Some("MS Mincho"));
    assert_eq!(cjk.font_size_half_pt, Some(22));
    assert!(cjk.bold, "w:b applies to East Asian text");

    let arabic = first_span(para(&ir, 2));
    assert_eq!(arabic.font_name.as_deref(), Some("Arial"));
    assert_eq!(arabic.font_size_half_pt, Some(28), "complex script takes w:szCs");
    assert!(!arabic.bold, "w:b does not apply to complex script; w:bCs is absent");

    let accented = spans(para(&ir, 3));
    let faces: Vec<_> = accented
        .iter()
        .map(|s| (s.text.as_str(), s.font_name.as_deref()))
        .collect();
    assert_eq!(faces, [("caf", Some("Calibri")), ("é", Some("Cambria"))], "é is hAnsi");

    let rtl = first_span(para(&ir, 4));
    assert_eq!(rtl.font_name.as_deref(), Some("Arial"), "a w:rtl run is complex script");
    assert_eq!(rtl.font_size_half_pt, Some(28));
    assert!(rtl.bold, "w:bCs");

    // A mixed run splits where the face changes, and loses no text.
    let mixed = spans(para(&ir, 5));
    let text: String = mixed.iter().map(|s| s.text.as_str()).collect();
    assert_eq!(text, "東京 2024");
    assert_eq!(mixed[0].font_name.as_deref(), Some("MS Mincho"));
    assert_eq!(mixed.last().unwrap().font_name.as_deref(), Some("Calibri"));
}

#[test]
fn test_complex_script_bold_and_size_survive_a_docx_round_trip() {
    let ir = Docx::new(
        r#"<w:p><w:r><w:rPr><w:b/><w:bCs/><w:i/><w:iCs/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr>
             <w:t>مرحبا</w:t></w:r></w:p>"#,
    )
    .ir();
    let before = first_span(para(&ir, 0)).clone();
    assert!(before.bold && before.italic);
    let bytes = docx_bytes(&ir);
    let again = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .expect("reread")
        .to_ir();
    let after = first_span(para(&again, 0));
    assert_eq!((after.bold, after.italic), (true, true));
    assert_eq!(after.font_size_half_pt, Some(30));
}

/// `w:asciiTheme` names a font of the theme's font scheme and supersedes
/// the literal `w:ascii` face (ECMA-376 Part 1 §17.3.2.26); it is how
/// every Office template assigns fonts, and it was never parsed.
#[test]
fn test_theme_font_references_resolve_to_the_font_scheme_faces() {
    let ir = Docx::new(
        r#"<w:p><w:r><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:ascii="Ignored"/></w:rPr><w:t>heading</w:t></w:r></w:p>
           <w:p><w:r><w:rPr><w:rFonts w:hAnsiTheme="minorHAnsi"/></w:rPr><w:t>body</w:t></w:r></w:p>
           <w:p><w:r><w:rPr><w:rFonts w:asciiTheme="minorEastAsia"/></w:rPr><w:t>ea</w:t></w:r></w:p>
           <w:p><w:r><w:rPr><w:rFonts w:asciiTheme="majorBidi" w:ascii="Literal"/></w:rPr><w:t>cs</w:t></w:r></w:p>
           <w:p><w:r><w:rPr><w:rFonts w:ascii="Direct Face"/></w:rPr><w:t>direct</w:t></w:r></w:p>"#,
    )
    .theme_fonts("Heading Face", "Body Face", "EA Face")
    .ir();
    let font = |i: usize| first_span(para(&ir, i)).font_name.clone();
    assert_eq!(font(0).as_deref(), Some("Heading Face"));
    assert_eq!(font(1).as_deref(), Some("Body Face"));
    assert_eq!(font(2).as_deref(), Some("EA Face"));
    // An empty scheme slot falls back to the literal face.
    assert_eq!(font(3).as_deref(), Some("Literal"));
    assert_eq!(font(4).as_deref(), Some("Direct Face"));
}

/// A theme reference inherited from a style is replaced, not kept, by a
/// face named directly on the run; a theme reference on the run replaces
/// a style's direct face.
#[test]
fn test_theme_font_and_direct_face_override_each_other_as_a_unit() {
    let ir = Docx::new(
        r#"<w:p><w:r><w:rPr><w:rStyle w:val="Themed"/><w:rFonts w:ascii="Run Face"/></w:rPr><w:t>a</w:t></w:r></w:p>
           <w:p><w:r><w:rPr><w:rStyle w:val="Plain"/><w:rFonts w:asciiTheme="minorAscii"/></w:rPr><w:t>b</w:t></w:r></w:p>
           <w:p><w:r><w:rPr><w:rStyle w:val="Themed"/></w:rPr><w:t>c</w:t></w:r></w:p>"#,
    )
    .styles(
        r#"<w:style w:type="character" w:styleId="Themed"><w:rPr><w:rFonts w:asciiTheme="majorAscii"/></w:rPr></w:style>
           <w:style w:type="character" w:styleId="Plain"><w:rPr><w:rFonts w:ascii="Style Face"/></w:rPr></w:style>"#,
    )
    .theme_fonts("Heading Face", "Body Face", "")
    .ir();
    let font = |i: usize| first_span(para(&ir, i)).font_name.clone();
    assert_eq!(font(0).as_deref(), Some("Run Face"));
    assert_eq!(font(1).as_deref(), Some("Body Face"));
    assert_eq!(font(2).as_deref(), Some("Heading Face"));
}

#[test]
fn test_run_caps_smallcaps_and_spacing_reach_the_ir() {
    let ir = Docx::new(
        r#"<w:p><w:r><w:rPr><w:caps/><w:smallCaps/><w:spacing w:val="-20"/></w:rPr>
             <w:t>x</w:t></w:r></w:p>"#,
    )
    .ir();
    let s = first_span(para(&ir, 0));
    assert!(s.all_caps, "w:caps");
    assert!(s.small_caps, "w:smallCaps");
    assert_eq!(s.char_spacing_half_pt, Some(-20));
}

#[test]
fn test_run_vert_align_reaches_the_ir() {
    let ir = Docx::new(
        r#"<w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r></w:p>"#,
    )
    .ir();
    assert_eq!(first_span(para(&ir, 0)).vertical_align, Some(VerticalAlign::Superscript));
}

#[test]
fn test_run_highlight_reads_both_encodings() {
    // Word's named palette …
    let ir = Docx::new(
        r#"<w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>x</w:t></w:r></w:p>"#,
    )
    .ir();
    assert_eq!(first_span(para(&ir, 0)).highlight, Some([0xFF, 0xFF, 0x00]));

    // … and the `w:shd` form this crate's own writer emits.
    let ir = Docx::new(
        r#"<w:p><w:r><w:rPr><w:shd w:val="clear" w:fill="00FF00"/></w:rPr><w:t>x</w:t></w:r></w:p>"#,
    )
    .ir();
    assert_eq!(first_span(para(&ir, 0)).highlight, Some([0x00, 0xFF, 0x00]));
}

// ---------------------------------------------------------------------------
// Theme colours and the `w:val` fallback
// ---------------------------------------------------------------------------

#[test]
fn test_theme_color_falls_back_to_w_val_when_no_theme_part() {
    // `w:themeColor` with no theme part in the package: the literal `w:val`
    // is the producer-supplied fallback and must be used, not discarded.
    let ir = Docx::new(
        r#"<w:p><w:r><w:rPr><w:color w:val="4472C4" w:themeColor="accent1"/></w:rPr>
             <w:t>x</w:t></w:r></w:p>"#,
    )
    .ir();
    assert_eq!(first_span(para(&ir, 0)).color, Some([0x44, 0x72, 0xC4]));
}

#[test]
fn test_plain_rgb_color_still_reaches_the_ir() {
    let ir =
        Docx::new(r#"<w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>x</w:t></w:r></w:p>"#)
            .ir();
    assert_eq!(first_span(para(&ir, 0)).color, Some([0xFF, 0x00, 0x00]));
}

#[test]
fn test_auto_color_stays_none() {
    let ir =
        Docx::new(r#"<w:p><w:r><w:rPr><w:color w:val="auto"/></w:rPr><w:t>x</w:t></w:r></w:p>"#)
            .ir();
    assert_eq!(first_span(para(&ir, 0)).color, None);
}

// ---------------------------------------------------------------------------
// Paragraph geometry
// ---------------------------------------------------------------------------

#[test]
fn test_paragraph_indent_spacing_and_keep_flags_reach_the_ir() {
    let ir = Docx::new(
        r#"<w:p><w:pPr>
             <w:ind w:left="720" w:right="360" w:firstLine="240"/>
             <w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/>
             <w:keepNext/><w:keepLines/><w:pageBreakBefore/>
           </w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#,
    )
    .ir();
    let p = para(&ir, 0);
    assert_eq!(p.indent_left_twips, Some(720));
    assert_eq!(p.indent_right_twips, Some(360));
    assert_eq!(p.first_line_indent_twips, Some(240));
    assert_eq!(p.space_before_twips, Some(120));
    assert_eq!(p.space_after_twips, Some(240));
    assert_eq!(p.line_spacing, Some(LineSpacing::Auto(360)));
    assert!(p.keep_with_next && p.keep_together && p.page_break_before);
}

#[test]
fn test_hanging_indent_is_a_negative_first_line_indent() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
           <w:r><w:t>x</w:t></w:r></w:p>"#,
    )
    .ir();
    assert_eq!(para(&ir, 0).first_line_indent_twips, Some(-360));
}

#[test]
fn test_exact_line_spacing_keeps_its_rule() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr>
           <w:r><w:t>x</w:t></w:r></w:p>"#,
    )
    .ir();
    assert_eq!(para(&ir, 0).line_spacing, Some(LineSpacing::Exact(240)));
}

#[test]
fn test_paragraph_tabs_and_shading_reach_the_ir() {
    let ir = Docx::new(
        r#"<w:p><w:pPr>
             <w:shd w:val="clear" w:fill="EEEEEE"/>
             <w:tabs><w:tab w:val="right" w:pos="9000" w:leader="dot"/></w:tabs>
           </w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#,
    )
    .ir();
    let p = para(&ir, 0);
    assert_eq!(p.background_color, Some([0xEE, 0xEE, 0xEE]));
    assert_eq!(
        p.tabs,
        vec![TabStop {
            position_twips: 9000,
            alignment: TabAlignment::Right,
            leader: TabLeader::Dot,
        }]
    );
}

// ---------------------------------------------------------------------------
// Paragraph borders keep their styling
// ---------------------------------------------------------------------------

#[test]
fn test_paragraph_borders_keep_style_size_and_colour() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pBdr>
             <w:top w:val="double" w:sz="12" w:space="4" w:color="FF0000"/>
             <w:bottom w:val="single" w:sz="6" w:space="1" w:color="0000FF"/>
           </w:pBdr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#,
    )
    .ir();
    let b = para(&ir, 0).border.as_ref().expect("paragraph border");
    let top = b.top.as_ref().expect("top edge");
    assert_eq!(top.style, BorderStyle::Double);
    assert_eq!(top.size, Some(12));
    assert_eq!(top.space, Some(4));
    assert_eq!(top.color, Some([0xFF, 0x00, 0x00]));
    assert_eq!(b.bottom.as_ref().unwrap().color, Some([0x00, 0x00, 0xFF]));
}

#[test]
fn test_empty_paragraph_with_only_a_bottom_border_is_still_a_thematic_break() {
    // The horizontal-rule encoding must keep working now that `w:pBdr` is
    // parsed in full rather than narrowed to a boolean.
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6"/></w:pBdr></w:pPr></w:p>"#,
    )
    .ir();
    assert!(matches!(first(&ir), Element::ThematicBreak));
}

// ---------------------------------------------------------------------------
// Table geometry and borders
// ---------------------------------------------------------------------------

/// A linked picture (`a:blip/@r:link`, an external image relationship)
/// has no bytes in the package, only a target. The reader looked at
/// `r:embed` only, so such a picture vanished without trace.
#[test]
fn test_a_linked_picture_reaches_the_ir_renderers_and_the_writer() {
    let bytes = Docx::new(
        r#"<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
             <wp:extent cx="100" cy="100"/><wp:docPr id="1" name="P" descr="Linked logo"/>
             <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData>
               <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                 <pic:blipFill><a:blip r:link="LINK"/></pic:blipFill></pic:pic>
             </a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#,
    )
    .external_rel("LINK", rel_types::IMAGE, "https://example.com/logo.png")
    .bytes();
    let image = |ir: &DocumentIR| -> Image {
        ir.sections[0]
            .elements
            .iter()
            .find_map(|e| match e {
                Element::Image(i) => Some(i.clone()),
                _ => None,
            })
            .expect("the linked picture reached the IR")
    };
    let ir = Document::from_reader(Cursor::new(bytes.clone()), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    let img = image(&ir);
    assert_eq!(img.source_url.as_deref(), Some("https://example.com/logo.png"));
    assert!(img.data.is_none());
    assert_eq!(img.alt_text.as_deref(), Some("Linked logo"));
    let want = "![Linked logo](https://example.com/logo.png)";
    assert!(ir.to_markdown().contains(want), "{}", ir.to_markdown());
    assert!(
        ir.to_html()
            .contains(r#"<img src="https://example.com/logo.png""#)
    );
    let direct = office_oxide::docx::DocxDocument::from_reader(Cursor::new(bytes))
        .unwrap()
        .to_markdown();
    assert!(direct.contains(want), "{direct}");

    let again = Document::from_reader(Cursor::new(docx_bytes(&ir)), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    assert_eq!(image(&again).source_url, img.source_url);
}

/// The direct markdown renderer wrote a picture as `![alt](rId7)`: the
/// relationship id is not a target any markdown reader can resolve. The
/// IR renderer, with no addressable source, writes the description as
/// italic text; both must agree.
#[test]
fn test_direct_markdown_does_not_use_a_relationship_id_as_an_image_target() {
    const PNG: &[u8] = &[0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a];
    let image = |alt: Option<&str>| {
        Element::Image(Image {
            data: Some(PNG.to_vec()),
            format: Some(ImageFormat::Png),
            display_width_emu: Some(100),
            display_height_emu: Some(100),
            alt_text: alt.map(str::to_string),
            ..Default::default()
        })
    };
    let ir = DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![image(Some("A *cat*")), image(None)],
            ..Default::default()
        }],
        defined_names: Vec::new(),
    };
    let bytes = docx_bytes(&ir);
    let direct = office_oxide::docx::DocxDocument::from_reader(Cursor::new(bytes.clone()))
        .unwrap()
        .to_markdown();
    let via_ir = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .unwrap()
        .to_ir()
        .to_markdown();
    assert!(!direct.contains("rId") && !direct.contains("]("), "{direct:?}");
    assert_eq!(direct.trim(), via_ir.trim());
    assert!(direct.contains(r"*A \*cat\**"), "{direct:?}");
}

/// `w:shd` (ECMA-376 §17.3.5) paints its pattern (`w:val`) in `w:color`
/// over `w:fill`. Only `w:fill` was read, so `solid` shading — the whole
/// area in `w:color` — and percentage patterns came out as the fill alone
/// (often `auto`, i.e. nothing).
#[test]
fn test_cell_and_paragraph_shading_honour_the_pattern_and_its_colour() {
    let cell = |shd: &str| {
        format!("<w:tc><w:tcPr>{shd}</w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>")
    };
    let body = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr>{}{}{}{}{}</w:tr></w:tbl>
           <w:p><w:pPr><w:shd w:val="solid" w:color="0000FF" w:fill="auto"/></w:pPr><w:r><w:t>p</w:t></w:r></w:p>"#,
        cell(r#"<w:shd w:val="solid" w:color="FF0000" w:fill="auto"/>"#),
        cell(r#"<w:shd w:val="pct50" w:color="000000" w:fill="FFFFFF"/>"#),
        cell(r#"<w:shd w:val="clear" w:color="auto" w:fill="00FF00"/>"#),
        cell(r#"<w:shd w:val="nil" w:fill="00FF00"/>"#),
        cell(r#"<w:shd w:val="pct25" w:color="auto" w:fill="auto"/>"#),
    );
    let ir = Docx::new(&body).ir();
    let fills: Vec<_> = table(&ir).rows[0]
        .cells
        .iter()
        .map(|c| c.background_color)
        .collect();
    assert_eq!(
        fills,
        [
            Some([0xFF, 0, 0]),
            Some([0x80, 0x80, 0x80]),
            Some([0, 0xFF, 0]),
            None,
            // Automatic pattern colour is black, automatic fill white.
            Some([0xBF, 0xBF, 0xBF]),
        ]
    );
    assert_eq!(para(&ir, 1).background_color, Some([0, 0, 0xFF]));
}

/// A table whose look lives in its table style: `w:style/w:tblPr` (borders)
/// and `w:tblStylePr` conditional formatting (header-row and banded-row
/// shading), switched on per table by `w:tblLook`. `Style.table_properties`
/// was hardcoded `None` and `w:tblStylePr`/`w:tblLook` were never read, so
/// such tables converted borderless and unshaded.
#[test]
fn test_table_style_borders_and_conditional_shading_reach_the_ir() {
    let row = |t: &str| format!("<w:tr><w:tc><w:p><w:r><w:t>{t}</w:t></w:r></w:p></w:tc></w:tr>");
    let body = format!(
        r#"<w:tbl><w:tblPr><w:tblStyle w:val="Banded"/>
             <w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>
           <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>{}{}{}{}
           <w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:fill="FF0000"/></w:tcPr><w:p><w:r><w:t>own</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
           <w:p/>
           <w:tbl><w:tblPr><w:tblStyle w:val="Banded"/><w:tblLook w:val="0600"/></w:tblPr>
           <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>{}{}</w:tbl>"#,
        row("head"),
        row("one"),
        row("two"),
        row("three"),
        row("plain head"),
        row("plain one"),
    );
    let ir = Docx::new(&body)
        .styles(
            r#"<w:style w:type="table" w:styleId="Base"><w:name w:val="Base"/>
                 <w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="000000"/>
                   <w:insideH w:val="single" w:sz="4" w:color="000000"/></w:tblBorders></w:tblPr></w:style>
               <w:style w:type="table" w:styleId="Banded"><w:name w:val="Banded"/><w:basedOn w:val="Base"/>
                 <w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr>
                   <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr></w:tblStylePr>
                 <w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr>
               </w:style>"#,
        )
        .ir();
    let tables: Vec<&Table> = ir.sections[0]
        .elements
        .iter()
        .filter_map(|e| match e {
            Element::Table(t) => Some(t),
            _ => None,
        })
        .collect();
    let fills = |t: &Table| -> Vec<Option<[u8; 3]>> {
        t.rows.iter().map(|r| r.cells[0].background_color).collect()
    };
    assert!(
        tables[0]
            .border
            .as_ref()
            .is_some_and(|b| b.top.is_some() && b.inside_h.is_some()),
        "the style's borders (through basedOn) apply: {:?}",
        tables[0].border
    );
    assert_eq!(
        fills(tables[0]),
        [
            Some([0x44, 0x72, 0xC4]),
            Some([0xD9, 0xE2, 0xF3]),
            None,
            Some([0xD9, 0xE2, 0xF3]),
            Some([0xFF, 0x00, 0x00]),
        ],
        "header row, banded rows, and a cell's own shading winning"
    );
    // The older bitmask form: 0x0600 is noHBand + noVBand with no
    // header/total/column flags, so every conditional is off.
    assert_eq!(fills(tables[1]), [None, None]);
    assert!(tables[1].border.is_some());
}

/// `w:trHeight/@w:hRule` (ECMA-376 §17.18) was parsed and read by
/// nothing, and the writer omitted it — the default is `atLeast`, so an
/// exact-height row (forms, labels) came back as a minimum height.
#[test]
fn test_row_height_rule_survives_a_round_trip() {
    let ir = Docx::new(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
             <w:tr><w:trPr><w:trHeight w:val="400" w:hRule="exact"/></w:trPr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>
             <w:tr><w:trPr><w:trHeight w:val="500" w:hRule="auto"/></w:trPr><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
             <w:tr><w:trPr><w:trHeight w:val="600"/></w:trPr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc></w:tr>
           </w:tbl>"#,
    )
    .ir();
    let rules = |ir: &DocumentIR| -> Vec<_> {
        table(ir)
            .rows
            .iter()
            .map(|r| (r.height_twips, r.height_rule))
            .collect()
    };
    let want = vec![
        (Some(400), Some(RowHeightRule::Exact)),
        (Some(500), Some(RowHeightRule::Auto)),
        (Some(600), None),
    ];
    assert_eq!(rules(&ir), want);
    let again = Document::from_reader(Cursor::new(docx_bytes(&ir)), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    assert_eq!(rules(&again), want);
}

#[test]
fn test_table_geometry_reaches_the_ir() {
    let ir = Docx::new(
        r#"<w:tbl>
             <w:tblPr>
               <w:tblW w:w="9000" w:type="dxa"/>
               <w:tblInd w:w="360" w:type="dxa"/>
               <w:jc w:val="center"/>
               <w:tblCaption w:val="Quarterly figures"/>
               <w:tblBorders>
                 <w:top w:val="single" w:sz="4" w:color="000000"/>
                 <w:insideV w:val="dotted" w:sz="2" w:color="808080"/>
               </w:tblBorders>
               <w:tblCellMar>
                 <w:top w:w="60" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>
                 <w:bottom w:w="60" w:type="dxa"/><w:right w:w="120" w:type="dxa"/>
               </w:tblCellMar>
             </w:tblPr>
             <w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="5000"/></w:tblGrid>
             <w:tr>
               <w:trPr><w:trHeight w:val="480"/><w:cantSplit/><w:tblHeader/></w:trPr>
               <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>
                       <w:shd w:val="clear" w:fill="DDEEFF"/>
                       <w:vAlign w:val="center"/>
                       <w:textDirection w:val="tbRl"/>
                       <w:tcBorders><w:left w:val="thick" w:sz="24" w:color="00FF00"/></w:tcBorders>
                       <w:tcMar><w:left w:w="200" w:type="dxa"/></w:tcMar>
                     </w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
               <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
             </w:tr>
           </w:tbl>"#,
    )
    .ir();
    let t = table(&ir);
    assert_eq!(t.width_twips, Some(9000));
    assert_eq!(t.indent_left_twips, Some(360));
    assert_eq!(t.alignment, Some(TableAlignment::Center));
    assert_eq!(t.caption.as_deref(), Some("Quarterly figures"));
    assert_eq!(t.column_widths_twips, vec![4000, 5000]);
    assert_eq!(t.cell_padding_twips, Some(120));
    let tb = t.border.as_ref().expect("table border");
    assert_eq!(tb.top.as_ref().unwrap().style, BorderStyle::Single);
    assert_eq!(tb.inside_v.as_ref().unwrap().style, BorderStyle::Dotted);

    let row = &t.rows[0];
    assert_eq!(row.height_twips, Some(480));
    assert!(!row.allow_break, "w:cantSplit");
    assert!(row.is_header && row.repeat_as_header);

    let cell = &row.cells[0];
    assert_eq!(cell.width_twips, Some(4000));
    assert_eq!(cell.background_color, Some([0xDD, 0xEE, 0xFF]));
    assert_eq!(cell.vertical_align, Some(CellVerticalAlign::Center));
    assert_eq!(cell.text_direction, Some(TextDirection::TbRl));
    assert_eq!(cell.border.as_ref().unwrap().left.as_ref().unwrap().style, BorderStyle::Thick);
    assert_eq!(cell.padding.as_ref().unwrap().left_twips, Some(200));
}

#[test]
fn test_auto_and_percentage_table_widths_report_no_twips() {
    // `w:type="auto"`/`"pct"` carry no absolute measure; reporting one would
    // be a confidently wrong number.
    let ir = Docx::new(
        r#"<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
             <w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    )
    .ir();
    assert_eq!(table(&ir).width_twips, None);
}

// ---------------------------------------------------------------------------
// Page and column breaks
// ---------------------------------------------------------------------------

#[test]
fn test_page_break_is_a_page_break_not_a_thematic_break() {
    let ir = Docx::new(
        r#"<w:p><w:r><w:t>before</w:t></w:r></w:p>
           <w:p><w:r><w:br w:type="page"/></w:r></w:p>
           <w:p><w:r><w:t>after</w:t></w:r></w:p>"#,
    )
    .ir();
    let kinds: Vec<&str> = ir.sections[0]
        .elements
        .iter()
        .map(|e| match e {
            Element::Paragraph(_) => "P",
            Element::PageBreak => "PB",
            Element::ColumnBreak => "CB",
            Element::ThematicBreak => "TB",
            _ => "?",
        })
        .collect();
    assert_eq!(kinds, vec!["P", "PB", "P"]);
}

#[test]
fn test_column_break_reaches_the_ir() {
    let ir = Docx::new(r#"<w:p><w:r><w:br w:type="column"/></w:r></w:p>"#).ir();
    assert!(
        ir.sections[0]
            .elements
            .iter()
            .any(|e| matches!(e, Element::ColumnBreak)),
        "expected a ColumnBreak, got {:?}",
        ir.sections[0].elements
    );
}

// ---------------------------------------------------------------------------
// Section break type and column layout
// ---------------------------------------------------------------------------

#[test]
fn test_section_break_type_comes_from_w_type_not_the_section_index() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/>
                <w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>
           <w:r><w:t>one</w:t></w:r></w:p>
           <w:p><w:r><w:t>two</w:t></w:r></w:p>
           <w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>"#,
    )
    .ir();
    assert_eq!(ir.sections.len(), 2);
    assert_eq!(ir.sections[0].break_type, SectionBreakType::NextPage);
    assert_eq!(ir.sections[1].break_type, SectionBreakType::Continuous);
}

#[test]
fn test_column_layout_keeps_space_separator_and_widths() {
    let ir = Docx::new(
        r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>
           <w:sectPr>
             <w:cols w:num="2" w:space="480" w:sep="1">
               <w:col w:w="4320" w:space="480"/><w:col w:w="4320"/>
             </w:cols>
           </w:sectPr>"#,
    )
    .ir();
    let c = ir.sections[0].columns.as_ref().expect("columns");
    assert_eq!(c.count, 2);
    assert_eq!(c.space_twips, Some(480));
    assert!(c.separator);
    assert_eq!(c.column_widths_twips, vec![4320, 4320]);
}

/// A section's `w:pgNumType`, `w:footnotePr` and `w:endnotePr` were never
/// parsed. Page numbering now reaches the IR and the writer; the note
/// settings are read onto the section's properties.
#[test]
fn test_section_page_and_note_numbering_are_read() {
    let body = r#"<w:p><w:r><w:t>body</w:t></w:r></w:p>
        <w:sectPr>
          <w:footnotePr><w:pos w:val="beneathText"/><w:numFmt w:val="lowerRoman"/>
            <w:numStart w:val="3"/><w:numRestart w:val="eachPage"/></w:footnotePr>
          <w:endnotePr><w:numFmt w:val="upperLetter"/></w:endnotePr>
          <w:pgSz w:w="12240" w:h="15840"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
          <w:pgNumType w:fmt="lowerRoman" w:start="5"/>
        </w:sectPr>"#;
    let bytes = Docx::new(body).bytes();
    let docx = office_oxide::docx::DocxDocument::from_reader(Cursor::new(bytes.clone())).unwrap();
    let sp = &docx.sections[0];
    let foot = sp.footnote_properties.as_ref().expect("footnotePr");
    assert_eq!(foot.position.as_deref(), Some("beneathText"));
    assert_eq!(foot.number_format.as_deref(), Some("lowerRoman"));
    assert_eq!(foot.start, Some(3));
    assert_eq!(foot.restart.as_deref(), Some("eachPage"));
    let end = sp.endnote_properties.as_ref().expect("endnotePr");
    assert_eq!(end.number_format.as_deref(), Some("upperLetter"));
    assert_eq!(end.position, None);

    let ir = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    let numbering = |ir: &DocumentIR| {
        let ps = ir.sections[0].page_setup.as_ref().expect("page setup");
        (ps.page_number_start, ps.page_number_format.clone())
    };
    assert_eq!(numbering(&ir), (Some(5), Some("lowerRoman".to_string())));
    let again = Document::from_reader(Cursor::new(docx_bytes(&ir)), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    assert_eq!(numbering(&again), numbering(&ir));
}

/// The section's `w:footnotePr` / `w:endnotePr` were read onto the
/// parser's section properties and then dropped: the IR had nowhere to
/// hold them and the writer emitted an empty `<w:footnotePr/>`, so a
/// round trip reset roman, restarting footnotes to plain decimal.
#[test]
fn test_section_note_settings_reach_the_ir_and_survive_a_round_trip() {
    let body = r#"<w:p><w:r><w:t>body</w:t></w:r></w:p>
        <w:sectPr>
          <w:footnotePr><w:pos w:val="beneathText"/><w:numFmt w:val="lowerRoman"/>
            <w:numStart w:val="3"/><w:numRestart w:val="eachPage"/></w:footnotePr>
          <w:endnotePr><w:numFmt w:val="upperLetter"/></w:endnotePr>
          <w:pgSz w:w="12240" w:h="15840"/>
        </w:sectPr>"#;
    let ir = Docx::new(body).ir();
    let foot = NoteSettings {
        position: Some("beneathText".into()),
        number_format: Some("lowerRoman".into()),
        start: Some(3),
        restart: Some("eachPage".into()),
    };
    let end = NoteSettings {
        number_format: Some("upperLetter".into()),
        ..Default::default()
    };
    let notes = |ir: &DocumentIR| {
        let s = &ir.sections[0];
        (s.footnote_settings.clone(), s.endnote_settings.clone())
    };
    assert_eq!(notes(&ir), (Some(foot.clone()), Some(end.clone())));

    let bytes = docx_bytes(&ir);
    let again = Document::from_reader(Cursor::new(bytes.clone()), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    assert_eq!(notes(&again), notes(&ir));

    // Also through an inline (non-final) section break.
    let mut two = ir.clone();
    two.sections.push(Section {
        elements: vec![Element::Paragraph(Paragraph {
            content: vec![InlineContent::Text(TextSpan::plain("second"))],
            ..Default::default()
        })],
        ..Default::default()
    });
    let again = Document::from_reader(Cursor::new(docx_bytes(&two)), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    assert_eq!(notes(&again), (Some(foot), Some(end)));
}

/// `w:pgMar/@w:gutter` was parsed and never read, and the writer
/// hardcoded `w:gutter="0"`, so a bound document's binding margin was
/// zeroed on every round-trip.
#[test]
fn test_gutter_margin_survives_a_round_trip() {
    let ir = Docx::new(
        r#"<w:p><w:r><w:t>body</w:t></w:r></w:p>
           <w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
             <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
                      w:header="720" w:footer="720" w:gutter="567"/></w:sectPr>"#,
    )
    .ir();
    let gutter = |ir: &DocumentIR| ir.sections[0].page_setup.as_ref().map(|p| p.gutter_twips);
    assert_eq!(gutter(&ir), Some(567));
    let again = Document::from_reader(Cursor::new(docx_bytes(&ir)), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    assert_eq!(gutter(&again), Some(567));
}

#[test]
fn test_a_sect_pr_with_no_page_size_reports_no_page_setup() {
    // Synthesising a Letter page for a section that states none would hand
    // the consumer a measurement the document never made.
    let ir = Docx::new(
        r#"<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:type w:val="continuous"/></w:sectPr>"#,
    )
    .ir();
    assert!(ir.sections[0].page_setup.is_none());
}

// ---------------------------------------------------------------------------
// Header / footer routing
// ---------------------------------------------------------------------------

#[test]
fn test_first_default_and_even_headers_land_in_distinct_slots() {
    let ir = Docx::new(
        r#"<w:p><w:r><w:t>body</w:t></w:r></w:p>
           <w:sectPr>
             <w:headerReference w:type="default" r:id="rIdH1"/>
             <w:headerReference w:type="first" r:id="rIdH2"/>
             <w:footerReference w:type="default" r:id="rIdF1"/>
             <w:titlePg/>
           </w:sectPr>"#,
    )
    .hf("header1.xml", "rIdH1", "DEFAULT HEADER")
    .hf("header2.xml", "rIdH2", "FIRST HEADER")
    .hf("footer1.xml", "rIdF1", "DEFAULT FOOTER")
    .ir();

    let s = &ir.sections[0];
    assert_eq!(hf_text(&s.header), "DEFAULT HEADER");
    assert_eq!(hf_text(&s.first_page_header), "FIRST HEADER");
    assert_eq!(hf_text(&s.footer), "DEFAULT FOOTER");
    // The footer must not have leaked into the header slot, which is what
    // the old cumulative-index split did as soon as the counts were unequal.
    assert!(!hf_text(&s.header).contains("FOOTER"));
}

/// A `w:type="first"` header or footer is shown only in a section with
/// `w:titlePg` (ECMA-376 §17.10). `title_page` was parsed and never
/// read: the inactive part reached every surface, and the writer, which
/// emits `w:titlePg` whenever a first-page part exists, made it visible
/// in Word after a round-trip.
#[test]
fn test_an_inactive_first_page_header_is_not_content() {
    let bytes = Docx::new(
        r#"<w:p><w:r><w:t>body</w:t></w:r></w:p>
           <w:sectPr>
             <w:headerReference w:type="default" r:id="rIdH1"/>
             <w:headerReference w:type="first" r:id="rIdH2"/>
             <w:footerReference w:type="first" r:id="rIdF2"/>
           </w:sectPr>"#,
    )
    .hf("header1.xml", "rIdH1", "DEFAULT HEADER")
    .hf("header2.xml", "rIdH2", "FIRST HEADER")
    .hf("footer2.xml", "rIdF2", "FIRST FOOTER")
    .bytes();
    let doc = Document::from_reader(Cursor::new(bytes.clone()), DocumentFormat::Docx).unwrap();
    let ir = doc.to_ir();
    assert_eq!(hf_text(&ir.sections[0].header), "DEFAULT HEADER");
    assert!(ir.sections[0].first_page_header.is_none());
    assert!(ir.sections[0].first_page_footer.is_none());
    let docx = office_oxide::docx::DocxDocument::from_reader(Cursor::new(bytes)).unwrap();
    for (surface, out) in [
        ("ir plain_text", ir.plain_text()),
        ("ir markdown", ir.to_markdown()),
        ("plain_text", docx.plain_text()),
        ("to_markdown", docx.to_markdown()),
    ] {
        assert!(out.contains("DEFAULT HEADER"), "{surface}: {out:?}");
        assert!(!out.contains("FIRST"), "{surface} shows an inactive part: {out:?}");
    }
    // Written back, the section stays without a distinct first page.
    let written = docx_bytes(&ir);
    let again = office_oxide::docx::DocxDocument::from_reader(Cursor::new(written)).unwrap();
    assert!(!again.sections.iter().any(|s| s.title_page));
}

fn hf_text(hf: &Option<HeaderFooter>) -> String {
    hf.as_ref()
        .map(|h| {
            h.content
                .iter()
                .map(|e| match e {
                    Element::Paragraph(p) => inline_to_text(&p.content),
                    _ => String::new(),
                })
                .collect()
        })
        .unwrap_or_default()
}

// ---------------------------------------------------------------------------
// Document metadata
// ---------------------------------------------------------------------------

#[test]
fn test_core_properties_populate_the_ir_metadata() {
    let ir = Docx::new(r#"<w:p><w:r><w:t>body</w:t></w:r></w:p>"#)
        .core_props(
            r#"<dc:title>Quarterly Report</dc:title>
               <dc:creator>A. Author</dc:creator>
               <dc:subject>Finance</dc:subject>
               <dc:description>Q3 numbers</dc:description>
               <cp:keywords>finance, q3; revenue</cp:keywords>
               <dcterms:created>2026-01-02T03:04:05Z</dcterms:created>
               <dcterms:modified>2026-02-03T04:05:06Z</dcterms:modified>"#,
        )
        .ir();
    let m = &ir.metadata;
    assert_eq!(m.title.as_deref(), Some("Quarterly Report"));
    assert_eq!(m.author.as_deref(), Some("A. Author"));
    assert_eq!(m.subject.as_deref(), Some("Finance"));
    assert_eq!(m.description.as_deref(), Some("Q3 numbers"));
    assert_eq!(m.keywords, vec!["finance", "q3", "revenue"]);
    assert_eq!(m.created.as_deref(), Some("2026-01-02T03:04:05Z"));
    assert_eq!(m.modified.as_deref(), Some("2026-02-03T04:05:06Z"));
}

#[test]
fn test_title_falls_back_to_the_first_heading_without_core_properties() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>Fallback</w:t></w:r></w:p>"#,
    )
    .ir();
    assert_eq!(ir.metadata.title.as_deref(), Some("Fallback"));
}

// ---------------------------------------------------------------------------
// Numbering
// ---------------------------------------------------------------------------

const NUMBERING: &str = r#"
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="5"/><w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%1)"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#;

#[test]
fn test_list_start_number_and_style_reach_the_ir() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
             <w:r><w:t>one</w:t></w:r></w:p>
           <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
             <w:r><w:t>two</w:t></w:r></w:p>"#,
    )
    .numbering(NUMBERING)
    .ir();
    let list = match first(&ir) {
        Element::List(l) => l,
        other => panic!("not a list: {other:?}"),
    };
    assert!(list.ordered);
    assert_eq!(list.start_number, Some(5));
    assert_eq!(list.style, Some(ListStyle::LowerAlpha));
}

/// A bullet level's `w:lvlText` is the glyph Word draws; the writer maps
/// `ListStyle::Square`/`Circle`/`Dash` to ▪/○/– and the reader turned every
/// glyph back into a plain bullet, so the marker changed on a round-trip.
#[test]
fn test_bullet_glyph_selects_the_list_style() {
    let level = |id: u32, glyph: &str| {
        format!(
            r#"<w:abstractNum w:abstractNumId="{id}"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/>
                 <w:lvlText w:val="{glyph}"/></w:lvl></w:abstractNum>
               <w:num w:numId="{id}"><w:abstractNumId w:val="{id}"/></w:num>"#
        )
    };
    let glyphs = [
        ("\u{25AA}", ListStyle::Square),
        ("\u{25CB}", ListStyle::Circle),
        ("o", ListStyle::Circle),
        ("\u{2013}", ListStyle::Dash),
        ("\u{2022}", ListStyle::Bullet),
        ("\u{F0A7}", ListStyle::Square),
    ];
    let mut numbering = String::new();
    let mut body = String::new();
    for (i, (glyph, _)) in glyphs.iter().enumerate() {
        let id = i as u32 + 1;
        numbering.push_str(&level(id, glyph));
        body.push_str(&format!(
            r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{id}"/></w:numPr></w:pPr>
                 <w:r><w:t>item {id}</w:t></w:r></w:p><w:p><w:r><w:t>gap</w:t></w:r></w:p>"#
        ));
    }
    let ir = Docx::new(&body).numbering(&numbering).ir();
    let styles: Vec<_> = ir.sections[0]
        .elements
        .iter()
        .filter_map(|e| match e {
            Element::List(l) => Some(l.style.clone()),
            _ => None,
        })
        .collect();
    let want: Vec<_> = glyphs.iter().map(|(_, s)| Some(s.clone())).collect();
    assert_eq!(styles, want);
}

/// A numbered heading is not turned into a list item, so its level's
/// `w:pPr/w:ind` is the only place its indentation comes from; direct
/// `w:ind` still wins.
#[test]
fn test_numbering_level_indent_applies_to_a_numbered_heading() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:outlineLvl w:val="0"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>One</w:t></w:r></w:p>
           <w:p><w:pPr><w:outlineLvl w:val="0"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="100"/></w:pPr><w:r><w:t>Two</w:t></w:r></w:p>"#,
    )
    .numbering(
        r#"<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
             <w:pPr><w:ind w:left="432" w:hanging="432"/></w:pPr></w:lvl></w:abstractNum>
           <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#,
    )
    .ir();
    let heading = |i: usize| match &ir.sections[0].elements[i] {
        Element::Heading(h) => h,
        other => panic!("not a heading: {other:?}"),
    };
    assert_eq!(heading(0).indent_left_twips, Some(432));
    assert_eq!(heading(0).first_line_indent_twips, Some(-432));
    assert_eq!(heading(1).indent_left_twips, Some(100));
}

/// `w:lvlText` (the marker pattern around the counter) and `w:lvlJc` were
/// parsed and never read: `a)`, `(1)` and `1.1.` all became `1.`, a
/// sub-list took the ordered flag of whichever item came last, and the
/// writer always wrote `%N.` with no `w:lvlJc`.
#[test]
fn test_numbering_marker_pattern_and_alignment_reach_the_ir_renderers_and_writer() {
    let numbering = r#"
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>
          <w:lvlText w:val="%1)"/><w:lvlJc w:val="right"/></w:lvl>
        <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/>
          <w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/></w:lvl>
        <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="bullet"/>
          <w:lvlText w:val="&#9642;"/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#;
    let para = |ilvl: u8, text: &str| {
        format!(
            r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
        )
    };
    let body = [
        para(0, "one"),
        para(1, "one-a"),
        para(2, "dot"),
        para(0, "two"),
    ]
    .concat();
    let bytes = Docx::new(&body).numbering(numbering).bytes();
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx).unwrap();

    let check = |ir: &DocumentIR| {
        let top = match first(ir) {
            Element::List(l) => l,
            other => panic!("not a list: {other:?}"),
        };
        assert!(top.ordered);
        assert_eq!(top.marker_pattern.as_deref(), Some("%1)"));
        assert_eq!(top.marker_alignment, Some(ParagraphAlignment::Right));
        let sub = top.items[0].nested.as_ref().expect("level 1");
        assert!(sub.ordered);
        assert_eq!(sub.marker_pattern.as_deref(), Some("%1.%2."));
        assert_eq!(sub.marker_alignment, None);
        let dots = sub.items[0].nested.as_ref().expect("level 2");
        assert!(!dots.ordered, "a bullet level under a numbered one is a bullet list");
        assert_eq!(dots.style, Some(ListStyle::Square));
        assert_eq!(dots.marker_pattern, None);
    };
    let ir = doc.to_ir();
    check(&ir);

    // CommonMark has `1)`; both markdown pipelines use it, and agree.
    let direct = doc.to_markdown();
    assert!(direct.contains("1) one") && direct.contains("2) two"), "{direct}");
    assert_eq!(ir.to_markdown().trim_end(), direct.trim_end());

    // The writer keeps the pattern and the alignment.
    let again = Document::from_reader(Cursor::new(docx_bytes(&ir)), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    check(&again);
}

#[test]
fn test_num_id_zero_is_not_a_list() {
    // `<w:numId w:val="0"/>` explicitly removes numbering. Treating it as a
    // list put a bullet in front of an ordinary paragraph.
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr>
             <w:r><w:t>not a bullet</w:t></w:r></w:p>"#,
    )
    .numbering(NUMBERING)
    .ir();
    assert!(matches!(first(&ir), Element::Paragraph(_)), "got {:?}", first(&ir));
}

// ---------------------------------------------------------------------------
// style-based formatting
// ---------------------------------------------------------------------------

#[test]
fn test_style_based_run_formatting_reaches_the_ir() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pStyle w:val="Emphatic"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p>"#,
    )
    .styles(
        r#"<w:style w:type="paragraph" w:styleId="Emphatic">
             <w:name w:val="Emphatic"/>
             <w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="112233"/></w:rPr>
           </w:style>"#,
    )
    .ir();
    let s = first_span(para(&ir, 0));
    assert!(s.bold, "bold from the paragraph style");
    assert_eq!(s.font_size_half_pt, Some(28));
    assert_eq!(s.color, Some([0x11, 0x22, 0x33]));
}

#[test]
fn test_style_inheritance_walks_based_on_and_direct_formatting_wins() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr>
             <w:r><w:rPr><w:i w:val="0"/></w:rPr><w:t>x</w:t></w:r></w:p>"#,
    )
    .styles(
        r#"<w:style w:type="paragraph" w:styleId="Parent">
             <w:name w:val="Parent"/><w:rPr><w:b/><w:i/></w:rPr>
           </w:style>
           <w:style w:type="paragraph" w:styleId="Child">
             <w:name w:val="Child"/><w:basedOn w:val="Parent"/>
           </w:style>"#,
    )
    .ir();
    let s = first_span(para(&ir, 0));
    assert!(s.bold, "inherited from Parent through basedOn");
    assert!(!s.italic, "direct <w:i w:val=\"0\"/> overrides the style");
}

#[test]
fn test_document_defaults_apply_when_no_style_is_named() {
    let ir = Docx::new(r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>"#)
        .styles(
            r#"<w:docDefaults><w:rPrDefault><w:rPr>
                 <w:rFonts w:ascii="Garamond"/><w:sz w:val="24"/>
               </w:rPr></w:rPrDefault></w:docDefaults>"#,
        )
        .ir();
    let s = first_span(para(&ir, 0));
    assert_eq!(s.font_name.as_deref(), Some("Garamond"));
    assert_eq!(s.font_size_half_pt, Some(24));
}

#[test]
fn test_character_style_sits_between_paragraph_style_and_direct_formatting() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr>
             <w:r><w:rPr><w:rStyle w:val="Code"/></w:rPr><w:t>x</w:t></w:r></w:p>"#,
    )
    .styles(
        r#"<w:style w:type="paragraph" w:styleId="Body">
             <w:name w:val="Body"/><w:rPr><w:rFonts w:ascii="Georgia"/><w:b/></w:rPr>
           </w:style>
           <w:style w:type="character" w:styleId="Code">
             <w:name w:val="Code"/><w:rPr><w:rFonts w:ascii="Consolas"/></w:rPr>
           </w:style>"#,
    )
    .ir();
    let s = first_span(para(&ir, 0));
    assert_eq!(s.font_name.as_deref(), Some("Consolas"), "character style wins");
    assert!(s.bold, "paragraph style's bold survives");
}

#[test]
fn test_style_based_paragraph_geometry_reaches_the_ir() {
    let ir =
        Docx::new(r#"<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#)
            .styles(
                r#"<w:style w:type="paragraph" w:styleId="Quote">
                 <w:name w:val="Quote"/>
                 <w:pPr><w:ind w:left="1440"/><w:jc w:val="center"/></w:pPr>
               </w:style>"#,
            )
            .ir();
    let p = para(&ir, 0);
    assert_eq!(p.indent_left_twips, Some(1440));
    assert_eq!(p.alignment, Some(ParagraphAlignment::Center));
}

#[test]
fn test_a_based_on_cycle_does_not_hang() {
    let ir = Docx::new(r#"<w:p><w:pPr><w:pStyle w:val="A"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#)
        .styles(
            r#"<w:style w:type="paragraph" w:styleId="A">
                 <w:name w:val="A"/><w:basedOn w:val="B"/><w:rPr><w:b/></w:rPr></w:style>
               <w:style w:type="paragraph" w:styleId="B">
                 <w:name w:val="B"/><w:basedOn w:val="A"/></w:style>"#,
        )
        .ir();
    assert!(first_span(para(&ir, 0)).bold);
}

// ---------------------------------------------------------------------------
// Heading detection via style id / name
// ---------------------------------------------------------------------------

/// Word writes a manual page break as `<w:br w:type="page"/>` at the
/// *start* of the following paragraph. The converter kept only the runs
/// before a break, so that paragraph's whole text vanished from
/// `to_ir()`/`to_html()` while `plain_text()` kept it — 75 corpus files.
/// A paragraph splits at every hard break; the text on both sides
/// survives and the breaks keep their place.
#[test]
fn test_text_after_a_hard_break_in_the_same_paragraph_is_kept() {
    let ir = Docx::new(
        r#"<w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t>After the page break</w:t></w:r></w:p>
           <w:p><w:r><w:t>before</w:t></w:r><w:r><w:br w:type="column"/></w:r><w:r><w:t>after</w:t></w:r></w:p>
           <w:p><w:r><w:br w:type="page"/></w:r></w:p>"#,
    )
    .ir();
    let kinds: Vec<String> = ir.sections[0]
        .elements
        .iter()
        .map(|e| match e {
            Element::Paragraph(p) => format!("p:{}", para_text(p)),
            Element::PageBreak => "page".into(),
            Element::ColumnBreak => "column".into(),
            other => format!("{other:?}"),
        })
        .collect();
    assert_eq!(
        kinds,
        [
            "page",
            "p:After the page break",
            "p:before",
            "column",
            "p:after",
            "page"
        ],
        "{:?}",
        ir.sections[0].elements
    );
    let html = ir.to_html();
    assert!(html.contains("After the page break") && html.contains("after"), "{html}");
}

/// Word's multilevel-list "Heading" gallery attaches `w:numPr` to the
/// heading styles, so a numbered heading (`1. Introduction`) is a list
/// member *and* a heading. List membership used to win in this converter
/// (as in the `.doc` one), turning every numbered heading into a list
/// item and leaving the IR with no headings at all. Both the direct
/// `w:numPr` and the style-inherited one must resolve to a Heading.
#[test]
fn test_numbered_heading_is_a_heading_not_a_list_item() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Introduction</w:t></w:r></w:p>
           <w:p><w:pPr><w:pStyle w:val="NumberedHeading"/></w:pPr><w:r><w:t>Scope</w:t></w:r></w:p>
           <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>a plain item</w:t></w:r></w:p>"#,
    )
    .styles(
        r#"<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
           <w:style w:type="paragraph" w:styleId="NumberedHeading"><w:name w:val="Numbered Heading"/><w:pPr><w:outlineLvl w:val="1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>"#,
    )
    .numbering(
        r#"<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>
           <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#,
    )
    .ir();
    let kinds: Vec<&str> = ir.sections[0]
        .elements
        .iter()
        .map(|e| match e {
            Element::Heading(h) => {
                if h.level == 1 {
                    "h1"
                } else {
                    "h2"
                }
            },
            Element::List(_) => "list",
            Element::Paragraph(_) => "para",
            _ => "other",
        })
        .collect();
    assert_eq!(kinds, ["h1", "h2", "list"], "{:?}", ir.sections[0].elements);
}

#[test]
fn test_heading_style_id_without_outline_level_is_still_a_heading() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Sub</w:t></w:r></w:p>"#,
    )
    .styles(
        r#"<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style>"#,
    )
    .ir();
    match first(&ir) {
        Element::Heading(h) => assert_eq!(h.level, 2),
        other => panic!("not a heading: {other:?}"),
    }
}

#[test]
fn test_heading_style_name_without_a_conventional_id_is_still_a_heading() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pStyle w:val="berschrift1"/></w:pPr><w:r><w:t>Top</w:t></w:r></w:p>"#,
    )
    .styles(
        r#"<w:style w:type="paragraph" w:styleId="berschrift1"><w:name w:val="heading 1"/></w:style>"#,
    )
    .ir();
    match first(&ir) {
        Element::Heading(h) => assert_eq!(h.level, 1),
        other => panic!("not a heading: {other:?}"),
    }
}

#[test]
fn test_a_non_heading_style_stays_a_paragraph() {
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#,
    )
    .styles(
        r#"<w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/></w:style>"#,
    )
    .ir();
    assert!(matches!(first(&ir), Element::Paragraph(_)));
}

// ---------------------------------------------------------------------------
// Internal hyperlinks
// ---------------------------------------------------------------------------

#[test]
fn test_internal_anchor_hyperlinks_become_fragment_links() {
    let ir = Docx::new(
        r#"<w:p><w:hyperlink w:anchor="section2"><w:r><w:t>Jump</w:t></w:r></w:hyperlink></w:p>"#,
    )
    .ir();
    assert_eq!(first_span(para(&ir, 0)).hyperlink.as_deref(), Some("#section2"));
}

// ---------------------------------------------------------------------------
// XML entities and character references
// ---------------------------------------------------------------------------

#[test]
fn test_entity_and_character_references_survive_extraction() {
    // quick-xml reports `&amp;` and `&#8212;` as their own events, so a
    // reader that only handled Event::Text deleted them outright.
    let ir =
        Docx::new(r#"<w:p><w:r><w:t>AT&amp;T &#8212; &lt;tag&gt; &#x2764;</w:t></w:r></w:p>"#).ir();
    assert_eq!(first_span(para(&ir, 0)).text, "AT&T — <tag> ❤");
}

#[test]
fn test_an_unresolvable_entity_is_preserved_verbatim() {
    // A DTD-declared entity we cannot expand must not silently vanish.
    let ir = Docx::new(r#"<w:p><w:r><w:t>a&nbsp;b</w:t></w:r></w:p>"#).ir();
    assert_eq!(first_span(para(&ir, 0)).text, "a&nbsp;b");
}

// ---------------------------------------------------------------------------
// w:cr, w:noBreakHyphen, w:softHyphen, w:sym
// ---------------------------------------------------------------------------

/// Concatenate a paragraph's text spans, ignoring breaks.
fn para_text(p: &Paragraph) -> String {
    p.content
        .iter()
        .filter_map(|c| match c {
            InlineContent::Text(t) => Some(t.text.as_str()),
            _ => None,
        })
        .collect()
}

#[test]
fn test_a_soft_hyphen_does_not_split_a_word() {
    // `<w:softHyphen/>` is a discretionary line-break hint. Emitting U+00AD
    // for it splits the word for every word-level consumer — one real corpus
    // file carries 68, turning `Fähigkeit` into `Fähig` + `keit`.
    let ir =
        Docx::new(r#"<w:p><w:r><w:t>Fähig</w:t><w:softHyphen/><w:t>keit</w:t></w:r></w:p>"#).ir();
    assert_eq!(para_text(para(&ir, 0)), "Fähigkeit");
}

#[test]
fn test_non_breaking_hyphen_is_not_dropped() {
    // "e-mail" used to extract as "email".
    let ir =
        Docx::new(r#"<w:p><w:r><w:t>e</w:t><w:noBreakHyphen/><w:t>mail</w:t></w:r></w:p>"#).ir();
    assert_eq!(para_text(para(&ir, 0)), "e\u{2011}mail");
}

#[test]
fn test_carriage_return_is_a_line_break() {
    let ir = Docx::new(r#"<w:p><w:r><w:t>a</w:t><w:cr/><w:t>b</w:t></w:r></w:p>"#).ir();
    assert!(
        para(&ir, 0)
            .content
            .iter()
            .any(|c| matches!(c, InlineContent::LineBreak)),
        "expected a LineBreak, got {:?}",
        para(&ir, 0).content
    );
}

#[test]
fn test_symbol_runs_produce_a_character() {
    // The Wingdings bullet lives in the private-use area; map it to U+2022
    // rather than emitting an unrenderable code point.
    let ir = Docx::new(
        r#"<w:p><w:r><w:sym w:font="Wingdings" w:char="F0B7"/><w:t> item</w:t></w:r></w:p>"#,
    )
    .ir();
    assert_eq!(para_text(para(&ir, 0)), "\u{2022} item");
}

// ---------------------------------------------------------------------------
// Transparent paragraph wrappers
// ---------------------------------------------------------------------------

#[test]
fn test_tracked_insertions_and_field_results_are_extracted() {
    let ir = Docx::new(
        r#"<w:p>
             <w:ins w:id="1" w:author="A"><w:r><w:t>INSERTED </w:t></w:r></w:ins>
             <w:fldSimple w:instr="PAGE"><w:r><w:t>7</w:t></w:r></w:fldSimple>
             <w:smartTag w:element="place"><w:r><w:t> Paris</w:t></w:r></w:smartTag>
             <w:sdt><w:sdtContent><w:r><w:t> SDT</w:t></w:r></w:sdtContent></w:sdt>
           </w:p>"#,
    )
    .ir();
    assert_eq!(para_text(para(&ir, 0)), "INSERTED 7 Paris SDT");
}

#[test]
fn test_tracked_deletions_are_not_document_content() {
    let ir = Docx::new(
        r#"<w:p>
             <w:r><w:t>kept</w:t></w:r>
             <w:del w:id="2" w:author="A"><w:r><w:delText> removed</w:delText></w:r></w:del>
           </w:p>"#,
    )
    .ir();
    assert_eq!(para_text(para(&ir, 0)), "kept");
}

// ---------------------------------------------------------------------------
// Text boxes, and AlternateContent taken once
// ---------------------------------------------------------------------------

/// A shape carrying a text box, wrapped in the compatibility element Word
/// emits. Both branches describe the same shape and the same text.
const ALTERNATE_TEXTBOX: &str = r#"<w:p><w:r>
  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
    <mc:Choice Requires="wps">
      <w:drawing><wp:anchor xmlns:wp="x">
        <a:graphic xmlns:a="y"><a:graphicData>
          <wps:wsp xmlns:wps="z"><wps:txbx><w:txbxContent>
            <w:p><w:r><w:t>BOXED TEXT</w:t></w:r></w:p>
          </w:txbxContent></wps:txbx></wps:wsp>
        </a:graphicData></a:graphic>
      </wp:anchor></w:drawing>
    </mc:Choice>
    <mc:Fallback>
      <w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:textbox><w:txbxContent>
        <w:p><w:r><w:t>BOXED TEXT</w:t></w:r></w:p>
      </w:txbxContent></v:textbox></v:shape></w:pict>
    </mc:Fallback>
  </mc:AlternateContent>
</w:r></w:p>"#;

fn text_box_texts(ir: &DocumentIR) -> Vec<String> {
    ir.sections[0]
        .elements
        .iter()
        .filter_map(|e| match e {
            Element::TextBox(tb) => Some(
                tb.content
                    .iter()
                    .map(|e| match e {
                        Element::Paragraph(p) => para_text(p),
                        _ => String::new(),
                    })
                    .collect::<String>(),
            ),
            _ => None,
        })
        .collect()
}

#[test]
fn test_vml_text_box_content_is_extracted() {
    let ir = Docx::new(
        r#"<w:p><w:r><w:pict>
             <v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:textbox><w:txbxContent>
               <w:p><w:r><w:t>SIDEBAR</w:t></w:r></w:p>
             </w:txbxContent></v:textbox></v:shape>
           </w:pict></w:r></w:p>"#,
    )
    .ir();
    assert_eq!(text_box_texts(&ir), vec!["SIDEBAR"]);
}

#[test]
fn test_drawingml_text_box_content_is_extracted_exactly_once() {
    // Extracting both branches duplicated the text; extracting neither
    // dropped it. This was 90% of the text in the original reproducer.
    let ir = Docx::new(ALTERNATE_TEXTBOX).ir();
    assert_eq!(text_box_texts(&ir), vec!["BOXED TEXT"]);
}

// ---------------------------------------------------------------------------
// Footnotes, endnotes and comments
// ---------------------------------------------------------------------------

#[test]
fn test_footnote_bodies_reach_the_ir() {
    let ir = Docx::new(r#"<w:p><w:r><w:t>body</w:t></w:r></w:p>"#)
        .notes(
            "footnotes.xml",
            rel_types::FOOTNOTES,
            CT_FOOTNOTES,
            r#"<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:t>SEP</w:t></w:r></w:p></w:footnote>
               <w:footnote w:id="1"><w:p><w:r><w:t>NOTE ONE</w:t></w:r></w:p></w:footnote>"#,
        )
        .ir();
    let notes: Vec<&Note> = ir.sections[0]
        .elements
        .iter()
        .filter_map(|e| match e {
            Element::Footnote(n) => Some(n),
            _ => None,
        })
        .collect();
    assert_eq!(notes.len(), 1, "separator pseudo-notes must be filtered out");
    assert_eq!(notes[0].id, 1);
    let text = ir.plain_text();
    assert!(text.contains("NOTE ONE"));
    assert!(!text.contains("SEP"));
}

#[test]
fn test_comment_bodies_reach_the_ir_with_their_author() {
    let ir = Docx::new(r#"<w:p><w:r><w:t>body</w:t></w:r></w:p>"#)
        .notes(
            "comments.xml",
            rel_types::COMMENTS,
            CT_COMMENTS,
            r#"<w:comment w:id="3" w:author="Reviewer"><w:p><w:r><w:t>Please check</w:t></w:r></w:p></w:comment>"#,
        )
        .ir();
    let note = ir.sections[0]
        .elements
        .iter()
        .find_map(|e| match e {
            Element::Endnote(n) => Some(n),
            _ => None,
        })
        .expect("comment reached the IR");
    assert_eq!(note.author.as_deref(), Some("Reviewer"));
    // The label every surface and format uses for a comment.
    assert_eq!(note.marker.as_deref(), Some("Comment (Reviewer)"));
    assert!(ir.plain_text().contains("Comment (Reviewer): "), "{}", ir.plain_text());
}

/// Where a comment's anchored range starts, and its citation point, as a
/// sequence of markers around the paragraph's text.
fn comment_shape(p: &Paragraph) -> Vec<String> {
    p.content
        .iter()
        .map(|c| match c {
            InlineContent::Text(s) => s.text.trim().to_string(),
            InlineContent::CommentStart(a) => format!("<{}", a.comment_id),
            InlineContent::CommentRef(a) => format!("{}>", a.comment_id),
            other => format!("{other:?}"),
        })
        .filter(|s| !s.is_empty())
        .collect()
}

const COMMENTED_BODY: &str = r#"<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>
  <w:commentRangeStart w:id="4"/><w:r><w:t>commented</w:t></w:r><w:commentRangeEnd w:id="4"/>
  <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="4"/></w:r>
  <w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>"#;
const COMMENTS: &str = r#"<w:comment w:id="4" w:author="Reviewer"><w:p><w:r><w:t>Please check</w:t></w:r></w:p></w:comment>"#;

/// `w:commentRangeStart` (ECMA-376 §17.13.4) and `w:commentReference`
/// (§17.13.4) were never read into the IR, so a comment's body survived
/// with no record of what it was about.
#[test]
fn test_comment_anchor_reaches_the_ir() {
    let ir = Docx::new(COMMENTED_BODY)
        .notes("comments.xml", rel_types::COMMENTS, CT_COMMENTS, COMMENTS)
        .ir();
    assert_eq!(comment_shape(para(&ir, 0)), ["Before", "<4", "commented", "4>", "after"]);
}

/// A DOCX comment read into the IR (an `Endnote` labelled "Comment (…)")
/// was written back as an endnote: the writer had no comments part at
/// all, so read→write turned every comment into an endnote detached from
/// its anchor.
#[test]
fn test_a_comment_round_trips_as_an_anchored_comment() {
    let ir = Docx::new(COMMENTED_BODY)
        .notes("comments.xml", rel_types::COMMENTS, CT_COMMENTS, COMMENTS)
        .ir();
    let bytes = docx_bytes(&ir);
    let doc = office_oxide::docx::DocxDocument::from_reader(Cursor::new(bytes.clone())).unwrap();
    assert_eq!(doc.comments.len(), 1, "written as a comment");
    assert!(doc.endnotes.is_empty(), "not as an endnote");
    assert_eq!(doc.comments[0].author.as_deref(), Some("Reviewer"));
    let again = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    assert_eq!(comment_shape(para(&again, 0)), ["Before", "<4", "commented", "4>", "after"]);
    assert!(
        again
            .plain_text()
            .contains("Comment (Reviewer): Please check")
    );
}

/// A comment with no anchor in the IR (one from a slide or built by hand)
/// is still a comment on write, cited at the end of the body.
#[test]
fn test_an_unanchored_comment_is_written_as_a_comment() {
    let ir = DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Docx,
            ..Default::default()
        },
        sections: vec![Section {
            elements: vec![
                Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain("Body"))],
                    ..Default::default()
                }),
                Element::Endnote(Note {
                    id: 9,
                    content: vec![Element::Paragraph(Paragraph {
                        content: vec![InlineContent::Text(TextSpan::plain("Loose remark"))],
                        ..Default::default()
                    })],
                    marker: Some("Comment (Ann)".to_string()),
                    author: Some("Ann".to_string()),
                }),
            ],
            ..Default::default()
        }],
        defined_names: Vec::new(),
    };
    let bytes = docx_bytes(&ir);
    let doc = office_oxide::docx::DocxDocument::from_reader(Cursor::new(bytes.clone())).unwrap();
    assert_eq!(doc.comments.len(), 1);
    assert!(doc.endnotes.is_empty());
    let again = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    assert!(
        comment_shape(para(&again, 0))
            .iter()
            .any(|s| s.ends_with('>')),
        "the comment is cited in the body: {:?}",
        para(&again, 0)
    );
}

// ---------------------------------------------------------------------------
// Image alt text must not also become body text
// ---------------------------------------------------------------------------

#[test]
fn test_image_alt_text_is_not_duplicated_as_body_text() {
    // Emitting the alt text as a run *and* on the Image made the IR
    // round-trip unbounded: the document grew on every read/write cycle.
    let ir = Docx::new(
        r#"<w:p><w:r><w:drawing>
             <wp:inline xmlns:wp="x">
               <wp:extent cx="100" cy="100"/>
               <wp:docPr id="1" name="Pic" descr="ALT TEXT"/>
               <a:graphic xmlns:a="y"><a:graphicData>
                 <pic:pic xmlns:pic="p"><pic:blipFill>
                   <a:blip r:embed="rIdImg"/>
                 </pic:blipFill></pic:pic>
               </a:graphicData></a:graphic>
             </wp:inline>
           </w:drawing></w:r></w:p>"#,
    )
    .ir();
    let body: String = ir.sections[0]
        .elements
        .iter()
        .filter_map(|e| match e {
            Element::Paragraph(p) => Some(para_text(p)),
            _ => None,
        })
        .collect();
    assert!(
        !body.contains("ALT TEXT"),
        "alt text must not appear as body text, got {body:?}"
    );
}

/// A picture in a header resolves through the header part's own
/// relationships. Images were loaded from the main document's rels only,
/// so a header logo was dropped — and since `r:id`s are per part, the
/// header's `rId1` is unrelated to the body's `rId1` (here: the header
/// relationship itself).
#[test]
fn test_a_picture_in_a_header_is_read_through_the_header_parts_rels() {
    const PNG: &[u8] = &[0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3];
    let ir = Docx::new(
        r#"<w:p><w:r><w:t>BODY</w:t></w:r></w:p>
           <w:sectPr><w:headerReference w:type="default" r:id="RID_H"/></w:sectPr>"#,
    )
    .header_with_picture("RID_H", PNG)
    .ir();
    let header = ir.sections[0]
        .header
        .as_ref()
        .expect("section should carry its header");
    let image = header
        .content
        .iter()
        .find_map(|e| match e {
            Element::Image(img) => Some(img),
            _ => None,
        })
        .unwrap_or_else(|| panic!("header should hold the picture: {header:?}"));
    assert_eq!(image.alt_text.as_deref(), Some("HEADER LOGO"));
    assert_eq!(
        image.data.as_deref(),
        Some(PNG),
        "the bytes must come from the header's own image relationship"
    );
}

/// A header part that is not well-formed (a damaged archive with an
/// intact body) is skipped with a warning; the document is still read,
/// as Tika/POI read it. It used to fail the whole file.
#[test]
fn test_an_unreadable_header_part_does_not_fail_the_document() {
    let ir = Docx::new(
        r#"<w:p><w:r><w:t>BODY TEXT</w:t></w:r></w:p>
           <w:sectPr><w:headerReference w:type="default" r:id="RID_H"/></w:sectPr>"#,
    )
    .header_raw("RID_H", b"<w:hdr xmlns:w=\"x\"><w:p><w:r w:a=\"unterminated")
    .ir();
    assert!(ir.plain_text().contains("BODY TEXT"));
    assert!(ir.sections[0].header.is_none(), "{:?}", ir.sections[0].header);
}

// ---------------------------------------------------------------------------
// altChunk embedded content
// ---------------------------------------------------------------------------

#[test]
fn test_html_alt_chunk_content_is_extracted_at_its_position() {
    // `altChunk` is how mail merge, report generators and CMS exporters
    // inject content — often the whole body, with document.xml holding only
    // a shell. Such a document extracted as almost nothing.
    let ir = Docx::new(
        r#"<w:p><w:r><w:t>BEFORE</w:t></w:r></w:p>
           <w:altChunk r:id="RID_A"/>
           <w:p><w:r><w:t>AFTER</w:t></w:r></w:p>"#,
    )
    .alt_chunk(
        "chunk1.html",
        CT_HTML,
        "RID_A",
        "<html><body><p>CHUNK ONE</p><p>CHUNK &amp; TWO</p>\
         <script>ignored()</script></body></html>",
    )
    .ir();

    let text = ir.plain_text();
    assert!(text.contains("CHUNK ONE"), "chunk text missing from {text:?}");
    assert!(text.contains("CHUNK & TWO"), "entity not resolved: {text:?}");
    assert!(!text.contains("ignored()"), "script body leaked: {text:?}");

    // And it must land between the two paragraphs.
    let before = text.find("BEFORE").expect("BEFORE");
    let chunk = text.find("CHUNK ONE").expect("chunk");
    let after = text.find("AFTER").expect("AFTER");
    assert!(before < chunk && chunk < after, "wrong position: {text:?}");
}

#[test]
fn test_a_plain_text_alt_chunk_is_extracted() {
    let ir = Docx::new(r#"<w:altChunk r:id="RID_A"/>"#)
        .alt_chunk("chunk1.txt", "text/plain", "RID_A", "LINE ONE\nLINE TWO")
        .ir();
    let text = ir.plain_text();
    assert!(text.contains("LINE ONE") && text.contains("LINE TWO"), "{text:?}");
}

#[test]
fn test_an_unsupported_alt_chunk_type_inserts_nothing_rather_than_garbage() {
    // A nested `.docx` chunk is a whole package; reading it is a bigger job
    // and is deliberately not attempted.
    let ir = Docx::new(r#"<w:p><w:r><w:t>ONLY</w:t></w:r></w:p><w:altChunk r:id="RID_A"/>"#)
        .alt_chunk(
            "chunk1.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "RID_A",
            "PK not really a package",
        )
        .ir();
    assert_eq!(ir.plain_text().trim(), "ONLY");
}

// ---------------------------------------------------------------------------
// Regression: numbering inherited from a style
// ---------------------------------------------------------------------------

#[test]
fn test_numbering_inherited_from_a_style_forms_a_list_and_terminates() {
    // The caller decides "this is a list" from the *effective* properties,
    // which include a `w:numPr` inherited from the paragraph style. The
    // group loop used to test the *direct* `w:pPr`, so it matched nothing,
    // consumed no paragraphs, and the caller looped forever appending empty
    // lists until the process was OOM-killed. Found on 8 real corpus files
    // in the 0.1.9 -> 0.1.10 sweep.
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pStyle w:val="ListPara"/></w:pPr>
             <w:r><w:t>one</w:t></w:r></w:p>
           <w:p><w:pPr><w:pStyle w:val="ListPara"/></w:pPr>
             <w:r><w:t>two</w:t></w:r></w:p>
           <w:p><w:r><w:t>after</w:t></w:r></w:p>"#,
    )
    .styles(
        r#"<w:style w:type="paragraph" w:styleId="ListPara">
             <w:name w:val="List Paragraph"/>
             <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
           </w:style>"#,
    )
    .numbering(NUMBERING)
    .ir();

    let els = &ir.sections[0].elements;
    // Exactly one list holding both items, then the trailing paragraph.
    let lists: Vec<&List> = els
        .iter()
        .filter_map(|e| match e {
            Element::List(l) => Some(l),
            _ => None,
        })
        .collect();
    assert_eq!(lists.len(), 1, "expected one list, got {} in {els:?}", lists.len());
    assert_eq!(lists[0].items.len(), 2, "both style-numbered paragraphs join the list");
    assert!(
        els.iter().any(|e| matches!(e, Element::Paragraph(_))),
        "the trailing non-list paragraph must survive"
    );
}

#[test]
fn test_a_list_group_that_matches_nothing_still_advances() {
    // Belt-and-braces for the same defect: even if the two membership tests
    // ever disagree again, conversion must terminate. A `w:numPr` with no
    // resolvable numbering definition is the shape that gets closest.
    let ir = Docx::new(
        r#"<w:p><w:pPr><w:pStyle w:val="Ghost"/></w:pPr>
             <w:r><w:t>alpha</w:t></w:r></w:p>
           <w:p><w:r><w:t>beta</w:t></w:r></w:p>"#,
    )
    .styles(
        r#"<w:style w:type="paragraph" w:styleId="Ghost">
             <w:name w:val="Ghost"/>
             <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="77"/></w:numPr></w:pPr>
           </w:style>"#,
    )
    .ir();
    // The assertion that matters is that we got here at all.
    let text = ir.plain_text();
    assert!(text.contains("alpha") && text.contains("beta"), "got {text:?}");
}

// ---------------------------------------------------------------------------
// Regression: text-box content must not fuse with the run after it
// ---------------------------------------------------------------------------

#[test]
fn test_text_box_content_does_not_fuse_with_the_following_run() {
    // Text-box prose is *block* content. Pasting it into the inline stream
    // bare glued the last word of the box to the first word after it —
    // `Linz` + `ANTRAG` became `LinzANTRAG` on a real corpus file.
    use office_oxide::{Document, DocumentFormat};

    let bytes = {
        let mut w = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
        let part = PartName::new("/word/document.xml").unwrap();
        w.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
        let xml = r#"<?xml version="1.0"?><w:document
             xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
           <w:body><w:p>
             <w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:textbox>
               <w:txbxContent><w:p><w:r><w:t>BOXEND</w:t></w:r></w:p></w:txbxContent>
             </v:textbox></v:shape></w:pict></w:r>
             <w:r><w:t>NEXTWORD</w:t></w:r>
           </w:p></w:body></w:document>"#;
        w.add_part(&part, CT_DOC, xml.as_bytes()).unwrap();
        w.finish().unwrap().into_inner()
    };

    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx).expect("parse");
    for (name, out) in [("plain", doc.plain_text()), ("markdown", doc.to_markdown())] {
        assert!(
            out.contains("BOXEND") && out.contains("NEXTWORD"),
            "{name} lost content: {out:?}"
        );
        assert!(
            !out.contains("BOXENDNEXTWORD"),
            "{name} fused the text box to the next run: {out:?}"
        );
    }
}

// ---------------------------------------------------------------------------
// Nested lists
// ---------------------------------------------------------------------------

const NESTED_WORDS: [&str; 6] = ["one", "one-a", "one-b", "two", "two-a", "three"];

/// Positions of each of `NESTED_WORDS` as a whole line/item in `s`.
fn nested_word_positions(s: &str) -> Vec<usize> {
    let s = format!("{s}\n");
    NESTED_WORDS
        .iter()
        .map(|w| {
            [format!(" {w}\n"), format!("\n{w}\n"), format!(">{w}<")]
                .iter()
                .find_map(|pat| s.find(pat.as_str()))
                .unwrap_or_else(|| panic!("{w} missing from {s}"))
        })
        .collect()
}

fn assert_nested_order(doc: &Document) {
    let ir = doc.to_ir();
    let list = match first(&ir) {
        Element::List(l) => l,
        other => panic!("not a list: {other:?}"),
    };
    assert_eq!(list.items.len(), 3, "top level: {list:?}");
    let kids = |i: usize| list.items[i].nested.as_ref().map_or(0, |l| l.items.len());
    assert_eq!((kids(0), kids(1), kids(2)), (2, 1, 0), "children moved: {list:?}");
    for s in [
        doc.to_markdown(),
        ir.to_markdown(),
        doc.to_html(),
        doc.plain_text(),
    ] {
        let pos = nested_word_positions(&format!("\n{s}"));
        assert!(pos.windows(2).all(|w| w[0] < w[1]), "out of order: {s}");
    }
}

/// A sub-list written by the DOCX writer belongs under the item it hangs
/// off. The writer emitted every item of a level first and the sub-lists
/// after them, so on the way back in both markdown pipelines (and HTML)
/// showed "one / two / three / one-a …": each child under the wrong parent.
#[test]
fn test_written_nested_list_keeps_children_under_their_parent() {
    let md = "- one\n  - one-a\n  - one-b\n- two\n  - two-a\n- three\n";
    let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
    let doc = Document::from_reader(Cursor::new(docx_bytes(&ir)), DocumentFormat::Docx).unwrap();
    assert_nested_order(&doc);

    // The same list read from Word-shaped paragraphs.
    let para = |ilvl: u8, text: &str| {
        format!(
            r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
        )
    };
    let body = [
        (0, "one"),
        (1, "one-a"),
        (1, "one-b"),
        (0, "two"),
        (1, "two-a"),
        (0, "three"),
    ]
    .iter()
    .map(|(l, t)| para(*l, t))
    .collect::<String>();
    let bytes = Docx::new(&body)
        .numbering(
            r#"<w:abstractNum w:abstractNumId="0">
                 <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#8226;"/></w:lvl>
                 <w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#8226;"/></w:lvl>
               </w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#,
        )
        .bytes();
    assert_nested_order(&Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx).unwrap());
}
