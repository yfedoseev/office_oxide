//! Pictures held as their stored streams (`lazy_pictures`).
//!
//! Opening a document inflated every picture it holds and kept them for its
//! life, although text, Markdown and HTML without images never read them,
//! and the IR copied each picture into every image that drew it. With
//! `ParseOptions::lazy_pictures` a deflated picture is kept as its stored
//! stream and inflated only when used, and the IR shares it. Every output
//! must be the same either way, and a picture that cannot be read must be
//! reported the same way.
//!
//! The memory checks count the heap with a global allocator, so this file is
//! its own test binary.

mod common;

use std::alloc::{GlobalAlloc, Layout, System};
use std::io::{Cursor, Write};
use std::sync::atomic::{AtomicUsize, Ordering};

use office_oxide::ir::Element;
use office_oxide::ir_render::{HtmlOptions, ImageEmbed, MarkdownOptions};
use office_oxide::options::ParseOptions;
use office_oxide::{Document, DocumentFormat};

struct Counting;
static LIVE: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);

fn grow(by: usize) {
    let now = LIVE.fetch_add(by, Ordering::Relaxed) + by;
    PEAK.fetch_max(now, Ordering::Relaxed);
}

// SAFETY: forwards every call to `System` unchanged and only counts.
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        grow(layout.size());
        // SAFETY: same layout contract as the caller's.
        unsafe { System.alloc(layout) }
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        LIVE.fetch_sub(layout.size(), Ordering::Relaxed);
        // SAFETY: `ptr` came from `System.alloc` with this layout.
        unsafe { System.dealloc(ptr, layout) }
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        LIVE.fetch_sub(layout.size(), Ordering::Relaxed);
        grow(new_size);
        // SAFETY: forwarded unchanged.
        unsafe { System.realloc(ptr, layout, new_size) }
    }
}

#[global_allocator]
static GLOBAL: Counting = Counting;

/// The counter is process-wide: every test holds this for its whole body.
fn serial() -> common::LimitScope {
    common::LimitScope::new()
}

/// What `f` left live, and the most that was live at once, above what was
/// live when it began.
fn measure<T>(f: impl FnOnce() -> T) -> (T, usize, usize) {
    let start = LIVE.load(Ordering::SeqCst);
    PEAK.store(start, Ordering::SeqCst);
    let out = f();
    let live = LIVE.load(Ordering::SeqCst).saturating_sub(start);
    (out, live, PEAK.load(Ordering::SeqCst).saturating_sub(start))
}

const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const WML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/// Pictures per document, and the inflated size of each.
const PICTURES: usize = 4;
const PICTURE_LEN: usize = 2 << 20;

/// A "PNG" that deflates to a few kilobytes: a signature, then a pattern.
fn picture(n: usize) -> Vec<u8> {
    let mut data = vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A];
    data.extend((0..PICTURE_LEN - 8).map(|i| ((i / 4096 + n) % 251) as u8));
    data
}

fn zip_of(parts: Vec<(String, Vec<u8>)>) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let deflated = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    for (name, data) in parts {
        zip.start_file(name, deflated).unwrap();
        zip.write_all(&data).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn rels(entries: &[(String, &str, String)]) -> Vec<u8> {
    let mut s = format!(r#"<Relationships xmlns="{PKG_REL}">"#);
    for (id, ty, target) in entries {
        s.push_str(&format!(r#"<Relationship Id="{id}" Type="{REL}/{ty}" Target="{target}"/>"#));
    }
    s.push_str("</Relationships>");
    s.into_bytes()
}

const CT: &str = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#;

/// A `.docx` whose body draws picture `i % pictures` in each of `draws`
/// paragraphs, between lines of text.
fn docx(pictures: usize, draws: usize) -> Vec<u8> {
    let mut body = String::new();
    for i in 0..draws {
        body.push_str(&format!(
            r#"<w:p><w:r><w:t>Paragraph {i} before a picture.</w:t></w:r></w:p><w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="100" cy="100"/><wp:docPr id="{}" name="p" descr="picture {i}"/><a:graphic xmlns:a="{DML}"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="rIdP{}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#,
            i + 1,
            i % pictures
        ));
    }
    let doc_rels: Vec<_> = (0..pictures)
        .map(|n| (format!("rIdP{n}"), "image", format!("media/img{n}.png")))
        .collect();
    let mut parts = vec![
        ("[Content_Types].xml".to_string(), CT.as_bytes().to_vec()),
        (
            "_rels/.rels".to_string(),
            rels(&[("rId1".into(), "officeDocument", "word/document.xml".into())]),
        ),
        (
            "word/document.xml".to_string(),
            format!(r#"<w:document xmlns:w="{WML}" xmlns:r="{REL}"><w:body>{body}</w:body></w:document>"#)
                .into_bytes(),
        ),
        ("word/_rels/document.xml.rels".to_string(), rels(&doc_rels)),
    ];
    for n in 0..pictures {
        parts.push((format!("word/media/img{n}.png"), picture(n)));
    }
    zip_of(parts)
}

/// A one-slide `.pptx` drawing picture `i % pictures` `draws` times.
fn pptx(pictures: usize, draws: usize) -> Vec<u8> {
    let mut shapes = String::new();
    for i in 0..draws {
        shapes.push_str(&format!(
            r#"<p:pic><p:nvPicPr><p:cNvPr id="{}" name="P{i}" descr="picture {i}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdP{}"/></p:blipFill><p:spPr><a:xfrm><a:off x="{}" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:pic>"#,
            i + 2,
            i % pictures,
            i * 1000
        ));
    }
    let slide = format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="999" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>Slide text</a:t></a:r></a:p></p:txBody></p:sp>{shapes}</p:spTree></p:cSld></p:sld>"#
    );
    let slide_rels: Vec<_> = (0..pictures)
        .map(|n| (format!("rIdP{n}"), "image", format!("../media/img{n}.png")))
        .collect();
    let mut parts = vec![
        ("[Content_Types].xml".to_string(), CT.as_bytes().to_vec()),
        (
            "_rels/.rels".to_string(),
            rels(&[("rId1".into(), "officeDocument", "ppt/presentation.xml".into())]),
        ),
        (
            "ppt/presentation.xml".to_string(),
            format!(r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#)
                .into_bytes(),
        ),
        (
            "ppt/_rels/presentation.xml.rels".to_string(),
            rels(&[("rId2".into(), "slide", "slides/slide1.xml".into())]),
        ),
        ("ppt/slides/slide1.xml".to_string(), slide.into_bytes()),
        ("ppt/slides/_rels/slide1.xml.rels".to_string(), rels(&slide_rels)),
    ];
    for n in 0..pictures {
        parts.push((format!("ppt/media/img{n}.png"), picture(n)));
    }
    zip_of(parts)
}

/// A one-sheet `.xlsx` whose drawing anchors picture `i % pictures` `draws`
/// times.
fn xlsx(pictures: usize, draws: usize) -> Vec<u8> {
    let mut drawing = String::from(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
    );
    for i in 0..draws {
        drawing.push_str(&format!(
            r#"<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{i}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="100" cy="100"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{}" name="P{i}" descr="picture {i}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdP{}"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>"#,
            i + 2,
            i % pictures
        ));
    }
    drawing.push_str("</xdr:wsDr>");
    let drawing_rels: Vec<_> = (0..pictures)
        .map(|n| (format!("rIdP{n}"), "image", format!("../media/img{n}.png")))
        .collect();
    let mut parts = vec![
        ("[Content_Types].xml".to_string(), CT.as_bytes().to_vec()),
        (
            "_rels/.rels".to_string(),
            rels(&[("rId1".into(), "officeDocument", "xl/workbook.xml".into())]),
        ),
        (
            "xl/workbook.xml".to_string(),
            format!(r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#)
                .into_bytes(),
        ),
        (
            "xl/_rels/workbook.xml.rels".to_string(),
            rels(&[("rId1".into(), "worksheet", "worksheets/sheet1.xml".into())]),
        ),
        (
            "xl/worksheets/sheet1.xml".to_string(),
            format!(r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>cell text</t></is></c></row></sheetData><drawing r:id="rId1"/></worksheet>"#)
                .into_bytes(),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels".to_string(),
            rels(&[("rId1".into(), "drawing", "../drawings/drawing1.xml".into())]),
        ),
        ("xl/drawings/drawing1.xml".to_string(), drawing.into_bytes()),
        ("xl/drawings/_rels/drawing1.xml.rels".to_string(), rels(&drawing_rels)),
    ];
    for n in 0..pictures {
        parts.push((format!("xl/media/img{n}.png"), picture(n)));
    }
    zip_of(parts)
}

fn options(lazy: bool) -> ParseOptions {
    let mut o = ParseOptions::default();
    o.lazy_pictures = lazy;
    o
}

fn open(bytes: &[u8], format: DocumentFormat, lazy: bool) -> Document {
    Document::from_reader_with(Cursor::new(bytes.to_vec()), format, &options(lazy)).unwrap()
}

fn formats() -> [(DocumentFormat, Vec<u8>); 3] {
    [
        (DocumentFormat::Docx, docx(PICTURES, PICTURES)),
        (DocumentFormat::Pptx, pptx(PICTURES, PICTURES)),
        (DocumentFormat::Xlsx, xlsx(PICTURES, PICTURES)),
    ]
}

/// Opening holds the pictures' stored size, not their inflated size, and
/// inflates at most about one picture at a time.
#[test]
fn test_a_lazy_open_holds_the_stored_pictures_not_the_inflated_ones() {
    let _serial = serial();
    let inflated = PICTURES * PICTURE_LEN;
    for (format, bytes) in formats() {
        let (eager, eager_live, _) = measure(|| open(&bytes, format, false));
        assert!(eager_live >= inflated, "{format:?}: an eager open held only {eager_live} bytes");
        drop(eager);
        let (lazy, lazy_live, lazy_peak) = measure(|| open(&bytes, format, true));
        assert!(
            lazy_live < inflated / 8,
            "{format:?}: a lazy open held {lazy_live} bytes of {inflated} inflated"
        );
        assert!(
            lazy_peak < 3 * PICTURE_LEN,
            "{format:?}: a lazy open peaked at {lazy_peak} bytes (one picture is {PICTURE_LEN})"
        );
        // Text never inflates a picture.
        let (_, _, text_peak) = measure(|| lazy.plain_text());
        assert!(text_peak < PICTURE_LEN, "{format:?}: plain_text peaked at {text_peak}");
    }
}

/// Every output is byte for byte the same whether pictures are lazy or not.
#[test]
fn test_every_output_is_the_same_with_lazy_pictures() {
    let _serial = serial();
    for (format, bytes) in formats() {
        let eager = open(&bytes, format, false);
        let lazy = open(&bytes, format, true);
        let md = |d: &Document| {
            d.to_markdown_with(MarkdownOptions {
                image_embed: ImageEmbed::Base64,
            })
        };
        let html = |d: &Document| {
            d.to_html_with(HtmlOptions {
                image_embed: ImageEmbed::Base64,
            })
        };
        assert_eq!(eager.plain_text(), lazy.plain_text(), "{format:?} text");
        assert_eq!(eager.to_markdown(), lazy.to_markdown(), "{format:?} markdown");
        assert_eq!(md(&eager), md(&lazy), "{format:?} markdown with images");
        assert!(md(&lazy).contains("[image-base64:iVBORw0KGgo"), "{format:?}: images embedded");
        assert_eq!(eager.to_html(), lazy.to_html(), "{format:?} html");
        assert_eq!(html(&eager), html(&lazy), "{format:?} html with images");
        let (e, l) = (eager.to_ir(), lazy.to_ir());
        assert_eq!(e.to_json_string().unwrap(), l.to_json_string().unwrap(), "{format:?} IR JSON");
        assert_eq!(eager.damage(), lazy.damage(), "{format:?} damage");
    }
}

fn images(elements: &[Element], out: &mut Vec<office_oxide::ir::Image>) {
    for e in elements {
        match e {
            Element::Image(i) => out.push(i.clone()),
            Element::TextBox(t) => images(&t.content, out),
            _ => {},
        }
    }
}

/// One picture drawn many times is one allocation in a lazy IR, where an
/// eager IR holds a copy per drawing.
#[test]
fn test_a_lazy_ir_shares_a_picture_drawn_many_times() {
    let _serial = serial();
    for (format, bytes) in [
        (DocumentFormat::Docx, docx(1, 2)),
        (DocumentFormat::Pptx, pptx(1, 2)),
        (DocumentFormat::Xlsx, xlsx(1, 2)),
    ] {
        let lazy = open(&bytes, format, true);
        let mut found = Vec::new();
        for s in &lazy.to_ir().sections {
            images(&s.elements, &mut found);
        }
        assert_eq!(found.len(), 2, "{format:?}");
        let (a, b) = (found[0].shared_data.as_ref(), found[1].shared_data.as_ref());
        assert!(found.iter().all(|i| i.data.is_none()), "{format:?}: no copies");
        assert!(a.unwrap().ptr_eq(b.unwrap()), "{format:?}: one allocation");
        assert_eq!(found[0].byte_len(), Some(PICTURE_LEN));
        assert_eq!(&found[0].bytes().unwrap()[..], &picture(0)[..]);
        let eager = open(&bytes, format, false);
        let mut copies = Vec::new();
        for s in &eager.to_ir().sections {
            images(&s.elements, &mut copies);
        }
        assert!(
            copies
                .iter()
                .all(|i| i.shared_data.is_none() && i.data.is_some())
        );
    }
}

/// The readers' own accessors give the pictures in both modes.
#[test]
fn test_the_readers_give_their_pictures_lazy_or_not() {
    let _serial = serial();
    let want = picture(0);
    for lazy in [false, true] {
        let d = open(&docx(1, 1), DocumentFormat::Docx, lazy);
        let (pic, ext) = d.as_docx().unwrap().picture("rIdP0").unwrap();
        assert_eq!(&pic.bytes().unwrap()[..], &want[..]);
        assert_eq!(ext, Some("png"));
        assert_eq!(pic.is_deflated(), lazy);

        let p = open(&pptx(1, 1), DocumentFormat::Pptx, lazy);
        let pic = p.as_pptx().unwrap().slide_picture(0, "rIdP0").unwrap();
        assert_eq!(&pic.bytes().unwrap()[..], &want[..]);

        let x = open(&xlsx(1, 1), DocumentFormat::Xlsx, lazy);
        let book = x.as_xlsx().unwrap();
        let ws = &book.worksheets[0];
        assert_eq!(ws.images.len(), 1);
        assert_eq!(&book.picture(&ws.images[0]).bytes().unwrap()[..], &want[..]);
    }
}

/// A picture over the part limit is refused and reported the same way.
#[test]
fn test_a_picture_over_a_limit_is_reported_the_same_lazy_or_not() {
    let _serial = serial();
    for (format, bytes) in formats() {
        let damage = |lazy: bool| {
            let mut o = options(lazy);
            o.limits.part_bytes = (PICTURE_LEN / 2) as u64;
            let d = Document::from_reader_with(Cursor::new(bytes.clone()), format, &o).unwrap();
            (d.damage(), d.plain_text())
        };
        let (eager, lazy) = (damage(false), damage(true));
        assert!(!eager.0.is_empty(), "{format:?}: the oversized pictures are reported");
        assert_eq!(eager, lazy, "{format:?}");
    }
}
