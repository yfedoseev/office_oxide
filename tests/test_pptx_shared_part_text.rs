//! Chart and SmartArt text is stored once in the package and named by every
//! graphic frame that references the part. The reader shares one copy of
//! the lines among the frames, and every copy a conversion or a renderer makes
//! is charged to the text budget before it is made.
//!
//! These tests lower the process-wide text budget and count allocated bytes,
//! so they have a test binary to themselves and run one at a time.
//!
//! Each test holds the serial lock for its whole body.

use std::alloc::{GlobalAlloc, Layout, System};
use std::io::{Cursor, Write};

use office_oxide::ir::{Element, InlineContent};
use office_oxide::limits::{max_text_chars, set_max_text_chars};
use office_oxide::{Document, DocumentFormat};

mod common;
use common::LimitScope;

struct Counting;

/// Bytes requested by the process. Opening a document runs on a helper
/// thread with a large stack, so a per-thread count would miss it; the tests
/// hold `serial()` so nothing else allocates meanwhile.
static BYTES: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

// SAFETY: every call is forwarded to `System` unchanged; the counter is the
// only addition and touches no allocator state.
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        BYTES.fetch_add(layout.size(), std::sync::atomic::Ordering::SeqCst);
        // SAFETY: same layout contract as the caller's.
        unsafe { System.alloc(layout) }
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        // SAFETY: `ptr` came from `System.alloc` with this layout.
        unsafe { System.dealloc(ptr, layout) }
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        BYTES.fetch_add(new_size, std::sync::atomic::Ordering::SeqCst);
        // SAFETY: forwarded unchanged.
        unsafe { System.realloc(ptr, layout, new_size) }
    }
}

#[global_allocator]
static GLOBAL: Counting = Counting;

fn bytes_during<T>(f: impl FnOnce() -> T) -> (T, usize) {
    let before = BYTES.load(std::sync::atomic::Ordering::SeqCst);
    let out = f();
    (out, BYTES.load(std::sync::atomic::Ordering::SeqCst) - before)
}

/// Tests change a process-wide limit and measure allocations: one at a time.
/// A limit set through the scope is back at its default when the test ends,
/// however it ends.
fn serial() -> LimitScope {
    LimitScope::new()
}

const PNS: &str = r#"xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#;
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

fn zip_of(parts: &[(String, Vec<u8>)]) -> Vec<u8> {
    let mut z = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let o: zip::write::FileOptions<'_, '_, ()> =
        zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for (n, d) in parts {
        z.start_file(n.as_str(), o).unwrap();
        z.write_all(d).unwrap();
    }
    z.finish().unwrap().into_inner()
}

/// A chart part of `series` series, each caching `values` values of
/// `value_len` characters: one text line per series ("Series N: v, v, ...").
/// A chart reader caps the values and series it keeps, so a chart is small,
/// but still large enough to matter when copied per reference.
fn chart_part(series: usize, values: usize, value_len: usize) -> Vec<u8> {
    let value = "v".repeat(value_len);
    let points: String = (0..values)
        .map(|i| format!(r#"<c:pt idx="{i}"><c:v>{value}</c:v></c:pt>"#))
        .collect();
    let ser = format!(
        r#"<c:ser><c:val><c:numRef><c:numCache>{points}</c:numCache></c:numRef></c:val></c:ser>"#
    );
    format!(
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart>{}</c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
        ser.repeat(series)
    )
    .into_bytes()
}

const CHART_FRAME: &str = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="1" name="c"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId9"/></a:graphicData></a:graphic></p:graphicFrame>"#;

fn slide_xml(tree: &str) -> Vec<u8> {
    format!(r#"<p:sld {PNS}><p:cSld><p:spTree>{tree}</p:spTree></p:cSld></p:sld>"#).into_bytes()
}

/// A deck of `slides` slides holding `frames` frames each, every frame naming
/// one chart part.
fn deck(slides: usize, frames: usize, chart: Vec<u8>) -> Vec<u8> {
    let slide_rels = format!(
        r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId9" Type="{REL}/chart" Target="../charts/chart1.xml"/></Relationships>"#
    );
    let mut ids = String::from("<p:sldIdLst>");
    let mut pres_rels = format!(r#"<Relationships xmlns="{PKG_REL}">"#);
    let mut ct = String::new();
    let mut parts = vec![("ppt/charts/chart1.xml".to_string(), chart)];
    let tree = CHART_FRAME.repeat(frames);
    for i in 0..slides {
        ids.push_str(&format!(r#"<p:sldId id="{}" r:id="s{i}"/>"#, 256 + i));
        pres_rels.push_str(&format!(
            r#"<Relationship Id="s{i}" Type="{REL}/slide" Target="slides/slide{i}.xml"/>"#
        ));
        ct.push_str(&format!(
            r#"<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#
        ));
        parts.push((format!("ppt/slides/slide{i}.xml"), slide_xml(&tree)));
        parts
            .push((format!("ppt/slides/_rels/slide{i}.xml.rels"), slide_rels.clone().into_bytes()));
    }
    ids.push_str("</p:sldIdLst>");
    pres_rels.push_str("</Relationships>");
    let mut all = vec![
        ("[Content_Types].xml".to_string(), format!(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>{ct}</Types>"#).into_bytes()),
        ("_rels/.rels".to_string(), format!(r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#).into_bytes()),
        ("ppt/_rels/presentation.xml.rels".to_string(), pres_rels.into_bytes()),
        ("ppt/presentation.xml".to_string(), format!(r#"<p:presentation {PNS}>{ids}</p:presentation>"#).into_bytes()),
    ];
    all.extend(parts);
    zip_of(&all)
}

fn open(bytes: Vec<u8>) -> Document {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Pptx).unwrap()
}

fn ir_text_chars(ir: &office_oxide::ir::DocumentIR) -> usize {
    fn walk(elements: &[Element]) -> usize {
        elements
            .iter()
            .map(|e| match e {
                Element::Paragraph(p) => p
                    .content
                    .iter()
                    .map(|c| match c {
                        InlineContent::Text(t) => t.text.chars().count(),
                        _ => 0,
                    })
                    .sum(),
                Element::TextBox(tb) => walk(&tb.content),
                _ => 0,
            })
            .sum()
    }
    ir.sections.iter().map(|s| walk(&s.elements)).sum()
}

#[test]
fn test_frames_naming_one_chart_part_share_its_text_at_open() {
    let _g = serial();
    // About 100 KB of chart text: 200 frames would be 20 MB if each held a copy.
    let bytes = deck(1, 200, chart_part(1, 4_000, 25));
    let (doc, allocated) = bytes_during(|| open(bytes));
    drop(doc);
    assert!(
        allocated < 3_000_000,
        "opening allocated {allocated} bytes: the chart text is copied per frame"
    );
}

#[test]
fn test_slides_naming_one_chart_part_share_its_text_at_open() {
    let _g = serial();
    // Each slide costs its own parsing buffers whatever it names; what the
    // chart text adds must not grow with the number of slides naming it.
    let small = deck(200, 1, chart_part(1, 1, 1));
    let (doc, baseline) = bytes_during(|| open(small));
    drop(doc);
    let big = deck(200, 1, chart_part(1, 4_000, 25));
    let (doc, allocated) = bytes_during(|| open(big));
    drop(doc);
    assert!(
        allocated < baseline + 3_000_000,
        "opening allocated {allocated} bytes against {baseline} for tiny charts: \
         the chart text is copied per slide"
    );
}

#[test]
fn test_ir_charges_every_frames_copy_of_chart_text() {
    let mut limits = serial();
    limits.set(max_text_chars, set_max_text_chars, 20_000);
    let doc = open(deck(1, 100, chart_part(1, 100, 8)));
    let ir = doc.to_ir();
    assert!(ir.metadata.text_truncated, "the cut must be flagged");
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("set_max_text_chars")),
        "{:?}",
        ir.metadata.warnings
    );
    // 100 frames of about 1,000 characters is 100,000; the budget allows a fifth.
    let kept = ir_text_chars(&ir);
    assert!(kept <= 20_000, "{kept} characters kept");
    assert!(kept >= 1_000, "the first copies fit the budget: {kept}");
}

#[test]
fn test_ir_charges_chart_text_copies_made_by_many_slides() {
    let mut limits = serial();
    limits.set(max_text_chars, set_max_text_chars, 20_000);
    let doc = open(deck(100, 1, chart_part(1, 100, 8)));
    let ir = doc.to_ir();
    assert!(ir.metadata.text_truncated);
    assert!(ir_text_chars(&ir) <= 20_000);
}

#[test]
fn test_ir_charges_a_copy_per_line_not_only_its_characters() {
    let mut limits = serial();
    // A SmartArt-like part of many one-character lines costs a paragraph
    // each in the IR.
    limits.set(max_text_chars, set_max_text_chars, 20_000);
    let doc = open(deck(1, 50, chart_part(512, 1, 1)));
    let ir = doc.to_ir();
    assert!(ir.metadata.text_truncated);
}

#[test]
fn test_plain_text_and_markdown_charge_chart_text_before_copying_it() {
    let mut limits = serial();
    limits.set(max_text_chars, set_max_text_chars, 300_000);
    let doc = open(deck(1, 200, chart_part(1, 4_000, 25)));
    let (plain, plain_bytes) = bytes_during(|| doc.plain_text());
    assert!(plain.contains("set_max_text_chars"), "plain text is not marked cut");
    assert!(
        plain_bytes < 4_000_000,
        "plain_text allocated {plain_bytes} bytes for a 300,000 character budget"
    );
    let (md, md_bytes) = bytes_during(|| doc.to_markdown());
    assert!(md.contains("set_max_text_chars"), "markdown is not marked cut");
    assert!(
        md_bytes < 4_000_000,
        "to_markdown allocated {md_bytes} bytes for a 300,000 character budget"
    );
}

#[test]
fn test_chart_text_within_budget_is_unchanged() {
    let _g = serial();
    let doc = open(deck(2, 3, chart_part(1, 2, 3)));
    let ir = doc.to_ir();
    assert!(!ir.metadata.text_truncated);
    // "Series 1: vvv, vvv" per frame.
    assert_eq!(ir_text_chars(&ir), 2 * 3 * "Series 1: vvv, vvv".len());
    assert_eq!(doc.plain_text().matches("Series 1: vvv, vvv").count(), 6);
}

/// One slide with a body placeholder of `paragraphs` paragraphs whose master
/// gives the body style a bullet character and an underline kind of `value`.
/// The paragraphs inherit both, so a reader that copies the inherited values
/// holds `value` once per paragraph.
fn inheriting_deck(value: &str, paragraphs: usize) -> Vec<u8> {
    let ct = |name: &str, kind: &str| {
        format!(
            r#"<Override PartName="/ppt/{name}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.{kind}+xml"/>"#
        )
    };
    let body = "<a:p><a:r><a:t>x</a:t></a:r></a:p>".repeat(paragraphs);
    let slide = format!(
        r#"<p:sld {PNS}><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="B"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/>{body}</p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#
    );
    let layout = format!(r#"<p:sldLayout {PNS}><p:cSld><p:spTree/></p:cSld></p:sldLayout>"#);
    let master = format!(
        r#"<p:sldMaster {PNS}><p:cSld><p:spTree/></p:cSld><p:txStyles><p:bodyStyle><a:lvl1pPr><a:buChar char="{value}"/><a:defRPr u="{value}"/></a:lvl1pPr></p:bodyStyle></p:txStyles></p:sldMaster>"#
    );
    let rels = |entries: &[(&str, &str, &str)]| {
        let mut s = format!(r#"<Relationships xmlns="{PKG_REL}">"#);
        for (id, ty, target) in entries {
            s.push_str(&format!(
                r#"<Relationship Id="{id}" Type="{REL}/{ty}" Target="{target}"/>"#
            ));
        }
        s.push_str("</Relationships>");
        s.into_bytes()
    };
    zip_of(&[
        ("[Content_Types].xml".to_string(), format!(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>{}{}{}{}</Types>"#, ct("presentation", "presentation.main"), ct("slides/slide1", "slide"), ct("slideLayouts/slideLayout1", "slideLayout"), ct("slideMasters/slideMaster1", "slideMaster")).into_bytes()),
        ("_rels/.rels".to_string(), rels(&[("rId1", "officeDocument", "ppt/presentation.xml")])),
        ("ppt/presentation.xml".to_string(), format!(r#"<p:presentation {PNS}><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>"#).into_bytes()),
        ("ppt/_rels/presentation.xml.rels".to_string(), rels(&[("rId2", "slide", "slides/slide1.xml")])),
        ("ppt/slides/slide1.xml".to_string(), slide.into_bytes()),
        ("ppt/slides/_rels/slide1.xml.rels".to_string(), rels(&[("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml")])),
        ("ppt/slideLayouts/slideLayout1.xml".to_string(), layout.into_bytes()),
        ("ppt/slideLayouts/_rels/slideLayout1.xml.rels".to_string(), rels(&[("rId1", "slideMaster", "../slideMasters/slideMaster1.xml")])),
        ("ppt/slideMasters/slideMaster1.xml".to_string(), master.into_bytes()),
    ])
}

#[test]
fn test_paragraphs_inheriting_a_master_value_do_not_each_hold_a_copy() {
    let _g = serial();
    // 300 paragraphs x 2 x 50 KB would be 30 MB if each held its own copy.
    let big = inheriting_deck(&"u".repeat(50_000), 300);
    let (doc, allocated) = bytes_during(|| open(big));
    drop(doc);
    assert!(
        allocated < 3_000_000,
        "opening allocated {allocated} bytes: an inherited style value is copied per paragraph"
    );
}

#[test]
fn test_ordinary_inherited_bullet_and_underline_are_unchanged() {
    let _g = serial();
    let doc = open(inheriting_deck("sng", 2));
    let p = doc.as_pptx().expect("a pptx");
    let office_oxide::pptx::Shape::AutoShape(shape) = &p.slides[0].shapes[0] else {
        panic!("a shape")
    };
    let body = shape.text_body.as_ref().expect("a body");
    assert_eq!(body.paragraphs.len(), 2);
    for para in &body.paragraphs {
        assert!(
            matches!(&para.bullet, Some(office_oxide::pptx::BulletStyle::Char(c)) if c == "sng")
        );
    }
}
