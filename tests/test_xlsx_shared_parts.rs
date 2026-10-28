//! A part stored once in a spreadsheet package can be named by every sheet:
//! a drawing, a comments part, a table definition. Each sheet used to read,
//! parse and copy the part for itself, so a package of a few hundred
//! kilobytes naming one large part from hundreds of sheets cost gigabytes.
//!
//! The part is parsed once and shared; what each sheet's use of it adds to
//! the IR and to the renderings is charged to the document's budgets before
//! it is copied. The heap counter makes the first deterministic: bytes
//! requested by the opening thread are compared between one sheet and many.
//!
//! The counter is a global allocator and the limits are process-global, so
//! this file is its own test binary and every test holds the
//! `LimitScope` of [`serial`].

use std::alloc::{GlobalAlloc, Layout, System};
use std::io::{Cursor, Write};

use office_oxide::limits::{max_text_chars, set_max_text_chars};
use office_oxide::{Document, DocumentFormat};

mod common;
use common::LimitScope;

struct Counting;

/// Bytes requested by the whole process. Opening a document runs on a thread
/// of its own, so a per-thread count would see none of it.
static BYTES: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn count_bytes(n: usize) {
    BYTES.fetch_add(n, std::sync::atomic::Ordering::Relaxed);
}

// SAFETY: every call is forwarded to `System` unchanged; the counter is the
// only addition and touches no allocator state.
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        count_bytes(layout.size());
        // SAFETY: same layout contract as the caller's.
        unsafe { System.alloc(layout) }
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        // SAFETY: `ptr` came from `System.alloc` with this layout.
        unsafe { System.dealloc(ptr, layout) }
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        count_bytes(new_size);
        // SAFETY: forwarded unchanged.
        unsafe { System.realloc(ptr, layout, new_size) }
    }
}

#[global_allocator]
static GLOBAL: Counting = Counting;

fn bytes_during<T>(f: impl FnOnce() -> T) -> (T, usize) {
    let before = BYTES.load(std::sync::atomic::Ordering::Relaxed);
    let out = f();
    (out, BYTES.load(std::sync::atomic::Ordering::Relaxed) - before)
}

/// Holds the process-global limits, at their defaults, for one test; a limit
/// the test sets through it is back at its default when the test ends.
fn serial() -> LimitScope {
    LimitScope::new()
}

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// A package of `sheets` sheets, each with the same `sheet_xml` and, when
/// `sheet_rels` is not empty, a relationships part of its own holding it.
/// `extra` are the parts the sheets share.
fn workbook(
    sheets: usize,
    sheet_xml: &str,
    sheet_rels: &str,
    extra: &[(&str, Vec<u8>)],
) -> Vec<u8> {
    workbook_of("worksheet", sheets, sheet_xml, sheet_rels, extra)
}

/// Like [`workbook`], with sheets of the relationship type `kind`
/// (`worksheet` or `chartsheet`); their parts are in `<kind>s/`.
fn workbook_of(
    kind: &str,
    sheets: usize,
    sheet_xml: &str,
    sheet_rels: &str,
    extra: &[(&str, Vec<u8>)],
) -> Vec<u8> {
    let mut names = String::new();
    let mut rels = String::new();
    let mut parts: Vec<(String, Vec<u8>)> = Vec::new();
    for i in 1..=sheets {
        names.push_str(&format!(r#"<sheet name="S{i}" sheetId="{i}" r:id="rId{i}"/>"#));
        rels.push_str(&format!(
            r#"<Relationship Id="rId{i}" Type="{R}/{kind}" Target="{kind}s/sheet{i}.xml"/>"#
        ));
        parts.push((format!("xl/{kind}s/sheet{i}.xml"), sheet_xml.as_bytes().to_vec()));
        if !sheet_rels.is_empty() {
            parts.push((
                format!("xl/{kind}s/_rels/sheet{i}.xml.rels"),
                format!(r#"<Relationships xmlns="{REL_NS}">{sheet_rels}</Relationships>"#)
                    .into_bytes(),
            ));
        }
    }
    let fixed: Vec<(String, Vec<u8>)> = vec![
        (
            "[Content_Types].xml".into(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".into(),
            format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#).into_bytes(),
        ),
        (
            "xl/workbook.xml".into(),
            format!(r#"<workbook xmlns="{NS}" xmlns:r="{R}"><sheets>{names}</sheets></workbook>"#).into_bytes(),
        ),
        (
            "xl/_rels/workbook.xml.rels".into(),
            format!(r#"<Relationships xmlns="{REL_NS}">{rels}</Relationships>"#).into_bytes(),
        ),
    ];
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts: zip::write::FileOptions<'_, '_, ()> =
        zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for (name, data) in fixed.iter().chain(parts.iter()) {
        zip.start_file(name.as_str(), opts).unwrap();
        zip.write_all(data).unwrap();
    }
    for (name, data) in extra {
        zip.start_file(*name, opts).unwrap();
        zip.write_all(data).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn open(bytes: Vec<u8>) -> Document {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Xlsx).unwrap()
}

/// The sheet body every sheet shares: one cell.
fn one_cell(tail: &str) -> String {
    format!(
        r#"<worksheet xmlns="{NS}" xmlns:r="{R}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>{tail}</worksheet>"#
    )
}

const XDR: &str = r#"xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;

/// A drawing of `anchors` pictures (or text shapes) each carrying `len`
/// characters of text, and the parts it names.
fn drawing_parts(anchors: usize, len: usize, shapes: bool) -> Vec<(&'static str, Vec<u8>)> {
    let text = "d".repeat(len);
    let mut d = format!("<xdr:wsDr {XDR}>");
    for i in 0..anchors {
        let from = "<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"100\" cy=\"100\"/>";
        if shapes {
            d.push_str(&format!(r#"<xdr:oneCellAnchor>{from}<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="{i}" name="t"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/><xdr:txBody><a:bodyPr/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>"#));
        } else {
            d.push_str(&format!(r#"<xdr:oneCellAnchor>{from}<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{i}" name="p" descr="{text}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>"#));
        }
    }
    d.push_str("</xdr:wsDr>");
    vec![
        ("xl/drawings/drawing1.xml", d.into_bytes()),
        (
            "xl/drawings/_rels/drawing1.xml.rels",
            format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/image" Target="../media/i.png"/></Relationships>"#).into_bytes(),
        ),
        ("xl/media/i.png", vec![0u8; 16]),
    ]
}

/// A relationship to nothing the reader uses, so that a sheet that does not
/// name the shared part still has a relationships part to read.
fn unrelated_rel() -> String {
    format!(
        r#"<Relationship Id="rId9" Type="{R}/printerSettings" Target="../printerSettings/p.bin"/>"#
    )
}

/// The part is in the package either way; `linked` says whether the sheets
/// name it.
fn shared_drawing(
    sheets: usize,
    anchors: usize,
    len: usize,
    shapes: bool,
    linked: bool,
) -> Vec<u8> {
    workbook(
        sheets,
        &one_cell(if linked {
            r#"<drawing r:id="rId1"/>"#
        } else {
            ""
        }),
        &if linked {
            format!(
                r#"<Relationship Id="rId1" Type="{R}/drawing" Target="../drawings/drawing1.xml"/>"#
            )
        } else {
            unrelated_rel()
        },
        &drawing_parts(anchors, len, shapes),
    )
}

fn shared_comments(sheets: usize, comments: usize, len: usize, linked: bool) -> Vec<u8> {
    let text = "c".repeat(len);
    let mut c =
        format!(r#"<comments xmlns="{NS}"><authors><author>a</author></authors><commentList>"#);
    for i in 0..comments {
        c.push_str(&format!(
            r#"<comment ref="A{}" authorId="0"><text><r><t>{text}</t></r></text></comment>"#,
            i + 1
        ));
    }
    c.push_str("</commentList></comments>");
    workbook(
        sheets,
        &one_cell(""),
        &if linked {
            format!(r#"<Relationship Id="rId1" Type="{R}/comments" Target="../comments1.xml"/>"#)
        } else {
            unrelated_rel()
        },
        &[("xl/comments1.xml", c.into_bytes())],
    )
}

fn shared_table(sheets: usize, columns: usize, len: usize, linked: bool) -> Vec<u8> {
    let mut t = format!(
        r#"<table xmlns="{NS}" id="1" name="T" displayName="T" ref="A1:B2"><tableColumns count="{columns}">"#
    );
    for i in 0..columns {
        t.push_str(&format!(r#"<tableColumn id="{}" name="{}{i}"/>"#, i + 1, "n".repeat(len)));
    }
    t.push_str("</tableColumns></table>");
    workbook(
        sheets,
        &one_cell(if linked {
            r#"<tableParts count="1"><tablePart r:id="rId1"/></tableParts>"#
        } else {
            ""
        }),
        &if linked {
            format!(r#"<Relationship Id="rId1" Type="{R}/table" Target="../tables/table1.xml"/>"#)
        } else {
            unrelated_rel()
        },
        &[("xl/tables/table1.xml", t.into_bytes())],
    )
}

/// Bytes requested while opening `build(sheets, linked)`.
fn open_bytes(build: &impl Fn(usize, bool) -> Vec<u8>, sheets: usize, linked: bool) -> usize {
    let bytes = build(sheets, linked);
    let (doc, used) = bytes_during(|| open(bytes));
    drop(doc);
    used
}

/// What naming the part from many sheets adds to opening is what naming it
/// from one does (it is parsed once), not the sheet count times as much.
fn assert_opening_is_shared(build: impl Fn(usize, bool) -> Vec<u8>, what: &str) {
    let once =
        open_bytes(&build, 1, true) - open_bytes(&build, 1, false).min(open_bytes(&build, 1, true));
    let many = open_bytes(&build, 30, true).saturating_sub(open_bytes(&build, 30, false));
    assert!(
        once > 100_000,
        "{what}: the part costs only {once} bytes to parse: too small to measure"
    );
    assert!(
        many < once * 3,
        "{what}: naming the part from 30 sheets adds {many} bytes to opening, from one sheet {once}: it is parsed per sheet"
    );
}

#[test]
fn test_a_drawing_named_by_many_sheets_is_parsed_once() {
    let _g = serial();
    assert_opening_is_shared(|n, l| shared_drawing(n, 600, 200, false, l), "pictures");
    assert_opening_is_shared(|n, l| shared_drawing(n, 600, 200, true, l), "text shapes");
}

#[test]
fn test_a_comments_part_named_by_many_sheets_is_parsed_once() {
    let _g = serial();
    assert_opening_is_shared(|n, l| shared_comments(n, 600, 200, l), "comments");
}

#[test]
fn test_a_table_part_named_by_many_sheets_is_parsed_once() {
    let _g = serial();
    assert_opening_is_shared(|n, l| shared_table(n, 600, 200, l), "table");
}

/// The characters of text the IR holds: every text run, alt text and note.
fn ir_chars(ir: &office_oxide::ir::DocumentIR) -> usize {
    serde_json::to_string(ir).unwrap().len()
}

#[test]
fn test_each_sheets_use_of_a_shared_drawing_is_charged_to_the_text_budget() {
    let mut limits = serial();
    limits.set(max_text_chars, set_max_text_chars, 150_000);
    // 20 sheets x 300 anchors x 200 characters is 1.2 million characters of
    // alternative text; the budget holds 150,000.
    for shapes in [false, true] {
        let doc = open(shared_drawing(20, 300, 200, shapes, true));
        let ir = doc.to_ir();
        assert!(ir.metadata.text_truncated, "shapes={shapes}: the cut is not flagged");
        assert!(
            ir.metadata
                .warnings
                .iter()
                .any(|w| w.contains("set_max_text_chars")),
            "shapes={shapes}: no warning: {:?}",
            ir.metadata.warnings
        );
        let size = ir_chars(&ir);
        assert!(size < 1_000_000, "shapes={shapes}: the IR holds {size} characters");
        // The sheet-by-sheet IR is bound by the same budget for the run.
        let total: usize = doc
            .as_xlsx()
            .unwrap()
            .sheets_ir()
            .map(|s| serde_json::to_string(&s.section).unwrap().len())
            .sum();
        assert!(total < 1_000_000, "shapes={shapes}: sheets_ir yields {total} characters");
    }
}

#[test]
fn test_a_workbook_within_the_budget_keeps_every_picture_and_shape() {
    let _g = serial();
    for shapes in [false, true] {
        let doc = open(shared_drawing(3, 20, 50, shapes, true));
        let ir = doc.to_ir();
        assert!(!ir.metadata.text_truncated, "{:?}", ir.metadata.warnings);
        let json = serde_json::to_string(&ir).unwrap();
        assert_eq!(json.matches(&"d".repeat(50)).count(), 60, "shapes={shapes}");
    }
}

#[test]
fn test_each_sheets_use_of_shared_comments_is_charged_to_the_text_budget() {
    let mut limits = serial();
    limits.set(max_text_chars, set_max_text_chars, 150_000);
    let doc = open(shared_comments(20, 300, 200, true));
    let ir = doc.to_ir();
    assert!(ir.metadata.text_truncated);
    assert!(ir_chars(&ir) < 1_000_000, "{}", ir_chars(&ir));
}

#[test]
fn test_a_comment_of_one_character_is_charged_as_the_note_that_holds_it() {
    let mut limits = serial();
    limits.set(max_text_chars, set_max_text_chars, 150_000);
    // 20 sheets x 300 one-character comments: the characters alone are far
    // inside the budget, the notes that hold them are not.
    let doc = open(shared_comments(20, 300, 1, true));
    let ir = doc.to_ir();
    assert!(ir.metadata.text_truncated, "{:?}", ir.metadata.warnings);
    let notes = serde_json::to_string(&ir)
        .unwrap()
        .matches("\"endnote\"")
        .count();
    assert!((1..2_000).contains(&notes), "{notes} notes for a budget of 150,000 characters");
    // The direct renderers agree on the cut.
    assert!(doc.plain_text().contains("truncated"));
    assert!(doc.to_markdown().contains("truncated"));
}

/// `sheets` chartsheets that display one chart whose title is `len`
/// characters; `linked` says whether they name its drawing.
fn shared_chartsheets(sheets: usize, len: usize, linked: bool) -> Vec<u8> {
    shared_chartsheets_padded(sheets, len, linked, 0)
}

/// Like [`shared_chartsheets`], with `pad` bytes of padding in the drawing's
/// relationships part.
fn shared_chartsheets_padded(sheets: usize, len: usize, linked: bool, pad: usize) -> Vec<u8> {
    let chart = format!(
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>{}</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>"#,
        "t".repeat(len)
    );
    let drawing = format!(
        r#"<xdr:wsDr {XDR}><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="100" cy="100"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="c"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>"#
    );
    workbook_of(
        "chartsheet",
        sheets,
        &if linked {
            format!(r#"<chartsheet xmlns="{NS}" xmlns:r="{R}"><drawing r:id="rId1"/></chartsheet>"#)
        } else {
            format!(r#"<chartsheet xmlns="{NS}" xmlns:r="{R}"/>"#)
        },
        &if linked {
            format!(r#"<Relationship Id="rId1" Type="{R}/drawing" Target="../drawings/drawing1.xml"/>"#)
        } else {
            unrelated_rel()
        },
        &[
            ("xl/drawings/drawing1.xml", drawing.into_bytes()),
            (
                "xl/drawings/_rels/drawing1.xml.rels",
                format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/chart" Target="../charts/chart1.xml"/><!--{}--></Relationships>"#, " ".repeat(pad)).into_bytes(),
            ),
            ("xl/charts/chart1.xml", chart.into_bytes()),
        ],
    )
}

#[test]
fn test_a_chart_displayed_by_many_chartsheets_is_read_once_and_charged() {
    let mut limits = serial();
    // A chartsheet's chart is read for the sheet, so the part is in the
    // package either way and only the cost of many sheets over one is
    // compared: the sheets' own cost, from a chart of ten characters, and
    // what the large chart adds to it.
    let by = |sheets, len| open_bytes(&|n, l| shared_chartsheets(n, len, l), sheets, true);
    let per_sheets = by(30, 10).saturating_sub(by(1, 10));
    let large = 2_000_000;
    let added = by(30, large).saturating_sub(by(1, large));
    assert!(
        added < per_sheets + large,
        "30 chartsheets add {added} bytes to opening over one, {per_sheets} without the chart: the chart is read per sheet"
    );
    // 20 sheets of a 20,000-character title against a budget of 150,000.
    limits.set(max_text_chars, set_max_text_chars, 150_000);
    let doc = open(shared_chartsheets(20, 20_000, true));
    let ir = doc.to_ir();
    assert!(ir.metadata.text_truncated, "{:?}", ir.metadata.warnings);
    assert!(ir_chars(&ir) < 400_000, "{}", ir_chars(&ir));
    assert!(doc.plain_text().len() < 400_000);
    assert!(doc.to_markdown().len() < 400_000);
}

/// One sheet with `pivots` pivot tables, all naming one cache definition of
/// about `pad` bytes whose source is not a worksheet (so the early return for
/// a source in the workbook does not apply). `linked` says whether the pivot
/// tables name it.
fn shared_pivot_definition(pivots: usize, pad: usize, linked: bool) -> Vec<u8> {
    let def = format!(
        r#"<pivotCacheDefinition xmlns="{NS}"><cacheSource type="consolidation"/><!--{}--></pivotCacheDefinition>"#,
        " ".repeat(pad)
    );
    let mut extra: Vec<(String, Vec<u8>)> =
        vec![("xl/pivotCache/pivotCacheDefinition1.xml".to_string(), def.into_bytes())];
    let mut rels = String::new();
    for i in 1..=pivots {
        extra.push((
            format!("xl/pivotTables/pivotTable{i}.xml"),
            format!(
                r#"<pivotTableDefinition xmlns="{NS}" name="p{i}"><location ref="A1:B2"/></pivotTableDefinition>"#
            )
            .into_bytes(),
        ));
        extra.push((
            format!("xl/pivotTables/_rels/pivotTable{i}.xml.rels"),
            format!(
                r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/{}" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>"#,
                if linked { "pivotCacheDefinition" } else { "unrelated" }
            )
            .into_bytes(),
        ));
        rels.push_str(&format!(
            r#"<Relationship Id="rId{i}" Type="{R}/pivotTable" Target="../pivotTables/pivotTable{i}.xml"/>"#
        ));
    }
    let extra: Vec<(&str, Vec<u8>)> = extra.iter().map(|(n, d)| (n.as_str(), d.clone())).collect();
    workbook(1, &one_cell(""), &rels, &extra)
}

#[test]
fn test_a_pivot_cache_definition_named_by_many_pivot_tables_is_read_once() {
    let _g = serial();
    // Opening with more pivot tables adds what the pivot tables themselves
    // cost, not another copy of the definition for each: one 2 MB definition
    // is read and scanned once whatever the number of tables that name it.
    let pad = 2_000_000;
    let open = |p, linked| {
        let bytes = shared_pivot_definition(p, pad, linked);
        let (doc, used) = bytes_during(|| open(bytes));
        drop(doc);
        used
    };
    let tables_alone = open(40, false).saturating_sub(open(1, false));
    let with_definition = open(40, true).saturating_sub(open(1, true));
    assert!(
        with_definition < tables_alone + pad,
        "40 pivot tables naming one {pad} byte definition add {with_definition} bytes to opening over one table, {tables_alone} when they do not name it: the definition is read per table"
    );
}

#[test]
fn test_the_relationships_of_a_drawing_named_by_many_chartsheets_are_read_once() {
    let _g = serial();
    // The drawing's relationships part, not the chart, is the large one: it
    // is read for the first chartsheet that names the drawing and not again.
    let pad = 2_000_000;
    let by = |sheets, linked| {
        let bytes = shared_chartsheets_padded(sheets, 10, linked, pad);
        let (doc, used) = bytes_during(|| open(bytes));
        drop(doc);
        used
    };
    let unlinked = by(30, false).saturating_sub(by(1, false));
    let linked = by(30, true).saturating_sub(by(1, true));
    assert!(
        linked < unlinked + pad,
        "30 chartsheets naming a drawing whose relationships are {pad} bytes add {linked} bytes to opening over one, {unlinked} when they do not name it: the relationships are read per chartsheet"
    );
}

/// A package of `t` threaded-comment parts of `entries` messages each and `l`
/// small legacy comments parts, with one sheet for each of the `t * l` pairs.
fn comment_pairs(t: usize, l: usize, entries: usize) -> Vec<u8> {
    const THREADED: &str =
        "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
    let mut parts: Vec<(String, String)> = Vec::new();
    let mut names = String::new();
    let mut wb_rels = String::new();
    let mut n = 0;
    for j in 0..t {
        let mut s = String::from(
            r#"<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">"#,
        );
        for i in 0..entries {
            s.push_str(&format!(
                r#"<threadedComment ref="A{}" id="{{{i}}}"><text>x</text></threadedComment>"#,
                i + 1
            ));
        }
        s.push_str("</ThreadedComments>");
        parts.push((format!("xl/threadedComments/t{j}.xml"), s));
    }
    for i in 0..l {
        parts.push((
            format!("xl/comments{i}.xml"),
            format!(
                r#"<comments xmlns="{NS}"><authors><author>a</author></authors><commentList><comment ref="B1" authorId="0"><text><t>y</t></text></comment></commentList></comments>"#
            ),
        ));
        for j in 0..t {
            n += 1;
            names.push_str(&format!(r#"<sheet name="S{n}" sheetId="{n}" r:id="rId{n}"/>"#));
            wb_rels.push_str(&format!(
                r#"<Relationship Id="rId{n}" Type="{R}/worksheet" Target="worksheets/sheet{n}.xml"/>"#
            ));
            parts.push((format!("xl/worksheets/sheet{n}.xml"), one_cell("")));
            parts.push((
                format!("xl/worksheets/_rels/sheet{n}.xml.rels"),
                format!(
                    r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/comments" Target="../comments{i}.xml"/><Relationship Id="rId2" Type="{THREADED}" Target="../threadedComments/t{j}.xml"/></Relationships>"#
                ),
            ));
        }
    }
    parts.push((
        "[Content_Types].xml".into(),
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#.into(),
    ));
    parts.push((
        "_rels/.rels".into(),
        format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#),
    ));
    parts.push((
        "xl/workbook.xml".into(),
        format!(r#"<workbook xmlns="{NS}" xmlns:r="{R}"><sheets>{names}</sheets></workbook>"#),
    ));
    parts.push((
        "xl/_rels/workbook.xml.rels".into(),
        format!(r#"<Relationships xmlns="{REL_NS}">{wb_rels}</Relationships>"#),
    ));
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts: zip::write::FileOptions<'_, '_, ()> =
        zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for (name, data) in parts {
        zip.start_file(name, opts).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

#[test]
fn test_large_threaded_parts_paired_with_many_legacy_parts_cost_nothing_once_the_budget_is_spent() {
    let mut limits = serial();
    // The first pairs spend the comment allowance; every later pair then has
    // nothing to merge and must not even group the threaded part's messages.
    limits.set(max_text_chars, set_max_text_chars, 1_000);
    // What each further legacy part adds to opening, against two threaded
    // parts of `entries` messages: the sheets' own cost, from messages of
    // few, and what the large parts add to it.
    let per_legacy_part = |entries| {
        let by = |l| {
            let bytes = comment_pairs(2, l, entries);
            let (doc, used) = bytes_during(|| open(bytes));
            assert!(doc.to_ir().metadata.text_truncated);
            drop(doc);
            used
        };
        by(22).saturating_sub(by(2)) / 20
    };
    let (small, large) = (per_legacy_part(20), per_legacy_part(5_000));
    assert!(
        large < small + 5_000 * 8,
        "each further legacy part costs {large} bytes to open against two threaded parts of 5,000 messages, {small} against two of 20: the messages are grouped per pair"
    );
}
