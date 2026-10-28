//! Tables that depend on the workbook alone (which cell styles are dates, the
//! custom number formats) are built once, not once per sheet: a workbook of
//! many sheets over a stylesheet of many cell formats cost their product in
//! every per-sheet and whole-workbook rendering.

use std::io::{Cursor, Write};
use std::time::{Duration, Instant};

use office_oxide::{Document, DocumentFormat};

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// `sheets` one-cell sheets over a stylesheet of `formats` custom date formats.
fn workbook(sheets: usize, formats: usize) -> Vec<u8> {
    let mut wb = format!(r#"<workbook xmlns="{NS}" xmlns:r="{R}"><sheets>"#);
    let mut rels = format!(r#"<Relationships xmlns="{REL_NS}">"#);
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts: zip::write::FileOptions<'_, '_, ()> = zip::write::FileOptions::default();
    for i in 1..=sheets {
        wb.push_str(&format!(r#"<sheet name="S{i}" sheetId="{i}" r:id="rId{i}"/>"#));
        rels.push_str(&format!(
            r#"<Relationship Id="rId{i}" Type="{R}/worksheet" Target="worksheets/sheet{i}.xml"/>"#
        ));
        zip.start_file(format!("xl/worksheets/sheet{i}.xml"), opts)
            .unwrap();
        zip.write_all(
            format!(
                r#"<worksheet xmlns="{NS}"><sheetData><row r="1"><c r="A1" s="{}"><v>{}</v></c></row></sheetData></worksheet>"#,
                i % formats + 1,
                38000 + i
            )
            .as_bytes(),
        )
        .unwrap();
    }
    wb.push_str("</sheets></workbook>");
    rels.push_str(&format!(
        r#"<Relationship Id="rIdS" Type="{R}/styles" Target="styles.xml"/></Relationships>"#
    ));
    let mut st = format!(r#"<styleSheet xmlns="{NS}"><numFmts count="{formats}">"#);
    for i in 0..formats {
        st.push_str(&format!(
            r#"<numFmt numFmtId="{}" formatCode="[$-409]dd/mm/yyyy hh:mm;@;0.0{i}"/>"#,
            164 + i
        ));
    }
    st.push_str(&format!(r#"</numFmts><cellXfs count="{}"><xf numFmtId="0"/>"#, formats + 1));
    for i in 0..formats {
        st.push_str(&format!(r#"<xf numFmtId="{}" applyNumberFormat="1"/>"#, 164 + i));
    }
    st.push_str("</cellXfs></styleSheet>");
    for (name, data) in [
        ("xl/workbook.xml", wb),
        ("xl/_rels/workbook.xml.rels", rels),
        ("xl/styles.xml", st),
    ] {
        zip.start_file(name, opts).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn timed<T>(f: impl FnOnce() -> T) -> (T, Duration) {
    let start = Instant::now();
    let out = f();
    (out, start.elapsed())
}

#[test]
fn test_per_sheet_and_whole_workbook_renderings_do_not_rebuild_workbook_tables_per_sheet() {
    let doc =
        Document::from_reader(Cursor::new(workbook(400, 20_000)), DocumentFormat::Xlsx).unwrap();
    let xlsx = doc.as_xlsx().unwrap();
    // Built per sheet, each of these costs seconds (debug) at this size.
    let (_, per_sheet) = timed(|| {
        for i in 0..400 {
            assert!(xlsx.sheet_ir(i).is_some());
        }
    });
    let (md, whole) = timed(|| doc.to_markdown());
    let (text, plain) = timed(|| doc.plain_text());
    assert!(md.contains("S400") && text.contains("S400"));
    for (what, took) in [
        ("sheet_ir", per_sheet),
        ("to_markdown", whole),
        ("plain_text", plain),
    ] {
        assert!(took < Duration::from_secs(8), "{what} took {took:?}");
    }
}
