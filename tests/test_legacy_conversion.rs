//! The public conversion surface end to end: legacy `.ppt` through
//! `Document::from_reader` to the IR, `Document::save_as` from each legacy
//! format to its OOXML counterpart, and an `.xlsb` package through the
//! unified API. Every fixture is synthesised in code.

mod common;
#[path = "common/ppt_builder.rs"]
mod ppt_builder;

use std::io::{Cursor, Write};

use common::{Para, biff, build_doc, cfb_with_stream, prose_grpprl};
use office_oxide::{Document, DocumentFormat};

/// A per-test scratch directory, removed on drop.
struct Scratch(std::path::PathBuf);

impl Scratch {
    fn new(tag: &str) -> Self {
        let dir = std::env::temp_dir().join(format!("oo-{tag}-{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        Self(dir)
    }
    fn path(&self, name: &str) -> std::path::PathBuf {
        self.0.join(name)
    }
}

impl Drop for Scratch {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}

/// A BIFF8 workbook with one sheet `S` whose A1 is a LABEL cell `text`.
fn build_xls(text: &str) -> Vec<u8> {
    let bof = |kind: u16| {
        let mut b = 0x0600u16.to_le_bytes().to_vec();
        b.extend_from_slice(&kind.to_le_bytes());
        b.extend_from_slice(&[0u8; 12]);
        biff(0x0809, &b)
    };
    let mut s = bof(0x0005);
    let mut bs = 0u32.to_le_bytes().to_vec();
    bs.extend_from_slice(&[0, 0, 1, 0, b'S']);
    s.extend(biff(0x0085, &bs));
    s.extend(biff(0x000A, &[]));
    s.extend(bof(0x0010));
    let mut label = 0u16.to_le_bytes().to_vec();
    label.extend_from_slice(&0u16.to_le_bytes());
    label.extend_from_slice(&0u16.to_le_bytes());
    label.extend_from_slice(&(text.len() as u16).to_le_bytes());
    label.push(0);
    label.extend_from_slice(text.as_bytes());
    s.extend(biff(0x0204, &label));
    s.extend(biff(0x000A, &[]));
    cfb_with_stream("Workbook", &s)
}

#[test]
fn test_ppt_reaches_the_ir_through_the_unified_api() {
    let bytes = ppt_builder::build_ppt(&["First slide text", "Second slide text"], &[]);
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Ppt).expect("opens");
    assert_eq!(doc.format(), DocumentFormat::Ppt);
    let text = doc.plain_text();
    assert!(text.contains("First slide text"), "{text:?}");
    assert!(text.contains("Second slide text"), "{text:?}");
    let first = text.find("First").unwrap();
    let second = text.find("Second").unwrap();
    assert!(first < second, "slide order: {text:?}");

    let ir = doc.to_ir();
    assert_eq!(ir.metadata.format, DocumentFormat::Ppt);
    assert_eq!(ir.sections.len(), 2, "one IR section per slide");
    assert!(doc.to_markdown().contains("Second slide text"));
}

#[test]
fn test_compound_file_without_a_powerpoint_stream_is_an_error() {
    let other = cfb_with_stream("SomeStream", b"data");
    assert!(Document::from_reader(Cursor::new(other), DocumentFormat::Ppt).is_err());
}

#[test]
fn test_save_as_converts_doc_to_docx() {
    let doc_bytes = build_doc(&[
        Para {
            text: "Converted paragraph one.",
            terminator: '\r',
            grpprl: prose_grpprl(),
        },
        Para {
            text: "Converted paragraph two.",
            terminator: '\r',
            grpprl: prose_grpprl(),
        },
    ]);
    let doc = Document::from_reader(Cursor::new(doc_bytes), DocumentFormat::Doc).unwrap();
    let dir = Scratch::new("save-doc");
    let out = dir.path("converted.docx");
    doc.save_as(&out).expect("doc -> docx");
    let back = Document::open(&out).expect("the written .docx opens");
    assert_eq!(back.format(), DocumentFormat::Docx);
    let text = back.plain_text();
    assert!(text.contains("Converted paragraph one."), "{text:?}");
    assert!(text.contains("Converted paragraph two."), "{text:?}");
}

#[test]
fn test_save_as_converts_xls_to_xlsx() {
    let doc =
        Document::from_reader(Cursor::new(build_xls("cell text")), DocumentFormat::Xls).unwrap();
    let dir = Scratch::new("save-xls");
    let out = dir.path("converted.xlsx");
    doc.save_as(&out).expect("xls -> xlsx");
    let back = Document::open(&out).expect("the written .xlsx opens");
    assert_eq!(back.format(), DocumentFormat::Xlsx);
    let text = back.plain_text();
    assert!(text.contains("cell text"), "{text:?}");
    let xlsx = back.as_xlsx().unwrap();
    assert_eq!(xlsx.workbook.sheets[0].name, "S");
}

#[test]
fn test_save_as_converts_ppt_to_pptx() {
    let bytes = ppt_builder::build_ppt(&["Deck slide one", "Deck slide two"], &[]);
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Ppt).unwrap();
    let dir = Scratch::new("save-ppt");
    let out = dir.path("converted.pptx");
    doc.save_as(&out).expect("ppt -> pptx");
    let back = Document::open(&out).expect("the written .pptx opens");
    assert_eq!(back.format(), DocumentFormat::Pptx);
    let text = back.plain_text();
    assert!(text.contains("Deck slide one"), "{text:?}");
    assert!(text.contains("Deck slide two"), "{text:?}");
    assert_eq!(back.as_pptx().unwrap().slides.len(), 2);
}

/// One BIFF12 record (MS-XLSB §2.1.4): id in one or two 7-bit bytes, size
/// in up to four 7-bit bytes.
fn xlsb_rec(id: u32, body: &[u8]) -> Vec<u8> {
    let mut v = Vec::new();
    if id < 0x80 {
        v.push(id as u8);
    } else {
        v.push((id & 0x7F) as u8 | 0x80);
        v.push((id >> 7) as u8);
    }
    let mut size = body.len();
    loop {
        let b = (size & 0x7F) as u8;
        size >>= 7;
        if size == 0 {
            v.push(b);
            break;
        }
        v.push(b | 0x80);
    }
    v.extend_from_slice(body);
    v
}

/// An `XLWideString` (MS-XLSB §2.5.172): u32 character count, UTF-16LE.
fn xl_wide(s: &str) -> Vec<u8> {
    let mut v = (s.encode_utf16().count() as u32).to_le_bytes().to_vec();
    v.extend(s.encode_utf16().flat_map(u16::to_le_bytes));
    v
}

#[test]
fn test_xlsb_package_reads_through_the_unified_api() {
    // BrtBundleSh (156): hsState, iTabID, strRelID, strName.
    let mut sh = 0u32.to_le_bytes().to_vec();
    sh.extend_from_slice(&1u32.to_le_bytes());
    sh.extend(xl_wide("rId1"));
    sh.extend(xl_wide("Binary"));
    let workbook = xlsb_rec(156, &sh);
    // BrtSSTItem (19): flags byte, then the string.
    let mut sst_item = vec![0u8];
    sst_item.extend(xl_wide("shared text"));
    let sst = xlsb_rec(19, &sst_item);
    // BrtRowHdr (0) for row 0, then BrtCellIsst (7) at A1 -> sst[0] and
    // BrtCellReal (5) at B1 = 42.5.
    let mut sheet = xlsb_rec(0, &[0u8; 17]);
    let cell = |col: u32, value: &[u8]| {
        let mut v = col.to_le_bytes().to_vec();
        v.extend_from_slice(&0u32.to_le_bytes());
        v.extend_from_slice(value);
        v
    };
    sheet.extend(xlsb_rec(7, &cell(0, &0u32.to_le_bytes())));
    sheet.extend(xlsb_rec(5, &cell(1, &42.5f64.to_le_bytes())));

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = zip::write::SimpleFileOptions::default();
    let mut add = |name: &str, data: &[u8]| {
        zip.start_file(name, opts).unwrap();
        zip.write_all(data).unwrap();
    };
    add(
        "[Content_Types].xml",
        br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>"#,
    );
    add(
        "_rels/.rels",
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/></Relationships>"#,
    );
    add("xl/workbook.bin", &workbook);
    add(
        "xl/_rels/workbook.bin.rels",
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.bin"/></Relationships>"#,
    );
    add("xl/sharedStrings.bin", &sst);
    add("xl/worksheets/sheet1.bin", &sheet);
    let bytes = zip.finish().unwrap().into_inner();

    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Xlsx)
        .expect("an .xlsb package opens through the unified API");
    let xlsx = doc.as_xlsx().expect("decoded into the XLSX model");
    assert_eq!(xlsx.workbook.sheets[0].name, "Binary");
    let text = doc.plain_text();
    assert!(text.contains("shared text"), "{text:?}");
    assert!(text.contains("42.5"), "{text:?}");
    let ir = doc.to_ir();
    assert_eq!(ir.sections.len(), 1);
    assert_eq!(ir.sections[0].title.as_deref(), Some("Binary"));
}

/// A Workbook stream whose sector chain ends before its declared size was
/// read short with nothing in the IR to say so; `.doc` and `.ppt` already
/// flag it. The workbook now reports the text as truncated and names the
/// stream in `Metadata::warnings`.
#[test]
fn test_truncated_workbook_stream_marks_xls_text_truncated() {
    let mut bytes = build_xls("Short");
    let doc = Document::from_reader(Cursor::new(bytes.clone()), DocumentFormat::Xls).unwrap();
    let ir = doc.to_ir();
    assert!(!ir.metadata.text_truncated);
    assert!(ir.metadata.warnings.is_empty(), "{:?}", ir.metadata.warnings);
    // Directory entry 1 is "Workbook"; declare it 4 KiB longer than its chain.
    let size_at = 512 + 128 + 0x78;
    let size = u32::from_le_bytes(bytes[size_at..size_at + 4].try_into().unwrap());
    bytes[size_at..size_at + 4].copy_from_slice(&(size + 4096).to_le_bytes());
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Xls).unwrap();
    assert!(doc.plain_text().contains("Short"), "short data stays readable");
    let ir = doc.to_ir();
    assert!(ir.metadata.text_truncated);
    assert!(
        ir.metadata.warnings.iter().any(|w| w.contains("Workbook")),
        "{:?}",
        ir.metadata.warnings
    );
}

/// The container's own structural warnings (header sector counts that
/// disagree with the chains present) reach `Metadata::warnings` for
/// `.xls` as they do for `.doc`/`.ppt`, without marking the text truncated.
#[test]
fn test_xls_container_warnings_reach_metadata() {
    let mut bytes = build_xls("Cell");
    // [MS-CFB] §2.2 header `Number of DIFAT Sectors` (offset 0x48): claim
    // one although the DIFAT chain (0x44) is empty.
    bytes[0x48..0x4C].copy_from_slice(&1u32.to_le_bytes());
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Xls).unwrap();
    let ir = doc.to_ir();
    assert!(!ir.metadata.text_truncated);
    assert!(
        ir.metadata.warnings.iter().any(|w| w.contains("DIFAT")),
        "{:?}",
        ir.metadata.warnings
    );
}
