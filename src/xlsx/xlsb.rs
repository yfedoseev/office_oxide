//! Excel Binary Workbook (`.xlsb`, [MS-XLSB]).
//!
//! The same OPC package as `.xlsx` — the relationships are still XML —
//! but every SpreadsheetML part is a BIFF12 record stream instead of XML:
//! `xl/workbook.bin` (the sheet list, `BrtBundleSh`), `xl/sharedStrings.bin`
//! (`BrtSSTItem`) and `xl/worksheets/sheetN.bin` (`BrtRowHdr` and the
//! `BrtCell*`/`BrtFmla*` records, `BrtMergeCell`). Each record is a
//! variable-length id (one or two 7-bit bytes) and a variable-length size
//! (up to four 7-bit bytes), then the body; strings are `XLWideString`,
//! a 32-bit character count and UTF-16LE.
//!
//! Decoded into the [`XlsxDocument`] model, so every converter and
//! renderer downstream is the `.xlsx` one. Formulas are not decoded to
//! text (their cached values are), and styles are not read, which is
//! what calamine's `xlsb` reader offers as well.

use std::io::{Read, Seek};

use zip::read::ZipArchive;

use std::collections::HashMap;

use super::cell::{Cell, CellRef, CellValue};
use super::shared_strings::{SharedString, SharedStringTable};
use super::styles::{CellFormat, StyleSheet};
use super::workbook::{SheetInfo, SheetState, WorkbookInfo};
use super::worksheet::{Row, Worksheet};
use super::{Result, XlsxDocument};
use crate::core::opc::{self, ZipEntryIndex};
use crate::core::relationships::{Relationships, rel_types};

// Record ids ([MS-XLSB] §2.3, decimal as the specification lists them).
const BRT_ROW_HDR: u32 = 0;
const BRT_CELL_BLANK: u32 = 1;
const BRT_CELL_RK: u32 = 2;
const BRT_CELL_ERROR: u32 = 3;
const BRT_CELL_BOOL: u32 = 4;
const BRT_CELL_REAL: u32 = 5;
const BRT_CELL_ST: u32 = 6;
const BRT_CELL_ISST: u32 = 7;
const BRT_FMLA_STRING: u32 = 8;
const BRT_FMLA_NUM: u32 = 9;
const BRT_FMLA_BOOL: u32 = 10;
const BRT_FMLA_ERROR: u32 = 11;
const BRT_SST_ITEM: u32 = 19;
/// `BrtFmt`: a number format id and its code.
const BRT_FMT: u32 = 44;
/// `BrtXF`: one cell format; the ones between `BrtBeginCellXFs` and
/// `BrtEndCellXFs` are what a cell's style reference indexes.
const BRT_XF: u32 = 47;
const BRT_BEGIN_CELL_XFS: u32 = 617;
const BRT_END_CELL_XFS: u32 = 618;
const BRT_WB_PROP: u32 = 153;
const BRT_BUNDLE_SH: u32 = 156;
const BRT_MERGE_CELL: u32 = 176;

/// The sheet list lives in `workbook.bin`; its relationships are XML.
const WORKBOOK: &str = "xl/workbook.bin";
const WORKBOOK_RELS: &str = "xl/_rels/workbook.bin.rels";
const SHARED_STRINGS: &str = "xl/sharedStrings.bin";
const STYLES: &str = "xl/styles.bin";

/// Whether the package is an `.xlsb` (a binary workbook part, no XML one).
pub(super) fn is_xlsb<R: Read + Seek>(archive: &ZipArchive<R>, entries: &ZipEntryIndex) -> bool {
    let has = |n: &str| archive.index_for_name(n).is_some() || entries.lookup(n).is_some();
    has(WORKBOOK) && !has("xl/workbook.xml")
}

pub(super) fn from_zip<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    entries: &ZipEntryIndex,
) -> Result<XlsxDocument> {
    let wb = opc::read_zip_entry(archive, entries, WORKBOOK)?;
    let rels = match opc::read_zip_entry(archive, entries, WORKBOOK_RELS) {
        Ok(d) => Relationships::parse(&d).unwrap_or_else(|_| Relationships::empty()),
        Err(_) => Relationships::empty(),
    };
    let (sheets, date1904) = parse_workbook(&wb);
    let shared_strings = match opc::read_zip_entry(archive, entries, SHARED_STRINGS) {
        Ok(d) => parse_shared_strings(&d),
        Err(_) => SharedStringTable {
            strings: Vec::new(),
        },
    };
    let styles = opc::read_zip_entry(archive, entries, STYLES)
        .ok()
        .map(|d| parse_styles(&d));
    let mut worksheets = Vec::with_capacity(sheets.len());
    for (i, info) in sheets.iter().enumerate() {
        let path = rels
            .get_by_id(&info.rel_id)
            .map(|r| super::resolve_relative_zip_path("xl/workbook.bin", &r.target))
            .unwrap_or_else(|| format!("xl/worksheets/sheet{}.bin", i + 1));
        let Ok(data) = opc::read_zip_entry(archive, entries, &path) else {
            continue;
        };
        worksheets.push(parse_sheet(&data, info.name.clone()));
    }
    let core_properties = XlsxDocument::read_xml_entry(archive, entries, "docProps/core.xml")
        .ok()
        .and_then(|d| crate::core::properties::CoreProperties::parse(&d).ok());
    let app_properties = XlsxDocument::read_xml_entry(archive, entries, "docProps/app.xml")
        .ok()
        .and_then(|d| crate::core::properties::AppProperties::parse(&d).ok());
    let has_macros = rels.first_by_type(rel_types::VBA_PROJECT).is_some();
    Ok(XlsxDocument {
        workbook: WorkbookInfo {
            sheets,
            defined_names: Vec::new(),
            date1904,
        },
        worksheets,
        shared_strings,
        styles,
        theme: None,
        chart_text: Vec::new(),
        embedded_fonts: Vec::new(),
        core_properties,
        app_properties,
        has_macros,
        styles_data: None,
        theme_data: None,
    })
}

/// One BIFF12 record: id and body.
struct Record<'a> {
    id: u32,
    data: &'a [u8],
}

/// Iterate the records of a BIFF12 stream. A truncated header or body
/// ends the iteration; nothing is read past the stream.
struct Records<'a> {
    data: &'a [u8],
    pos: usize,
}

impl<'a> Iterator for Records<'a> {
    type Item = Record<'a>;

    fn next(&mut self) -> Option<Record<'a>> {
        let d = self.data;
        let mut p = self.pos;
        // Id: one byte, or two when bit 7 of the first is set.
        let b0 = *d.get(p)?;
        p += 1;
        let mut id = (b0 & 0x7F) as u32;
        if b0 & 0x80 != 0 {
            let b1 = *d.get(p)?;
            p += 1;
            id |= ((b1 & 0x7F) as u32) << 7;
        }
        // Size: up to four 7-bit bytes.
        let mut size = 0usize;
        for i in 0..4 {
            let b = *d.get(p)?;
            p += 1;
            size |= ((b & 0x7F) as usize) << (7 * i);
            if b & 0x80 == 0 {
                break;
            }
        }
        let end = p.checked_add(size)?;
        if end > d.len() {
            return None;
        }
        self.pos = end;
        Some(Record {
            id,
            data: &d[p..end],
        })
    }
}

fn records(data: &[u8]) -> Records<'_> {
    Records { data, pos: 0 }
}

fn u32_at(d: &[u8], o: usize) -> Option<u32> {
    d.get(o..o + 4)
        .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

/// `XLWideString` at `o`: character count then UTF-16LE. Returns the text
/// and the offset just past it; a count of `0xFFFFFFFF` is the nullable
/// form's "no string".
fn wide_string(d: &[u8], o: usize) -> Option<(String, usize)> {
    let cch = u32_at(d, o)?;
    if cch == 0xFFFF_FFFF {
        return Some((String::new(), o + 4));
    }
    let bytes = (cch as usize).checked_mul(2)?;
    let start = o + 4;
    let end = start.checked_add(bytes)?;
    let raw = d.get(start..end)?;
    let units: Vec<u16> = raw
        .as_chunks::<2>()
        .0
        .iter()
        .map(|c| u16::from_le_bytes(*c))
        .collect();
    Some((String::from_utf16_lossy(&units), end))
}

/// The sheet list and the 1904 flag from `workbook.bin`.
fn parse_workbook(data: &[u8]) -> (Vec<SheetInfo>, bool) {
    let mut sheets = Vec::new();
    let mut date1904 = false;
    for rec in records(data) {
        match rec.id {
            BRT_WB_PROP => {
                // BrtWbProp flags: bit 0 is f1904.
                if let Some(flags) = u32_at(rec.data, 0) {
                    date1904 = flags & 1 != 0;
                }
            },
            BRT_BUNDLE_SH => {
                // hsState, iTabID, strRelID (nullable), strName. Excel 2007
                // Beta 2 wrote one more dword before the strings; a record
                // whose strings do not parse at the released offset is
                // read at that one.
                let Some(hs_state) = u32_at(rec.data, 0) else {
                    continue;
                };
                let Some(sheet_id) = u32_at(rec.data, 4) else {
                    continue;
                };
                let strings_at = |o: usize| {
                    let (rel_id, next) = wide_string(rec.data, o)?;
                    let (name, _) = wide_string(rec.data, next)?;
                    (!rel_id.chars().any(char::is_control) && !name.is_empty())
                        .then_some((rel_id, name))
                };
                let Some((rel_id, name)) = strings_at(8).or_else(|| strings_at(12)) else {
                    continue;
                };
                sheets.push(SheetInfo {
                    name,
                    sheet_id,
                    rel_id,
                    state: match hs_state {
                        1 => SheetState::Hidden,
                        2 => SheetState::VeryHidden,
                        _ => SheetState::Visible,
                    },
                });
            },
            _ => {},
        }
    }
    (sheets, date1904)
}

/// `sharedStrings.bin`: one `BrtSSTItem` per entry — a flags byte, the
/// text, then the rich runs and phonetic data the flags announce, which
/// this reader does not need.
fn parse_shared_strings(data: &[u8]) -> SharedStringTable {
    let mut strings = Vec::new();
    for rec in records(data) {
        if rec.id == BRT_SST_ITEM {
            if let Some((text, _)) = wide_string(rec.data, 1) {
                strings.push(SharedString {
                    text,
                    rich_text: None,
                });
            }
        }
    }
    SharedStringTable { strings }
}

/// `styles.bin`: the number-format codes and the `cellXfs` array, which
/// is what turns a serial into a date downstream — the same fields the
/// `.xlsx` stylesheet supplies; fonts, fills and borders are not read.
fn parse_styles(data: &[u8]) -> StyleSheet {
    let mut number_formats = HashMap::new();
    let mut cell_formats = Vec::new();
    let mut in_cell_xfs = false;
    for rec in records(data) {
        let d = rec.data;
        match rec.id {
            BRT_FMT => {
                // ifmt (u16), stFmtCode.
                if let (Some(b), Some((code, _))) = (d.get(0..2), wide_string(d, 2)) {
                    number_formats.insert(u16::from_le_bytes([b[0], b[1]]) as u32, code);
                }
            },
            BRT_BEGIN_CELL_XFS => in_cell_xfs = true,
            BRT_END_CELL_XFS => in_cell_xfs = false,
            BRT_XF if in_cell_xfs => {
                // ixfeParent, iFmt, iFont, iFill, ixBorder (u16 each).
                if let Some(b) = d.get(0..10) {
                    cell_formats.push(CellFormat {
                        number_format_id: u16::from_le_bytes([b[2], b[3]]) as u32,
                        font_index: Some(u16::from_le_bytes([b[4], b[5]]) as u32),
                        fill_index: Some(u16::from_le_bytes([b[6], b[7]]) as u32),
                        border_index: Some(u16::from_le_bytes([b[8], b[9]]) as u32),
                        apply_number_format: true,
                        xf_id: Some(u16::from_le_bytes([b[0], b[1]]) as u32),
                    });
                }
            },
            _ => {},
        }
    }
    StyleSheet {
        number_formats,
        fonts: Vec::new(),
        fills: Vec::new(),
        borders: Vec::new(),
        cell_formats,
        cell_style_formats: Vec::new(),
    }
}

/// The `Cell` structure every cell record starts with: the column, then
/// the style reference (24 bits) and flags.
fn cell_head(d: &[u8]) -> Option<(u32, u32)> {
    Some((u32_at(d, 0)?, u32_at(d, 4)? & 0x00FF_FFFF))
}

fn error_text(code: u8) -> String {
    match code {
        0x00 => "#NULL!",
        0x07 => "#DIV/0!",
        0x0F => "#VALUE!",
        0x17 => "#REF!",
        0x1D => "#NAME?",
        0x24 => "#NUM!",
        0x2A => "#N/A",
        0x2B => "#GETTING_DATA",
        _ => "#ERR",
    }
    .to_string()
}

fn parse_sheet(data: &[u8], name: String) -> Worksheet {
    let mut rows: Vec<Row> = Vec::new();
    let mut merged_cells = Vec::new();
    let mut current: Option<Row> = None;
    let mut row_index: u32 = 0;
    let push = |rows: &mut Vec<Row>,
                current: &mut Option<Row>,
                row: u32,
                col: u32,
                style: u32,
                value: CellValue| {
        // Past XFD1048576 is not a cell (a corrupted index would size
        // the table's rows to it downstream).
        if col > super::cell::MAX_COL || row >= super::cell::MAX_ROWS {
            return;
        }
        if current.as_ref().is_some_and(|r| r.index != row + 1) {
            if let Some(r) = current.take() {
                rows.push(r);
            }
        }
        let r = current.get_or_insert_with(|| Row {
            index: row + 1,
            cells: Vec::new(),
        });
        r.cells.push(Cell {
            reference: CellRef { col, row },
            value,
            style_index: Some(style),
            formula: None,
            rich_runs: None,
            vm: None,
        });
    };
    for rec in records(data) {
        let d = rec.data;
        match rec.id {
            BRT_ROW_HDR => {
                if let Some(rw) = u32_at(d, 0) {
                    row_index = rw;
                }
            },
            BRT_CELL_BLANK => {},
            BRT_CELL_RK => {
                if let (Some((col, style)), Some(rk)) = (cell_head(d), u32_at(d, 8)) {
                    let n = crate::xls::decode_rk(rk);
                    push(&mut rows, &mut current, row_index, col, style, CellValue::Number(n));
                }
            },
            BRT_CELL_ERROR | BRT_FMLA_ERROR => {
                if let (Some((col, style)), Some(&code)) = (cell_head(d), d.get(8)) {
                    push(
                        &mut rows,
                        &mut current,
                        row_index,
                        col,
                        style,
                        CellValue::Error(error_text(code)),
                    );
                }
            },
            BRT_CELL_BOOL | BRT_FMLA_BOOL => {
                if let (Some((col, style)), Some(&b)) = (cell_head(d), d.get(8)) {
                    push(
                        &mut rows,
                        &mut current,
                        row_index,
                        col,
                        style,
                        CellValue::Boolean(b != 0),
                    );
                }
            },
            BRT_CELL_REAL | BRT_FMLA_NUM => {
                if let (Some((col, style)), Some(b)) = (cell_head(d), d.get(8..16)) {
                    let n = f64::from_le_bytes([b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7]]);
                    push(&mut rows, &mut current, row_index, col, style, CellValue::Number(n));
                }
            },
            BRT_CELL_ST | BRT_FMLA_STRING => {
                if let (Some((col, style)), Some((text, _))) = (cell_head(d), wide_string(d, 8)) {
                    push(&mut rows, &mut current, row_index, col, style, CellValue::String(text));
                }
            },
            BRT_CELL_ISST => {
                if let (Some((col, style)), Some(isst)) = (cell_head(d), u32_at(d, 8)) {
                    push(
                        &mut rows,
                        &mut current,
                        row_index,
                        col,
                        style,
                        CellValue::SharedString(isst),
                    );
                }
            },
            BRT_MERGE_CELL => {
                if let (Some(r1), Some(r2), Some(c1), Some(c2)) =
                    (u32_at(d, 0), u32_at(d, 4), u32_at(d, 8), u32_at(d, 12))
                {
                    merged_cells.push(format!(
                        "{}{}:{}{}",
                        CellRef::col_name(c1),
                        r1 + 1,
                        CellRef::col_name(c2),
                        r2 + 1
                    ));
                }
            },
            _ => {},
        }
    }
    if let Some(r) = current.take() {
        rows.push(r);
    }
    Worksheet {
        name,
        dimension: None,
        rows,
        merged_cells,
        hyperlinks: Vec::new(),
        page_setup: None,
        images: Vec::new(),
        comments: Vec::new(),
        text_shapes: Vec::new(),
        conditional_formats: Vec::new(),
        data_validations: Vec::new(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A record: id in the one- or two-byte form, size in 7-bit bytes.
    fn rec(id: u32, body: &[u8]) -> Vec<u8> {
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

    fn wide(s: &str) -> Vec<u8> {
        let mut v = (s.encode_utf16().count() as u32).to_le_bytes().to_vec();
        v.extend(s.encode_utf16().flat_map(u16::to_le_bytes));
        v
    }

    fn cell(col: u32, style: u32, value: &[u8]) -> Vec<u8> {
        let mut v = col.to_le_bytes().to_vec();
        v.extend_from_slice(&style.to_le_bytes());
        v.extend_from_slice(value);
        v
    }

    #[test]
    fn test_records_decode_one_and_two_byte_ids_and_multibyte_sizes() {
        let mut s = rec(5, &[1, 2, 3]);
        s.extend(rec(156, &vec![7u8; 300]));
        let got: Vec<(u32, usize)> = records(&s).map(|r| (r.id, r.data.len())).collect();
        assert_eq!(got, vec![(5, 3), (156, 300)]);
        // A truncated body ends the stream.
        let mut t = rec(5, &[1, 2, 3]);
        t.truncate(t.len() - 1);
        assert_eq!(records(&t).count(), 0);
    }

    #[test]
    fn test_workbook_sheet_list_and_1904_flag() {
        let mut wb = rec(BRT_WB_PROP, &1u32.to_le_bytes());
        let mut sh = 1u32.to_le_bytes().to_vec(); // hidden
        sh.extend_from_slice(&7u32.to_le_bytes());
        sh.extend(wide("rId3"));
        sh.extend(wide("Data"));
        wb.extend(rec(BRT_BUNDLE_SH, &sh));
        let (sheets, date1904) = parse_workbook(&wb);
        assert!(date1904);
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "Data");
        assert_eq!(sheets[0].rel_id, "rId3");
        assert_eq!(sheets[0].sheet_id, 7);
        assert_eq!(sheets[0].state, SheetState::Hidden);
    }

    /// Excel 2007 Beta 2 wrote one more dword before the sheet's strings.
    #[test]
    fn test_beta_bundle_sheet_layout_is_read_at_the_second_offset() {
        let mut sh = 0u32.to_le_bytes().to_vec();
        sh.extend_from_slice(&0u32.to_le_bytes());
        sh.extend_from_slice(&1u32.to_le_bytes());
        sh.extend(wide("rId1"));
        sh.extend(wide("Sheet1"));
        let wb = rec(BRT_BUNDLE_SH, &sh);
        let (sheets, _) = parse_workbook(&wb);
        assert_eq!(sheets.len(), 1);
        assert_eq!((sheets[0].rel_id.as_str(), sheets[0].name.as_str()), ("rId1", "Sheet1"));
    }

    #[test]
    fn test_sheet_cells_of_every_kind_land_in_their_rows() {
        let mut s = rec(BRT_ROW_HDR, &[0u8; 17]);
        s.extend(rec(BRT_CELL_ISST, &cell(0, 3, &2u32.to_le_bytes())));
        s.extend(rec(BRT_CELL_RK, &cell(1, 0, &(100u32 << 2 | 2).to_le_bytes())));
        s.extend(rec(BRT_CELL_REAL, &cell(2, 0, &2.5f64.to_le_bytes())));
        let mut row2 = 4u32.to_le_bytes().to_vec();
        row2.extend_from_slice(&[0u8; 13]);
        s.extend(rec(BRT_ROW_HDR, &row2));
        s.extend(rec(BRT_CELL_ST, &cell(0, 0, &wide("inline"))));
        s.extend(rec(BRT_CELL_BOOL, &cell(1, 0, &[1])));
        s.extend(rec(BRT_CELL_ERROR, &cell(2, 0, &[0x07])));
        let mut f = cell(3, 0, &wide("cached"));
        f.extend_from_slice(&[0, 0, 0, 0, 0, 0]);
        s.extend(rec(BRT_FMLA_STRING, &f));
        let mut m = 0u32.to_le_bytes().to_vec();
        m.extend_from_slice(&4u32.to_le_bytes());
        m.extend_from_slice(&0u32.to_le_bytes());
        m.extend_from_slice(&1u32.to_le_bytes());
        s.extend(rec(BRT_MERGE_CELL, &m));
        let ws = parse_sheet(&s, "S".into());
        assert_eq!(ws.rows.len(), 2);
        assert_eq!(ws.rows[0].index, 1);
        assert_eq!(ws.rows[1].index, 5);
        let r0 = &ws.rows[0].cells;
        assert!(matches!(r0[0].value, CellValue::SharedString(2)));
        assert_eq!(r0[0].style_index, Some(3));
        assert!(matches!(r0[1].value, CellValue::Number(n) if n == 100.0));
        assert!(matches!(r0[2].value, CellValue::Number(n) if n == 2.5));
        let r1 = &ws.rows[1].cells;
        assert!(matches!(&r1[0].value, CellValue::String(s) if s == "inline"));
        assert!(matches!(r1[1].value, CellValue::Boolean(true)));
        assert!(matches!(&r1[2].value, CellValue::Error(e) if e == "#DIV/0!"));
        assert!(matches!(&r1[3].value, CellValue::String(s) if s == "cached"));
        assert_eq!(ws.merged_cells, vec!["A1:B5"]);
    }

    #[test]
    fn test_styles_map_cell_formats_to_number_format_codes() {
        let mut st = Vec::new();
        let mut fmt = 164u16.to_le_bytes().to_vec();
        fmt.extend(wide("yyyy-mm-dd"));
        st.extend(rec(BRT_FMT, &fmt));
        st.extend(rec(BRT_BEGIN_CELL_XFS, &1u32.to_le_bytes()));
        let mut xf = 0u16.to_le_bytes().to_vec(); // ixfeParent
        xf.extend_from_slice(&164u16.to_le_bytes()); // iFmt
        xf.extend_from_slice(&[0u8; 12]);
        st.extend(rec(BRT_XF, &xf));
        st.extend(rec(BRT_END_CELL_XFS, &[]));
        let sheet = parse_styles(&st);
        assert_eq!(sheet.number_format_for(0), Some("yyyy-mm-dd"));
    }

    #[test]
    fn test_shared_strings_take_the_text_and_skip_rich_runs() {
        let mut sst = Vec::new();
        let mut plain = vec![0u8];
        plain.extend(wide("plain"));
        sst.extend(rec(BRT_SST_ITEM, &plain));
        let mut rich = vec![1u8];
        rich.extend(wide("rich"));
        rich.extend_from_slice(&1u32.to_le_bytes());
        rich.extend_from_slice(&[0u8; 4]);
        sst.extend(rec(BRT_SST_ITEM, &rich));
        let table = parse_shared_strings(&sst);
        let texts: Vec<&str> = table.strings.iter().map(|s| s.text.as_str()).collect();
        assert_eq!(texts, ["plain", "rich"]);
    }
}
