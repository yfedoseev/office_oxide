//! Excel 2.x–4.0 (BIFF2, BIFF3, BIFF4) worksheet streams.
//!
//! A pre-OLE2 Excel file has no compound-file container: the file *is*
//! one BIFF record stream, starting with its own `BOF` (`0x0009`,
//! `0x0209` or `0x0409` by version) and holding one worksheet. The cell
//! records are the ancestors of BIFF8's: from BIFF3 on `NUMBER`, `RK`,
//! `BLANK` and `BOOLERR` already have their final layout, `LABEL` is a
//! 16-bit length and raw 8-bit characters (no option-flags byte), and
//! `FORMULA` carries its cached result the way every later version does,
//! with a string result delivered by the `STRING` record that follows.
//! BIFF2 cells carry a 3-byte cell-attribute field where later versions
//! carry a 16-bit XF index, and their own record ids.
//!
//! This is the record set Apache POI's `OldExcelExtractor` and xlrd's
//! `biffh` read for these files. Number formats are not applied — the
//! XF and FORMAT records of these versions are different again — so a
//! date renders as its serial, as xlrd shows it.

use std::sync::Arc;

use super::cell::{Cell, CellValue, decode_rk};
use super::error::{Result, XlsError};
use super::records::RecordIter;
use super::workbook::{NumberFormats, Sheet, XlsDocument, build_grid};

// Record ids by version ([MS-XLS] record enumeration; the BIFF2 ids are
// the original 1987 numbering, BIFF3 added 0x0200, BIFF4 0x0400).
const BIFF2_BLANK: u16 = 0x0001;
const BIFF2_INTEGER: u16 = 0x0002;
const BIFF2_NUMBER: u16 = 0x0003;
const BIFF2_LABEL: u16 = 0x0004;
const BIFF2_BOOLERR: u16 = 0x0005;
const BIFF2_FORMULA: u16 = 0x0006;
const BIFF2_STRING: u16 = 0x0007;
const BLANK: u16 = 0x0201;
const NUMBER: u16 = 0x0203;
const LABEL: u16 = 0x0204;
const BOOLERR: u16 = 0x0205;
const BIFF3_FORMULA: u16 = 0x0206;
const STRING: u16 = 0x0207;
const BIFF4_FORMULA: u16 = 0x0406;
const RK: u16 = 0x027E;
/// `CODEPAGE` ([MS-XLS] §2.4.52): the 8-bit character set of every string.
const CODEPAGE: u16 = 0x0042;
const EOF: u16 = 0x000A;

/// Parse a raw BIFF2/3/4 stream into a one-sheet workbook.
pub(super) fn parse(data: &[u8], biff: u8) -> Result<XlsDocument> {
    let mut cells: Vec<Cell> = Vec::new();
    let mut codepage: Option<u16> = None;
    // A `FORMULA` whose cached result is a string leaves its cell here
    // until the `STRING` record that follows supplies the text.
    let mut pending_string: Option<(u16, u16, u16)> = None;
    let mut records = 0u32;
    for rec in RecordIter::new(data) {
        let rec = rec?;
        records += 1;
        if records > 4_000_000 {
            return Err(XlsError::InvalidRecord("record budget exhausted".into()));
        }
        let d: &[u8] = &rec.data;
        match (biff, rec.record_type) {
            (_, CODEPAGE) if d.len() >= 2 => {
                codepage = Some(u16::from_le_bytes([d[0], d[1]]));
            },
            (_, EOF) => break,
            // ---- BIFF2: rw, col, 3-byte cell attributes, then the value.
            (2, BIFF2_BLANK) => {},
            (2, BIFF2_INTEGER) if d.len() >= 9 => {
                cells.push(cell2(d, CellValue::Number(u16::from_le_bytes([d[7], d[8]]) as f64)));
            },
            (2, BIFF2_NUMBER) if d.len() >= 15 => {
                cells.push(cell2(d, CellValue::Number(f64_at(d, 7))));
            },
            (2, BIFF2_LABEL) if d.len() >= 8 => {
                let n = d[7] as usize;
                let end = (8 + n).min(d.len());
                let s = super::codepage::decode_biff5_text(&d[8..end], codepage);
                cells.push(cell2(d, CellValue::String(s)));
            },
            (2, BIFF2_BOOLERR) if d.len() >= 9 => {
                let value = if d[8] != 0 {
                    CellValue::Error(d[7])
                } else {
                    CellValue::Bool(d[7] != 0)
                };
                cells.push(cell2(d, value));
            },
            (2, BIFF2_FORMULA) if d.len() >= 15 => {
                let (row, col) =
                    (u16::from_le_bytes([d[0], d[1]]), u16::from_le_bytes([d[2], d[3]]));
                match cached_result(&d[7..15]) {
                    Cached::String => pending_string = Some((row, col, 0)),
                    Cached::Value(v) => cells.push(Cell {
                        row,
                        col,
                        xf_index: 0,
                        value: v,
                    }),
                }
            },
            (2, BIFF2_STRING) if !d.is_empty() => {
                if let Some((row, col, xf_index)) = pending_string.take() {
                    let n = d[0] as usize;
                    let end = (1 + n).min(d.len());
                    let s = super::codepage::decode_biff5_text(&d[1..end], codepage);
                    cells.push(Cell {
                        row,
                        col,
                        xf_index,
                        value: CellValue::String(s),
                    });
                }
            },
            // ---- BIFF3/4: rw, col, ixfe, then the value.
            (3 | 4, BLANK) => {},
            (3 | 4, NUMBER) if d.len() >= 14 => {
                cells.push(cell3(d, CellValue::Number(f64_at(d, 6))));
            },
            (3 | 4, RK) if d.len() >= 10 => {
                let rk = u32::from_le_bytes([d[6], d[7], d[8], d[9]]);
                cells.push(cell3(d, CellValue::Number(decode_rk(rk))));
            },
            (3 | 4, LABEL) if d.len() >= 8 => {
                let n = u16::from_le_bytes([d[6], d[7]]) as usize;
                let end = (8 + n).min(d.len());
                let s = super::codepage::decode_biff5_text(&d[8..end], codepage);
                cells.push(cell3(d, CellValue::String(s)));
            },
            (3 | 4, BOOLERR) if d.len() >= 8 => {
                let value = if d[7] != 0 {
                    CellValue::Error(d[6])
                } else {
                    CellValue::Bool(d[6] != 0)
                };
                cells.push(cell3(d, value));
            },
            (3, BIFF3_FORMULA) | (4, BIFF4_FORMULA) if d.len() >= 14 => {
                let (row, col, xf_index) = header3(d);
                match cached_result(&d[6..14]) {
                    Cached::String => pending_string = Some((row, col, xf_index)),
                    Cached::Value(v) => cells.push(Cell {
                        row,
                        col,
                        xf_index,
                        value: v,
                    }),
                }
            },
            (3 | 4, STRING) if d.len() >= 2 => {
                if let Some((row, col, xf_index)) = pending_string.take() {
                    let n = u16::from_le_bytes([d[0], d[1]]) as usize;
                    let end = (2 + n).min(d.len());
                    let s = super::codepage::decode_biff5_text(&d[2..end], codepage);
                    cells.push(Cell {
                        row,
                        col,
                        xf_index,
                        value: CellValue::String(s),
                    });
                }
            },
            _ => {},
        }
    }
    let (rows, xf) = build_grid(&mut cells);
    Ok(XlsDocument::from_one_sheet(Sheet {
        name: "Sheet1".to_string(),
        rows,
        xf,
        formats: Arc::new(NumberFormats::default()),
        ..Default::default()
    }))
}

fn f64_at(d: &[u8], o: usize) -> f64 {
    f64::from_le_bytes([
        d[o],
        d[o + 1],
        d[o + 2],
        d[o + 3],
        d[o + 4],
        d[o + 5],
        d[o + 6],
        d[o + 7],
    ])
}

fn cell2(d: &[u8], value: CellValue) -> Cell {
    Cell {
        row: u16::from_le_bytes([d[0], d[1]]),
        col: u16::from_le_bytes([d[2], d[3]]),
        xf_index: 0,
        value,
    }
}

fn header3(d: &[u8]) -> (u16, u16, u16) {
    (
        u16::from_le_bytes([d[0], d[1]]),
        u16::from_le_bytes([d[2], d[3]]),
        u16::from_le_bytes([d[4], d[5]]),
    )
}

fn cell3(d: &[u8], value: CellValue) -> Cell {
    let (row, col, xf_index) = header3(d);
    Cell {
        row,
        col,
        xf_index,
        value,
    }
}

enum Cached {
    String,
    Value(CellValue),
}

/// A formula's 8-byte cached result: a number, unless the last two bytes
/// are `FF FF`, in which case the first byte says string (the text is in
/// the following `STRING` record), boolean, error or empty string.
fn cached_result(r: &[u8]) -> Cached {
    if r[6] == 0xFF && r[7] == 0xFF {
        match r[0] {
            0 => Cached::String,
            1 => Cached::Value(CellValue::Bool(r[2] != 0)),
            2 => Cached::Value(CellValue::Error(r[2])),
            _ => Cached::Value(CellValue::String(String::new())),
        }
    } else {
        Cached::Value(CellValue::Number(f64_at(r, 0)))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn rec(rt: u16, body: &[u8]) -> Vec<u8> {
        let mut v = rt.to_le_bytes().to_vec();
        v.extend_from_slice(&(body.len() as u16).to_le_bytes());
        v.extend_from_slice(body);
        v
    }

    fn biff3_stream() -> Vec<u8> {
        let mut s = rec(0x0209, &[0, 0, 0x10, 0, 0, 0]);
        s.extend(rec(CODEPAGE, &0x8001u16.to_le_bytes()));
        // A1 "Season" — LABEL: rw, col, ixfe, cch u16, chars.
        let mut l = vec![0, 0, 0, 0, 0x31, 0];
        l.extend_from_slice(&6u16.to_le_bytes());
        l.extend_from_slice(b"Season");
        s.extend(rec(LABEL, &l));
        // B1 376.2 — NUMBER.
        let mut n = vec![0, 0, 1, 0, 0x39, 0];
        n.extend_from_slice(&376.2f64.to_le_bytes());
        s.extend(rec(NUMBER, &n));
        // C1 100 — RK (integer 100 << 2).
        let mut r = vec![0, 0, 2, 0, 0x39, 0];
        r.extend_from_slice(&(100u32 << 2 | 2).to_le_bytes());
        s.extend(rec(RK, &r));
        // A2 formula with a string result, then the STRING.
        let mut f = vec![1, 0, 0, 0, 0x39, 0];
        f.extend_from_slice(&[0, 0, 0, 0, 0, 0, 0xFF, 0xFF]);
        f.extend_from_slice(&[0, 0, 0, 0]);
        s.extend(rec(BIFF3_FORMULA, &f));
        let mut st = 5u16.to_le_bytes().to_vec();
        st.extend_from_slice(b"total");
        s.extend(rec(STRING, &st));
        // B2 TRUE — BOOLERR.
        s.extend(rec(BOOLERR, &[1, 0, 1, 0, 0x39, 0, 1, 0]));
        s.extend(rec(EOF, &[]));
        s
    }

    #[test]
    fn test_biff3_cells_are_read() {
        let doc = parse(&biff3_stream(), 3).unwrap();
        let sheet = &doc.sheets[0];
        assert_eq!(sheet.display_text(0, 0).unwrap(), "Season");
        assert_eq!(sheet.display_text(0, 1).unwrap(), "376.2");
        assert_eq!(sheet.display_text(0, 2).unwrap(), "100");
        assert_eq!(sheet.display_text(1, 0).unwrap(), "total");
        assert_eq!(sheet.display_text(1, 1).unwrap(), "TRUE");
    }

    #[test]
    fn test_biff2_cells_are_read() {
        let mut s = rec(0x0009, &[0, 0, 0x10, 0]);
        // INTEGER: rw, col, 3 attribute bytes, u16.
        let mut i = vec![0, 0, 0, 0, 0, 0, 0];
        i.extend_from_slice(&42u16.to_le_bytes());
        s.extend(rec(BIFF2_INTEGER, &i));
        // LABEL: rw, col, attrs, cch u8, chars.
        let mut l = vec![0, 0, 1, 0, 0, 0, 0, 5];
        l.extend_from_slice(b"hello");
        s.extend(rec(BIFF2_LABEL, &l));
        let mut n = vec![1, 0, 0, 0, 0, 0, 0];
        n.extend_from_slice(&2.5f64.to_le_bytes());
        s.extend(rec(BIFF2_NUMBER, &n));
        s.extend(rec(EOF, &[]));
        let doc = parse(&s, 2).unwrap();
        let sheet = &doc.sheets[0];
        assert_eq!(sheet.display_text(0, 0).unwrap(), "42");
        assert_eq!(sheet.display_text(0, 1).unwrap(), "hello");
        assert_eq!(sheet.display_text(1, 0).unwrap(), "2.5");
    }

    /// Truncated and garbage records never index out of bounds.
    #[test]
    fn test_short_records_are_skipped() {
        let mut s = rec(0x0209, &[0, 0]);
        s.extend(rec(LABEL, &[0, 0, 0, 0, 0]));
        s.extend(rec(NUMBER, &[0, 0, 0, 0, 0, 0, 1]));
        s.extend(rec(BIFF3_FORMULA, &[0; 13]));
        s.extend(rec(STRING, &[9, 0, b'x']));
        let doc = parse(&s, 3).unwrap();
        assert!(doc.sheets[0].rows.is_empty());
    }
}
