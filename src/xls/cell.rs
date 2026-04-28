//! Cell value parsing for BIFF8 records.

use super::error::{Result, XlsError};
use super::records::*;
use super::sst::read_unicode_string;

/// A cell value in an XLS spreadsheet.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum CellValue {
    /// Cell contains no value.
    #[default]
    Empty,
    /// Floating-point number.
    Number(f64),
    /// Text string (from SST or inline label).
    String(String),
    /// Boolean value.
    Bool(bool),
    /// Error code byte (e.g. `0x07` = `#DIV/0!`).
    Error(u8),
}

impl CellValue {
    /// Get the display text of a cell value.
    pub fn as_text(&self) -> String {
        match self {
            Self::Empty => String::new(),
            Self::Number(n) => {
                if *n == (*n as i64) as f64 && n.abs() < 1e15 {
                    format!("{}", *n as i64)
                } else {
                    format!("{n}")
                }
            },
            Self::String(s) => s.clone(),
            Self::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            Self::Error(code) => match code {
                0x00 => "#NULL!".into(),
                0x07 => "#DIV/0!".into(),
                0x0F => "#VALUE!".into(),
                0x17 => "#REF!".into(),
                0x1D => "#NAME?".into(),
                0x24 => "#NUM!".into(),
                0x2A => "#N/A".into(),
                _ => format!("#ERR({code})"),
            },
        }
    }
}

/// A cell with its position and value.
#[derive(Debug, Clone)]
pub struct Cell {
    /// 0-based row index.
    pub row: u16,
    /// 0-based column index.
    pub col: u16,
    /// The parsed cell value.
    pub value: CellValue,
}

/// Parse cells from a BIFF record.
///
/// Returns a list of cells extracted from a single record.
pub fn parse_cell_record(record: &BiffRecord, sst: &[String]) -> Result<Vec<Cell>> {
    match record.record_type {
        RT_LABELSST => parse_labelsst(&record.data, sst),
        RT_NUMBER => parse_number(&record.data),
        RT_RK => parse_rk_record(&record.data),
        RT_MULRK => parse_mulrk(&record.data),
        RT_BOOLERR => parse_boolerr(&record.data),
        RT_LABEL | RT_RSTRING => parse_label(&record.data),
        RT_BLANK => parse_blank(&record.data),
        RT_MULBLANK => parse_mulblank(&record.data),
        RT_FORMULA => parse_formula(&record.data),
        _ => Ok(Vec::new()),
    }
}

fn parse_labelsst(data: &[u8], sst: &[String]) -> Result<Vec<Cell>> {
    if data.len() < 10 {
        return Err(XlsError::InvalidRecord("LABELSST too short".into()));
    }
    let row = u16::from_le_bytes([data[0], data[1]]);
    let col = u16::from_le_bytes([data[2], data[3]]);
    let sst_index = u32::from_le_bytes([data[6], data[7], data[8], data[9]]) as usize;

    let value = if sst_index < sst.len() {
        CellValue::String(sst[sst_index].clone())
    } else {
        CellValue::String(String::new())
    };

    Ok(vec![Cell { row, col, value }])
}

fn parse_number(data: &[u8]) -> Result<Vec<Cell>> {
    if data.len() < 14 {
        return Err(XlsError::InvalidRecord("NUMBER too short".into()));
    }
    let row = u16::from_le_bytes([data[0], data[1]]);
    let col = u16::from_le_bytes([data[2], data[3]]);
    let value = f64::from_le_bytes([
        data[6], data[7], data[8], data[9], data[10], data[11], data[12], data[13],
    ]);
    Ok(vec![Cell {
        row,
        col,
        value: CellValue::Number(value),
    }])
}

fn parse_rk_record(data: &[u8]) -> Result<Vec<Cell>> {
    if data.len() < 10 {
        return Err(XlsError::InvalidRecord("RK too short".into()));
    }
    let row = u16::from_le_bytes([data[0], data[1]]);
    let col = u16::from_le_bytes([data[2], data[3]]);
    let rk_val = u32::from_le_bytes([data[6], data[7], data[8], data[9]]);
    let value = decode_rk(rk_val);
    Ok(vec![Cell {
        row,
        col,
        value: CellValue::Number(value),
    }])
}

fn parse_mulrk(data: &[u8]) -> Result<Vec<Cell>> {
    if data.len() < 6 {
        return Err(XlsError::InvalidRecord("MULRK too short".into()));
    }
    let row = u16::from_le_bytes([data[0], data[1]]);
    let first_col = u16::from_le_bytes([data[2], data[3]]);
    // Last 2 bytes = last_col.
    // Each RK entry: 2 bytes XF index + 4 bytes RK value = 6 bytes.
    let rk_data = &data[4..data.len() - 2];
    let count = rk_data.len() / 6;

    let mut cells = Vec::with_capacity(count);
    for i in 0..count {
        let off = i * 6;
        let rk_val = u32::from_le_bytes([
            rk_data[off + 2],
            rk_data[off + 3],
            rk_data[off + 4],
            rk_data[off + 5],
        ]);
        cells.push(Cell {
            row,
            col: first_col + i as u16,
            value: CellValue::Number(decode_rk(rk_val)),
        });
    }
    Ok(cells)
}

fn parse_boolerr(data: &[u8]) -> Result<Vec<Cell>> {
    if data.len() < 8 {
        return Err(XlsError::InvalidRecord("BOOLERR too short".into()));
    }
    let row = u16::from_le_bytes([data[0], data[1]]);
    let col = u16::from_le_bytes([data[2], data[3]]);
    let val = data[6];
    let is_error = data[7];
    let value = if is_error != 0 {
        CellValue::Error(val)
    } else {
        CellValue::Bool(val != 0)
    };
    Ok(vec![Cell { row, col, value }])
}

fn parse_label(data: &[u8]) -> Result<Vec<Cell>> {
    if data.len() < 8 {
        return Err(XlsError::InvalidRecord("LABEL too short".into()));
    }
    let row = u16::from_le_bytes([data[0], data[1]]);
    let col = u16::from_le_bytes([data[2], data[3]]);
    // Try BIFF8 unicode string first; fall back to raw bytes for BIFF5.
    let s = match read_unicode_string(data, 6) {
        Ok((s, end)) if end <= data.len() + 4 => s,
        _ => {
            // BIFF5 LABEL: [u16 len][raw bytes] at offset 6.
            let str_len = u16::from_le_bytes([data[6], data[7]]) as usize;
            let start = 8;
            let end = (start + str_len).min(data.len());
            data[start..end].iter().map(|&b| b as char).collect()
        },
    };
    Ok(vec![Cell {
        row,
        col,
        value: CellValue::String(s),
    }])
}

fn parse_blank(data: &[u8]) -> Result<Vec<Cell>> {
    if data.len() < 6 {
        return Err(XlsError::InvalidRecord("BLANK too short".into()));
    }
    let row = u16::from_le_bytes([data[0], data[1]]);
    let col = u16::from_le_bytes([data[2], data[3]]);
    Ok(vec![Cell {
        row,
        col,
        value: CellValue::Empty,
    }])
}

fn parse_mulblank(data: &[u8]) -> Result<Vec<Cell>> {
    if data.len() < 6 {
        return Err(XlsError::InvalidRecord("MULBLANK too short".into()));
    }
    let row = u16::from_le_bytes([data[0], data[1]]);
    let first_col = u16::from_le_bytes([data[2], data[3]]);
    let last_col = u16::from_le_bytes([data[data.len() - 2], data[data.len() - 1]]);
    let count = (last_col - first_col + 1) as usize;
    let cells = (0..count)
        .map(|i| Cell {
            row,
            col: first_col + i as u16,
            value: CellValue::Empty,
        })
        .collect();
    Ok(cells)
}

fn parse_formula(data: &[u8]) -> Result<Vec<Cell>> {
    // Use the cached result value from the FORMULA record.
    if data.len() < 14 {
        return Err(XlsError::InvalidRecord("FORMULA too short".into()));
    }
    let row = u16::from_le_bytes([data[0], data[1]]);
    let col = u16::from_le_bytes([data[2], data[3]]);
    // Cached result is at bytes 6..14 (8 bytes).
    // If byte 6 == 0xFF and byte 7 == 0xFF, it's a special type:
    //   byte 6 = value type: 0=string(in following STRING record), 1=bool, 2=error, 3=empty
    let val_bytes = &data[6..14];

    // Check if it's a special value (non-numeric).
    if val_bytes[6] == 0xFF && val_bytes[7] == 0xFF {
        let special_type = val_bytes[0];
        match special_type {
            0 => {
                // String follows in a STRING record — we'll handle this at a higher level.
                // For now, return empty.
                Ok(vec![Cell {
                    row,
                    col,
                    value: CellValue::String(String::new()),
                }])
            },
            1 => Ok(vec![Cell {
                row,
                col,
                value: CellValue::Bool(val_bytes[2] != 0),
            }]),
            2 => Ok(vec![Cell {
                row,
                col,
                value: CellValue::Error(val_bytes[2]),
            }]),
            3 => Ok(vec![Cell {
                row,
                col,
                value: CellValue::Empty,
            }]),
            _ => Ok(vec![Cell {
                row,
                col,
                value: CellValue::Empty,
            }]),
        }
    } else {
        // It's a regular IEEE 754 double.
        let value = f64::from_le_bytes([
            val_bytes[0],
            val_bytes[1],
            val_bytes[2],
            val_bytes[3],
            val_bytes[4],
            val_bytes[5],
            val_bytes[6],
            val_bytes[7],
        ]);
        Ok(vec![Cell {
            row,
            col,
            value: CellValue::Number(value),
        }])
    }
}

/// Decode an RK value to f64.
///
/// Bit 0: 0 = IEEE float, 1 = integer
/// Bit 1: 0 = not /100, 1 = value /100
/// Bits 2-31: the value
pub fn decode_rk(rk: u32) -> f64 {
    let is_integer = (rk & 0x02) != 0;
    let div_100 = (rk & 0x01) != 0;

    let val = if is_integer {
        // Signed 30-bit integer.
        (rk as i32 >> 2) as f64
    } else {
        // IEEE 754 double with top 30 bits from RK and bottom 34 bits zero.
        let bits = (rk as u64 & 0xFFFFFFFC) << 32;
        f64::from_bits(bits)
    };

    if div_100 { val / 100.0 } else { val }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn decode_rk_integer() {
        // Integer 42: value = 42 << 2 | 0x02 = 170
        let rk = (42u32 << 2) | 0x02;
        assert_eq!(decode_rk(rk), 42.0);
    }

    #[test]
    fn decode_rk_integer_div100() {
        // 1234 / 100 = 12.34, flags = 0x03
        let rk = (1234u32 << 2) | 0x03;
        let val = decode_rk(rk);
        assert!((val - 12.34).abs() < 1e-10);
    }

    #[test]
    fn decode_rk_float() {
        // Float: encode 1.0 as RK.
        // 1.0 in IEEE 754: 0x3FF0_0000_0000_0000
        // Top 30 bits: 0x3FF00000 >> 2 = keep bits 2-31 of 0x3FF00000
        let ieee_bits = 1.0f64.to_bits();
        let top32 = (ieee_bits >> 32) as u32;
        let rk = top32 & 0xFFFFFFFC; // clear bottom 2 bits
        assert_eq!(decode_rk(rk), 1.0);
    }

    #[test]
    fn cell_value_display() {
        assert_eq!(CellValue::Number(42.0).as_text(), "42");
        assert_eq!(CellValue::Number(3.15).as_text(), "3.15");
        assert_eq!(CellValue::String("hello".into()).as_text(), "hello");
        assert_eq!(CellValue::Bool(true).as_text(), "TRUE");
        assert_eq!(CellValue::Error(0x07).as_text(), "#DIV/0!");
        assert_eq!(CellValue::Empty.as_text(), "");
    }

    #[test]
    fn parse_labelsst_record() {
        let sst = vec!["Hello".into(), "World".into()];
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes()); // row 0
        data.extend_from_slice(&1u16.to_le_bytes()); // col 1
        data.extend_from_slice(&0u16.to_le_bytes()); // XF index
        data.extend_from_slice(&1u32.to_le_bytes()); // SST index 1
        let rec = BiffRecord {
            record_type: RT_LABELSST,
            data,
        };
        let cells = parse_cell_record(&rec, &sst).unwrap();
        assert_eq!(cells.len(), 1);
        assert_eq!(cells[0].row, 0);
        assert_eq!(cells[0].col, 1);
        assert_eq!(cells[0].value, CellValue::String("World".into()));
    }

    #[test]
    fn parse_number_record() {
        let mut data = Vec::new();
        data.extend_from_slice(&3u16.to_le_bytes()); // row 3
        data.extend_from_slice(&0u16.to_le_bytes()); // col 0
        data.extend_from_slice(&0u16.to_le_bytes()); // XF
        data.extend_from_slice(&42.5f64.to_le_bytes());
        let rec = BiffRecord {
            record_type: RT_NUMBER,
            data,
        };
        let cells = parse_cell_record(&rec, &[]).unwrap();
        assert_eq!(cells[0].value, CellValue::Number(42.5));
    }

    #[test]
    fn parse_boolerr_bool() {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes()); // XF
        data.push(1); // true
        data.push(0); // is_error = false (it's a bool)
        let rec = BiffRecord {
            record_type: RT_BOOLERR,
            data,
        };
        let cells = parse_cell_record(&rec, &[]).unwrap();
        assert_eq!(cells[0].value, CellValue::Bool(true));
    }

    #[test]
    fn parse_boolerr_error() {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.push(0x07); // #DIV/0!
        data.push(1); // is_error = true
        let rec = BiffRecord {
            record_type: RT_BOOLERR,
            data,
        };
        let cells = parse_cell_record(&rec, &[]).unwrap();
        assert_eq!(cells[0].value, CellValue::Error(0x07));
    }

    #[test]
    fn parse_mulrk_record() {
        let mut data = Vec::new();
        data.extend_from_slice(&5u16.to_le_bytes()); // row 5
        data.extend_from_slice(&0u16.to_le_bytes()); // first_col 0
        // RK entry 1: XF=0, RK = integer 10
        data.extend_from_slice(&0u16.to_le_bytes()); // XF
        let rk1 = (10u32 << 2) | 0x02;
        data.extend_from_slice(&rk1.to_le_bytes());
        // RK entry 2: XF=0, RK = integer 20
        data.extend_from_slice(&0u16.to_le_bytes());
        let rk2 = (20u32 << 2) | 0x02;
        data.extend_from_slice(&rk2.to_le_bytes());
        // last_col
        data.extend_from_slice(&1u16.to_le_bytes());

        let rec = BiffRecord {
            record_type: RT_MULRK,
            data,
        };
        let cells = parse_cell_record(&rec, &[]).unwrap();
        assert_eq!(cells.len(), 2);
        assert_eq!(cells[0].col, 0);
        assert_eq!(cells[0].value, CellValue::Number(10.0));
        assert_eq!(cells[1].col, 1);
        assert_eq!(cells[1].value, CellValue::Number(20.0));
    }
}
