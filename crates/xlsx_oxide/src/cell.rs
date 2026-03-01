use std::fmt;

use crate::date::DateTimeValue;

/// Zero-based column/row indices for a cell reference like "A1".
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellRef {
    /// 0-based column index (A=0, B=1, ..., XFD=16383).
    pub col: u32,
    /// 0-based row index (displayed as 1-based in Excel).
    pub row: u32,
}

impl CellRef {
    /// Parse a cell reference string like "A1" into zero-based indices.
    pub fn parse(reference: &str) -> Option<Self> {
        let bytes = reference.as_bytes();
        // Find where letters end and digits begin
        let col_end = bytes
            .iter()
            .position(|b| b.is_ascii_digit())
            .filter(|&p| p > 0)?;

        let col_str = &reference[..col_end];
        let row_str = &reference[col_end..];

        let col = Self::parse_col(col_str)?;
        let row: u32 = row_str.parse().ok()?;
        if row == 0 {
            return None;
        }

        Some(CellRef { col, row: row - 1 })
    }

    /// Convert a column letter string to a 0-based index: "A" -> 0, "Z" -> 25, "AA" -> 26.
    pub fn parse_col(col_str: &str) -> Option<u32> {
        if col_str.is_empty() {
            return None;
        }
        let mut result: u32 = 0;
        for &b in col_str.as_bytes() {
            if !b.is_ascii_alphabetic() {
                return None;
            }
            let digit = (b.to_ascii_uppercase() - b'A') as u32;
            result = result.checked_mul(26)?.checked_add(digit + 1)?;
        }
        Some(result - 1)
    }

    /// Convert a 0-based column index to a column letter string: 0 -> "A", 25 -> "Z", 26 -> "AA".
    pub fn col_name(col: u32) -> String {
        let mut result = Vec::new();
        let mut n = col + 1; // 1-based for the math
        while n > 0 {
            n -= 1;
            result.push(b'A' + (n % 26) as u8);
            n /= 26;
        }
        result.reverse();
        String::from_utf8(result).unwrap()
    }
}

impl fmt::Display for CellRef {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}{}", Self::col_name(self.col), self.row + 1)
    }
}

/// Parsed cell value (after shared string dereferencing).
#[derive(Debug, Clone)]
pub enum CellValue {
    Empty,
    Number(f64),
    String(String),
    /// Index into the shared string table (resolved after SST is loaded).
    SharedString(u32),
    Boolean(bool),
    /// Error value like #DIV/0!, #REF!, etc.
    Error(String),
    /// Resolved from a number + date format detection.
    Date(DateTimeValue),
}

/// A raw cell as parsed from worksheet XML.
#[derive(Debug, Clone)]
pub struct Cell {
    /// Parsed cell reference (column, row).
    pub reference: CellRef,
    /// The cell's value.
    pub value: CellValue,
    /// Style index (`s` attribute) referencing the cellXfs array.
    pub style_index: Option<u32>,
    /// Formula content from the `<f>` element, if present.
    pub formula: Option<String>,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_a1() {
        let r = CellRef::parse("A1").unwrap();
        assert_eq!(r.col, 0);
        assert_eq!(r.row, 0);
    }

    #[test]
    fn parse_z1() {
        let r = CellRef::parse("Z1").unwrap();
        assert_eq!(r.col, 25);
        assert_eq!(r.row, 0);
    }

    #[test]
    fn parse_aa1() {
        let r = CellRef::parse("AA1").unwrap();
        assert_eq!(r.col, 26);
        assert_eq!(r.row, 0);
    }

    #[test]
    fn parse_xfd1048576() {
        let r = CellRef::parse("XFD1048576").unwrap();
        assert_eq!(r.col, 16383);
        assert_eq!(r.row, 1048575);
    }

    #[test]
    fn col_name_round_trip() {
        assert_eq!(CellRef::col_name(0), "A");
        assert_eq!(CellRef::col_name(25), "Z");
        assert_eq!(CellRef::col_name(26), "AA");
        assert_eq!(CellRef::col_name(16383), "XFD");
    }

    #[test]
    fn display_round_trip() {
        let r = CellRef { col: 0, row: 0 };
        assert_eq!(r.to_string(), "A1");

        let r = CellRef {
            col: 26,
            row: 99,
        };
        assert_eq!(r.to_string(), "AA100");
    }

    #[test]
    fn parse_col_values() {
        assert_eq!(CellRef::parse_col("A"), Some(0));
        assert_eq!(CellRef::parse_col("Z"), Some(25));
        assert_eq!(CellRef::parse_col("AA"), Some(26));
        assert_eq!(CellRef::parse_col("AZ"), Some(51));
        assert_eq!(CellRef::parse_col("BA"), Some(52));
    }

    #[test]
    fn parse_invalid() {
        assert!(CellRef::parse("").is_none());
        assert!(CellRef::parse("1").is_none());
        assert!(CellRef::parse("A").is_none());
        assert!(CellRef::parse("A0").is_none());
    }
}
