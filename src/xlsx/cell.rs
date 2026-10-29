use std::fmt;

use super::date::DateTimeValue;

/// Zero-based column/row indices for a cell reference like "A1".
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellRef {
    /// 0-based column index (A=0, B=1, ..., XFD=16383).
    pub col: u32,
    /// 0-based row index (displayed as 1-based in Excel).
    pub row: u32,
}

/// The last column of the grid, `XFD` (0-based).
pub const MAX_COL: u32 = 16_383;
/// The last row of the grid (1-based).
pub const MAX_ROWS: u32 = 1_048_576;

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
        // Excel's grid ends at XFD1048576. A reference past it is not a
        // cell: `ZZZZZZ1` parsed to column 321,272,405 and the IR
        // converter sized every row of the table to it — a 95 GB
        // allocation from a 2 KB package, and an abort.
        if row == 0 || row > MAX_ROWS || col > MAX_COL {
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
    /// Cell is empty.
    Empty,
    /// Numeric value.
    Number(f64),
    /// Inline or resolved string.
    String(String),
    /// Index into the shared string table (resolved after SST is loaded).
    SharedString(u32),
    /// Boolean value.
    Boolean(bool),
    /// Error value like `#DIV/0!`, `#REF!`, etc.
    Error(String),
    /// Resolved from a number + date format detection.
    Date(DateTimeValue),
}

/// A raw cell as parsed from worksheet XML.
///
/// Most cells are a value and a style. The three attributes that few cells
/// carry (a formula, rich-text runs, a value-metadata index) live in one boxed
/// extension, so the common cell is 56 bytes instead of 96: a worksheet of
/// millions of cells holds millions of these. Read them through
/// [`Cell::formula`], [`Cell::rich_runs`] and [`Cell::vm`].
#[derive(Debug, Clone)]
pub struct Cell {
    /// Parsed cell reference (column, row).
    pub reference: CellRef,
    /// The cell's value.
    pub value: CellValue,
    /// Style index (`s` attribute) referencing the cellXfs array.
    pub style_index: Option<u32>,
    /// What few cells have; `None` for the rest.
    extras: Option<Box<CellExtras>>,
}

/// The attributes of a [`Cell`] that few cells carry.
#[derive(Debug, Clone, Default)]
struct CellExtras {
    formula: Option<String>,
    rich_runs: Option<std::sync::Arc<[crate::xlsx::shared_strings::RichTextRun]>>,
    vm: Option<u32>,
}

impl Cell {
    /// A cell with no formula, rich runs or value metadata.
    pub fn new(reference: CellRef, value: CellValue, style_index: Option<u32>) -> Self {
        Self {
            reference,
            value,
            style_index,
            extras: None,
        }
    }

    /// Formula content from the `<f>` element, if present.
    pub fn formula(&self) -> Option<&str> {
        self.extras.as_ref()?.formula.as_deref()
    }

    /// Rich-run formatting for an inline (`t="inlineStr"`) string cell: the
    /// mirror of `SharedString.rich_text` for a cell whose text lives directly
    /// in the `<c>` element instead of the shared string table.
    pub fn rich_runs(&self) -> Option<&std::sync::Arc<[crate::xlsx::shared_strings::RichTextRun]>> {
        self.extras.as_ref()?.rich_runs.as_ref()
    }

    /// The cell's `vm` attribute: a 1-based index into the workbook's
    /// `xl/metadata.xml` `<valueMetadata>` array. Excel 365's in-cell
    /// rich-value images (`=IMAGE(...)`/"Place in Cell") are stored as a
    /// `t="e"` cell with a literal `<v>#VALUE!</v>` fallback for old readers,
    /// plus this pointer to the real rich-value chain.
    pub fn vm(&self) -> Option<u32> {
        self.extras.as_ref()?.vm
    }

    /// Set or clear the formula.
    pub fn set_formula(&mut self, formula: Option<String>) {
        if formula.is_some() || self.extras.is_some() {
            self.extras.get_or_insert_with(Default::default).formula = formula;
        }
    }

    /// Set or clear the rich runs.
    pub fn set_rich_runs(
        &mut self,
        runs: Option<std::sync::Arc<[crate::xlsx::shared_strings::RichTextRun]>>,
    ) {
        if runs.is_some() || self.extras.is_some() {
            self.extras.get_or_insert_with(Default::default).rich_runs = runs;
        }
    }

    /// Set or clear the value-metadata index.
    pub fn set_vm(&mut self, vm: Option<u32>) {
        if vm.is_some() || self.extras.is_some() {
            self.extras.get_or_insert_with(Default::default).vm = vm;
        }
    }

    /// Take the formula out of the cell.
    pub fn take_formula(&mut self) -> Option<String> {
        self.extras.as_mut()?.formula.take()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A worksheet of millions of cells holds millions of these: the three
    /// attributes few cells carry are boxed, and a regression that inlines
    /// them again fails here.
    #[test]
    fn test_a_plain_cell_is_small() {
        assert!(std::mem::size_of::<Cell>() <= 56, "{}", std::mem::size_of::<Cell>());
        let mut cell = Cell::new(CellRef { col: 0, row: 0 }, CellValue::Number(1.0), None);
        assert_eq!((cell.formula(), cell.vm()), (None, None));
        assert!(cell.rich_runs().is_none());
        // Clearing what was never set allocates nothing.
        cell.set_formula(None);
        assert!(cell.extras.is_none());
        cell.set_formula(Some("A1+1".into()));
        cell.set_vm(Some(3));
        assert_eq!((cell.formula(), cell.vm()), (Some("A1+1"), Some(3)));
        assert_eq!(cell.take_formula().as_deref(), Some("A1+1"));
        assert_eq!(cell.formula(), None);
    }

    #[test]
    fn test_parse_a1() {
        let r = CellRef::parse("A1").unwrap();
        assert_eq!(r.col, 0);
        assert_eq!(r.row, 0);
    }

    #[test]
    fn test_parse_z1() {
        let r = CellRef::parse("Z1").unwrap();
        assert_eq!(r.col, 25);
        assert_eq!(r.row, 0);
    }

    #[test]
    fn test_parse_aa1() {
        let r = CellRef::parse("AA1").unwrap();
        assert_eq!(r.col, 26);
        assert_eq!(r.row, 0);
    }

    #[test]
    fn test_parse_xfd1048576() {
        let r = CellRef::parse("XFD1048576").unwrap();
        assert_eq!(r.col, 16383);
        assert_eq!(r.row, 1048575);
    }

    #[test]
    fn test_col_name_round_trip() {
        assert_eq!(CellRef::col_name(0), "A");
        assert_eq!(CellRef::col_name(25), "Z");
        assert_eq!(CellRef::col_name(26), "AA");
        assert_eq!(CellRef::col_name(16383), "XFD");
    }

    #[test]
    fn test_display_round_trip() {
        let r = CellRef { col: 0, row: 0 };
        assert_eq!(r.to_string(), "A1");

        let r = CellRef { col: 26, row: 99 };
        assert_eq!(r.to_string(), "AA100");
    }

    #[test]
    fn test_parse_col_values() {
        assert_eq!(CellRef::parse_col("A"), Some(0));
        assert_eq!(CellRef::parse_col("Z"), Some(25));
        assert_eq!(CellRef::parse_col("AA"), Some(26));
        assert_eq!(CellRef::parse_col("AZ"), Some(51));
        assert_eq!(CellRef::parse_col("BA"), Some(52));
    }

    #[test]
    fn test_parse_invalid() {
        assert!(CellRef::parse("").is_none());
        assert!(CellRef::parse("1").is_none());
        assert!(CellRef::parse("A").is_none());
        assert!(CellRef::parse("A0").is_none());
    }

    /// The grid ends at XFD1048576; a reference past it is not a cell.
    /// `ZZZZZZ1` parsed to column 321,272,405 and the converter sized
    /// every row to it.
    #[test]
    fn test_references_past_the_grid_are_not_cells() {
        assert_eq!(
            CellRef::parse("XFD1048576"),
            Some(CellRef {
                col: 16_383,
                row: 1_048_575
            })
        );
        assert!(CellRef::parse("XFE1").is_none());
        assert!(CellRef::parse("A1048577").is_none());
        assert!(CellRef::parse("ZZZZZZ1").is_none());
    }
}
