use crate::core::units::Twip;

use super::document::BlockElement;
use super::formatting::Justification;

/// A table element (`w:tbl`).
#[derive(Debug, Clone)]
pub struct Table {
    pub properties: Option<TableProperties>,
    /// Column widths from `w:tblGrid/w:gridCol`.
    pub grid: Vec<Twip>,
    pub rows: Vec<TableRow>,
}

/// A table row (`w:tr`).
#[derive(Debug, Clone)]
pub struct TableRow {
    pub properties: Option<TableRowProperties>,
    pub cells: Vec<TableCell>,
}

/// A table cell (`w:tc`).
#[derive(Debug, Clone)]
pub struct TableCell {
    pub properties: Option<TableCellProperties>,
    pub content: Vec<BlockElement>,
}

/// Table-level properties (`w:tblPr`).
#[derive(Debug, Clone, Default)]
pub struct TableProperties {
    pub width: Option<TableWidth>,
    pub justification: Option<Justification>,
    pub style_id: Option<String>,
}

/// Table row properties (`w:trPr`).
#[derive(Debug, Clone, Default)]
pub struct TableRowProperties {
    pub is_header: bool,
}

/// Table cell properties (`w:tcPr`).
#[derive(Debug, Clone, Default)]
pub struct TableCellProperties {
    pub width: Option<TableWidth>,
    pub vertical_merge: Option<MergeType>,
    pub grid_span: Option<u32>,
    pub shading: Option<Shading>,
}

/// Width specification for tables/cells.
#[derive(Debug, Clone)]
pub struct TableWidth {
    pub value: i32,
    pub width_type: TableWidthType,
}

/// How the table/cell width value is interpreted.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableWidthType {
    /// Width in fiftieths of a percent.
    Pct,
    /// Width in twips.
    Dxa,
    /// Automatically determined.
    Auto,
    /// No width specified.
    Nil,
}

/// Vertical merge type for table cells.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MergeType {
    /// Start of a vertical merge.
    Restart,
    /// Continuation of a vertical merge.
    Continue,
}

/// Cell/paragraph shading.
#[derive(Debug, Clone)]
pub struct Shading {
    pub fill: Option<String>,
    pub color: Option<String>,
    pub pattern: Option<String>,
}
