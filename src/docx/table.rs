use crate::core::units::Twip;

use super::document::BlockElement;
use super::formatting::Justification;

/// A table element (`w:tbl`).
#[derive(Debug, Clone)]
pub struct Table {
    /// Table-level formatting.
    pub properties: Option<TableProperties>,
    /// Column widths from `w:tblGrid/w:gridCol`.
    pub grid: Vec<Twip>,
    /// Rows in the table.
    pub rows: Vec<TableRow>,
}

/// A table row (`w:tr`).
#[derive(Debug, Clone)]
pub struct TableRow {
    /// Row-level formatting.
    pub properties: Option<TableRowProperties>,
    /// Cells in this row.
    pub cells: Vec<TableCell>,
}

/// A table cell (`w:tc`).
#[derive(Debug, Clone)]
pub struct TableCell {
    /// Cell-level formatting.
    pub properties: Option<TableCellProperties>,
    /// Block content within the cell.
    pub content: Vec<BlockElement>,
}

/// Table-level properties (`w:tblPr`).
#[derive(Debug, Clone, Default)]
pub struct TableProperties {
    /// Preferred table width.
    pub width: Option<TableWidth>,
    /// Table justification.
    pub justification: Option<Justification>,
    /// Applied table style ID.
    pub style_id: Option<String>,
}

/// Table row properties (`w:trPr`).
#[derive(Debug, Clone, Default)]
pub struct TableRowProperties {
    /// Whether this row is a table header row.
    pub is_header: bool,
}

/// Table cell properties (`w:tcPr`).
#[derive(Debug, Clone, Default)]
pub struct TableCellProperties {
    /// Preferred cell width.
    pub width: Option<TableWidth>,
    /// Vertical merge type (for spanning rows).
    pub vertical_merge: Option<MergeType>,
    /// Horizontal grid span (number of columns spanned).
    pub grid_span: Option<u32>,
    /// Cell shading/background.
    pub shading: Option<Shading>,
}

/// Width specification for tables/cells.
#[derive(Debug, Clone)]
pub struct TableWidth {
    /// Numeric width value (interpretation depends on `width_type`).
    pub value: i32,
    /// How `value` is measured.
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
    /// Background fill color (hex or "auto").
    pub fill: Option<String>,
    /// Foreground/pattern color.
    pub color: Option<String>,
    /// Shading pattern value.
    pub pattern: Option<String>,
}
