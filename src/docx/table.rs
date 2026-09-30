use crate::core::units::Twip;

use super::document::BlockElement;
use super::formatting::{Justification, TableBorders};

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
    /// Table border edges (`w:tblBorders`).
    pub borders: Option<TableBorders>,
    /// Default cell margins (`w:tblCellMar`), in twips.
    pub cell_margins: Option<CellMargins>,
    /// Table indent from the left margin (`w:tblInd`), in twips.
    pub indent: Option<Twip>,
    /// Accessibility caption (`w:tblCaption`).
    pub caption: Option<String>,
    /// Which of the table style's conditional formats apply (`w:tblLook`).
    pub look: Option<TableLook>,
    /// Rows per band for banded-row formatting (`w:tblStyleRowBandSize`,
    /// set in a table style's `w:tblPr`); `None` means 1.
    pub row_band_size: Option<u32>,
    /// Columns per band (`w:tblStyleColBandSize`); `None` means 1.
    pub col_band_size: Option<u32>,
}

/// `w:tblLook` (ECMA-376 §17.4): which conditional formats of the table's
/// style are switched on. Each flag is off unless the table says
/// otherwise.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TableLook {
    /// Header-row formatting (`firstRow`).
    pub first_row: bool,
    /// Total-row formatting (`lastRow`).
    pub last_row: bool,
    /// First-column formatting (`firstColumn`).
    pub first_column: bool,
    /// Last-column formatting (`lastColumn`).
    pub last_column: bool,
    /// Banded rows switched off (`noHBand`).
    pub no_h_band: bool,
    /// Banded columns switched off (`noVBand`).
    pub no_v_band: bool,
}

impl TableLook {
    /// `w:tblLook`: the explicit attributes, or the older `w:val` bitmask
    /// (firstRow 0x0020, lastRow 0x0040, firstColumn 0x0080, lastColumn
    /// 0x0100, noHBand 0x0200, noVBand 0x0400), which an explicit
    /// attribute overrides.
    pub(crate) fn parse(e: &quick_xml::events::BytesStart) -> Self {
        let mut look = TableLook::default();
        if let Ok(Some(v)) = crate::core::xml::optional_attr_str(e, "w:val") {
            if let Ok(bits) = u16::from_str_radix(v.trim(), 16) {
                look.first_row = bits & 0x0020 != 0;
                look.last_row = bits & 0x0040 != 0;
                look.first_column = bits & 0x0080 != 0;
                look.last_column = bits & 0x0100 != 0;
                look.no_h_band = bits & 0x0200 != 0;
                look.no_v_band = bits & 0x0400 != 0;
            }
        }
        for (attr, flag) in [
            ("w:firstRow", &mut look.first_row),
            ("w:lastRow", &mut look.last_row),
            ("w:firstColumn", &mut look.first_column),
            ("w:lastColumn", &mut look.last_column),
            ("w:noHBand", &mut look.no_h_band),
            ("w:noVBand", &mut look.no_v_band),
        ] {
            if let Ok(Some(v)) = crate::core::xml::optional_attr_str(e, attr) {
                *flag = matches!(v.as_ref(), "1" | "true" | "on");
            }
        }
        look
    }
}

/// Cell margins from `w:tblCellMar` / `w:tcMar`, in twips.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CellMargins {
    /// Top margin.
    pub top: Option<i32>,
    /// Bottom margin.
    pub bottom: Option<i32>,
    /// Left (start) margin.
    pub left: Option<i32>,
    /// Right (end) margin.
    pub right: Option<i32>,
}

/// How a `w:trHeight` value is applied.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RowHeightRule {
    /// Height is a minimum; the row grows to fit content.
    AtLeast,
    /// Height is fixed.
    Exact,
    /// Height is determined by content.
    Auto,
}

/// Table row properties (`w:trPr`).
#[derive(Debug, Clone, Default)]
pub struct TableRowProperties {
    /// Whether this row is a table header row.
    pub is_header: bool,
    /// Row height from `w:trHeight`, in twips.
    pub height: Option<i32>,
    /// How `height` is applied.
    pub height_rule: Option<RowHeightRule>,
    /// `<w:cantSplit/>` — the row may not break across pages.
    pub cant_split: bool,
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
    /// Cell shading/background. Boxed: rare on real cells in a large
    /// table — `Option<T>` reserves `size_of(T)` even when `None`.
    pub shading: Option<Box<Shading>>,
    /// Cell border edges (`w:tcBorders`). Boxed: rare on real cells in a
    /// large table.
    pub borders: Option<Box<TableBorders>>,
    /// Vertical alignment of cell content (`w:vAlign`).
    pub v_align: Option<CellVAlign>,
    /// Text flow direction (`w:textDirection`).
    pub text_direction: Option<String>,
    /// Per-cell margin overrides (`w:tcMar`), in twips.
    pub margins: Option<CellMargins>,
    /// `true` when the cell carries `<w:cellDel>` — deleted via tracked
    /// changes, pending acceptance. Mirrors the policy already applied to
    /// run-level `w:del`: excluded from the accepted view.
    pub deleted: bool,
}

/// Vertical alignment of a cell's content (`w:vAlign`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellVAlign {
    /// Align to the top of the cell.
    Top,
    /// Center vertically.
    Center,
    /// Align to the bottom of the cell.
    Bottom,
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

#[cfg(test)]
mod tests {
    use super::*;

    /// Regression: `TableCellProperties`'s `shading`/
    /// `borders` must stay boxed. `TableCellProperties` was 536 bytes
    /// before this fix (dominated by an unboxed `Shading` + `TableBorders`
    /// pair), and `TableCell` (which every cell in a large table pays for)
    /// was 560 bytes. Headroom above the measured post-fix sizes (96B /
    /// 120B) keeps this from being flaky on an unrelated new field, while
    /// still catching the actual regression: a large struct un-boxed back
    /// into `Option<T>`.
    #[test]
    fn test_table_cell_properties_size_stays_boxed() {
        assert!(
            std::mem::size_of::<TableCellProperties>() <= 200,
            "TableCellProperties grew to {} bytes — check shading/borders \
             are still Option<Box<T>>, not Option<T>",
            std::mem::size_of::<TableCellProperties>()
        );
        assert!(
            std::mem::size_of::<TableCell>() <= 250,
            "TableCell grew to {} bytes — a TableCellProperties field \
             regression would show up here too",
            std::mem::size_of::<TableCell>()
        );
    }
}
