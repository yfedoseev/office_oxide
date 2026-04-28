//! XLSX creation (write) module.
//!
//! Provides a builder API for creating XLSX files from scratch.
//!
//! # Example
//!
//! ```rust,no_run
//! use office_oxide::xlsx::write::{XlsxWriter, CellData, CellStyle, NumberFormat, HAlign};
//!
//! let mut wb = XlsxWriter::new();
//! let mut sheet = wb.add_sheet("Sales");
//!
//! // Bold header row
//! let header_style = CellStyle::new().bold();
//! sheet.set_cell_styled(0, 0, CellData::String("Item".into()),  header_style.clone());
//! sheet.set_cell_styled(0, 1, CellData::String("Amount".into()), header_style);
//!
//! // Data rows
//! sheet.set_cell(1, 0, CellData::String("Widget".into()));
//! sheet.set_cell(1, 1, CellData::Number(1500.0));
//! sheet.set_cell(2, 0, CellData::String("Gadget".into()));
//! sheet.set_cell(2, 1, CellData::Number(2400.0));
//!
//! // SUM formula with currency formatting
//! let currency = CellStyle::new().number_format(NumberFormat::Currency);
//! sheet.set_cell_styled(3, 1, CellData::Formula("SUM(B2:B3)".into()), currency);
//!
//! sheet.set_column_width(0, 20.0);
//! sheet.set_column_width(1, 15.0);
//!
//! wb.save("sales.xlsx").unwrap();
//! ```

use std::collections::HashMap;
use std::io::{Seek, Write};
use std::path::Path;

use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};

use crate::core::opc::{OpcWriter, PartName};
use crate::core::relationships::rel_types;

use super::Result;

// ---------------------------------------------------------------------------
// SML namespace constants
// ---------------------------------------------------------------------------

use crate::core::xml::ns::{R_STR as NS_REL, SML_STR as NS_SML};

// ---------------------------------------------------------------------------
// Content type constants
// ---------------------------------------------------------------------------

const CT_WORKBOOK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_STYLES: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";

// ---------------------------------------------------------------------------
// Public: CellStyle
// ---------------------------------------------------------------------------

/// Horizontal alignment inside a cell.
#[derive(Debug, Clone, PartialEq)]
pub enum HAlign {
    /// Align content to the left edge.
    Left,
    /// Center content horizontally.
    Center,
    /// Align content to the right edge.
    Right,
}

/// Pre-defined number formats.  Custom format strings can be set via
/// `CellStyle::custom_format`.
#[derive(Debug, Clone, Default, PartialEq)]
pub enum NumberFormat {
    /// General text (default).
    #[default]
    General,
    /// Integer: `0`
    Integer,
    /// Two decimal places: `0.00`
    Decimal2,
    /// Currency: `#,##0.00`
    Currency,
    /// Percentage: `0%`
    Percent,
    /// Percentage with two decimals: `0.00%`
    Percent2,
    /// Short date: `yyyy-mm-dd`
    Date,
    /// Short datetime: `yyyy-mm-dd hh:mm`
    DateTime,
}

impl NumberFormat {
    /// Return the OOXML built-in `numFmtId` (0 = general) or `None` if a
    /// custom format string must be emitted.
    fn builtin_id(&self) -> Option<u32> {
        match self {
            Self::General => Some(0),
            Self::Integer => Some(1),
            Self::Decimal2 => Some(2),
            Self::Currency => Some(7),  // #,##0.00
            Self::Percent => Some(9),
            Self::Percent2 => Some(10),
            Self::Date => Some(14),
            Self::DateTime => Some(22),
        }
    }
}

/// Style applied to a single cell.
#[derive(Debug, Clone, Default)]
pub struct CellStyle {
    /// Apply bold font weight.
    pub bold: bool,
    /// Apply italic style.
    pub italic: bool,
    /// Apply single underline.
    pub underline: bool,
    /// RGB hex color for the font, e.g. `"FF0000"` for red. No leading `#`.
    pub font_color: Option<String>,
    /// Font size in points.
    pub font_size_pt: Option<f32>,
    /// Font family name, e.g. `"Arial"`.
    pub font_name: Option<String>,
    /// Background fill color (RGB hex, no `#`).
    pub background_color: Option<String>,
    /// Number format to apply to the cell value.
    pub number_format: NumberFormat,
    /// Horizontal alignment override.
    pub h_align: Option<HAlign>,
    /// Wrap text within the cell.
    pub wrap_text: bool,
}

impl CellStyle {
    /// Create a new default (unstyled) `CellStyle`.
    pub fn new() -> Self {
        Self::default()
    }

    /// Enable bold. Builder-style: returns `Self`.
    pub fn bold(mut self) -> Self {
        self.bold = true;
        self
    }

    /// Enable italic.
    pub fn italic(mut self) -> Self {
        self.italic = true;
        self
    }

    /// Enable underline.
    pub fn underline(mut self) -> Self {
        self.underline = true;
        self
    }

    /// Set font color (RGB hex without `#`, e.g. `"FF0000"`).
    pub fn font_color(mut self, color: impl Into<String>) -> Self {
        self.font_color = Some(color.into());
        self
    }

    /// Set font size in points.
    pub fn font_size(mut self, pt: f32) -> Self {
        self.font_size_pt = Some(pt);
        self
    }

    /// Set font family.
    pub fn font_name(mut self, name: impl Into<String>) -> Self {
        self.font_name = Some(name.into());
        self
    }

    /// Set background fill color (RGB hex without `#`).
    pub fn background(mut self, color: impl Into<String>) -> Self {
        self.background_color = Some(color.into());
        self
    }

    /// Set a number format.
    pub fn number_format(mut self, fmt: NumberFormat) -> Self {
        self.number_format = fmt;
        self
    }

    /// Set horizontal alignment.
    pub fn align(mut self, align: HAlign) -> Self {
        self.h_align = Some(align);
        self
    }

    /// Enable text wrapping.
    pub fn wrap(mut self) -> Self {
        self.wrap_text = true;
        self
    }
}

// ---------------------------------------------------------------------------
// Public: CellData
// ---------------------------------------------------------------------------

/// A value to write into a cell.
#[derive(Debug, Clone)]
pub enum CellData {
    /// An empty cell.
    Empty,
    /// A string value (written as an inline string).
    String(String),
    /// A numeric value.
    Number(f64),
    /// A boolean value.
    Boolean(bool),
    /// A formula, e.g. `"SUM(A1:A10)"`. Do not include the leading `=`.
    Formula(String),
}

// ---------------------------------------------------------------------------
// XlsxWriter
// ---------------------------------------------------------------------------

/// Builder for creating XLSX files.
pub struct XlsxWriter {
    sheets: Vec<SheetDataInner>,
}

/// Full internal sheet representation.
struct SheetDataInner {
    pub name: String,
    pub rows: Vec<Vec<Option<StoredCellInner>>>,
    pub col_widths: HashMap<usize, f64>,
    pub cell_styles: HashMap<(usize, usize), CellStyle>,
}

impl SheetDataInner {
    fn new(name: &str) -> Self {
        Self {
            name: name.to_string(),
            rows: Vec::new(),
            col_widths: HashMap::new(),
            cell_styles: HashMap::new(),
        }
    }

    fn ensure_cell(&mut self, row: usize, col: usize) {
        if self.rows.len() <= row {
            self.rows.resize_with(row + 1, Vec::new);
        }
        if self.rows[row].len() <= col {
            self.rows[row].resize_with(col + 1, || None);
        }
    }

    pub fn add_row(&mut self, cells: Vec<CellData>) -> &mut Self {
        let stored: Vec<Option<StoredCellInner>> = cells
            .into_iter()
            .map(|v| Some(StoredCellInner { value: v }))
            .collect();
        self.rows.push(stored);
        self
    }

    pub fn set_cell(&mut self, row: usize, col: usize, value: CellData) -> &mut Self {
        self.ensure_cell(row, col);
        self.rows[row][col] = Some(StoredCellInner { value });
        self
    }

    pub fn set_cell_styled(
        &mut self,
        row: usize,
        col: usize,
        value: CellData,
        style: CellStyle,
    ) -> &mut Self {
        self.ensure_cell(row, col);
        self.rows[row][col] = Some(StoredCellInner { value });
        self.cell_styles.insert((row, col), style);
        self
    }

    pub fn set_column_width(&mut self, col: usize, width: f64) -> &mut Self {
        self.col_widths.insert(col, width);
        self
    }
}

struct StoredCellInner {
    value: CellData,
}

// ---------------------------------------------------------------------------
// Public SheetData — wraps SheetDataInner via &mut reference returned by
// XlsxWriter::add_sheet so callers get a single ergonomic type.
// We expose SheetData as a thin newtype.
// ---------------------------------------------------------------------------

/// Data for a single worksheet — returned by `XlsxWriter::add_sheet`.
pub struct SheetData<'a>(&'a mut SheetDataInner);

impl<'a> SheetData<'a> {
    /// Append a row of `CellData` values.
    pub fn add_row(&mut self, cells: Vec<CellData>) -> &mut Self {
        self.0.add_row(cells);
        self
    }

    /// Set a cell value (0-based row and column).
    pub fn set_cell(&mut self, row: usize, col: usize, value: CellData) -> &mut Self {
        self.0.set_cell(row, col, value);
        self
    }

    /// Set a cell value with explicit formatting.
    ///
    /// # Example
    /// ```rust,no_run
    /// # use office_oxide::xlsx::write::*;
    /// # let mut wb = XlsxWriter::new();
    /// # let mut sheet = wb.add_sheet("S");
    /// sheet.set_cell_styled(0, 0, CellData::String("Total".into()),
    ///     CellStyle::new().bold().align(HAlign::Right));
    /// sheet.set_cell_styled(0, 1, CellData::Formula("SUM(B2:B100)".into()),
    ///     CellStyle::new().bold().number_format(NumberFormat::Currency));
    /// ```
    pub fn set_cell_styled(
        &mut self,
        row: usize,
        col: usize,
        value: CellData,
        style: CellStyle,
    ) -> &mut Self {
        self.0.set_cell_styled(row, col, value, style);
        self
    }

    /// Set the width of a column in character units.
    pub fn set_column_width(&mut self, col: usize, width: f64) -> &mut Self {
        self.0.set_column_width(col, width);
        self
    }
}

// ---------------------------------------------------------------------------
// XlsxWriter impl
// ---------------------------------------------------------------------------

impl Default for XlsxWriter {
    fn default() -> Self {
        Self::new()
    }
}

impl XlsxWriter {
    /// Create a new, empty XLSX writer.
    pub fn new() -> Self {
        Self { sheets: Vec::new() }
    }

    /// Add a worksheet and return a mutable handle to it.
    pub fn add_sheet(&mut self, name: &str) -> SheetData<'_> {
        self.sheets.push(SheetDataInner::new(name));
        SheetData(self.sheets.last_mut().unwrap())
    }

    /// Save the workbook to a file.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        let mut opc = OpcWriter::create(path)?;
        self.write_parts(&mut opc)?;
        opc.finish()?;
        Ok(())
    }

    /// Write the workbook to any `Write + Seek` destination.
    pub fn write_to<W: Write + Seek>(&self, writer: W) -> Result<()> {
        let mut opc = OpcWriter::new(writer)?;
        self.write_parts(&mut opc)?;
        opc.finish()?;
        Ok(())
    }

    // -----------------------------------------------------------------------
    // Internal
    // -----------------------------------------------------------------------

    fn write_parts<W: Write + Seek>(&self, opc: &mut OpcWriter<W>) -> Result<()> {
        let wb_part = PartName::new("/xl/workbook.xml")?;

        opc.add_package_rel(rel_types::OFFICE_DOCUMENT, "xl/workbook.xml");

        let mut sheet_rids = Vec::with_capacity(self.sheets.len());
        for (i, _) in self.sheets.iter().enumerate() {
            let target = format!("worksheets/sheet{}.xml", i + 1);
            let rid = opc.add_part_rel(&wb_part, rel_types::WORKSHEET, &target);
            sheet_rids.push(rid);
        }
        opc.add_part_rel(&wb_part, rel_types::STYLES, "styles.xml");

        let wb_xml = self.build_workbook_xml(&sheet_rids)?;
        opc.add_part(&wb_part, CT_WORKBOOK, &wb_xml)?;

        // Collect all unique styles across all sheets, assign indices.
        let style_table = StyleTable::build(&self.sheets);

        for (i, sheet) in self.sheets.iter().enumerate() {
            let part_name_str = format!("/xl/worksheets/sheet{}.xml", i + 1);
            let part_name = PartName::new(&part_name_str)?;
            let ws_xml = Self::build_worksheet_xml(sheet, &style_table)?;
            opc.add_part(&part_name, CT_WORKSHEET, &ws_xml)?;
        }

        let styles_part = PartName::new("/xl/styles.xml")?;
        let styles_xml = style_table.build_styles_xml()?;
        opc.add_part(&styles_part, CT_STYLES, &styles_xml)?;

        Ok(())
    }

    fn build_workbook_xml(&self, sheet_rids: &[String]) -> crate::core::Result<Vec<u8>> {
        let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

        w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))?;

        let mut root = BytesStart::new("workbook");
        root.push_attribute(("xmlns", NS_SML));
        root.push_attribute(("xmlns:r", NS_REL));
        w.write_event(Event::Start(root))?;

        w.write_event(Event::Start(BytesStart::new("sheets")))?;

        for (i, sheet) in self.sheets.iter().enumerate() {
            let mut elem = BytesStart::new("sheet");
            elem.push_attribute(("name", sheet.name.as_str()));
            let sheet_id = (i + 1).to_string();
            elem.push_attribute(("sheetId", sheet_id.as_str()));
            elem.push_attribute(("r:id", sheet_rids[i].as_str()));
            w.write_event(Event::Empty(elem))?;
        }

        w.write_event(Event::End(BytesEnd::new("sheets")))?;
        w.write_event(Event::End(BytesEnd::new("workbook")))?;

        Ok(w.into_inner())
    }

    fn build_worksheet_xml(
        sheet: &SheetDataInner,
        style_table: &StyleTable,
    ) -> crate::core::Result<Vec<u8>> {
        let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

        w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))?;

        let mut root = BytesStart::new("worksheet");
        root.push_attribute(("xmlns", NS_SML));
        w.write_event(Event::Start(root))?;

        // Column widths
        if !sheet.col_widths.is_empty() {
            let mut sorted_cols: Vec<usize> = sheet.col_widths.keys().copied().collect();
            sorted_cols.sort_unstable();

            w.write_event(Event::Start(BytesStart::new("cols")))?;
            for col_idx in &sorted_cols {
                let width = sheet.col_widths[col_idx];
                let col_num = (*col_idx + 1).to_string();
                let width_str = format!("{width:.2}");

                let mut col_elem = BytesStart::new("col");
                col_elem.push_attribute(("min", col_num.as_str()));
                col_elem.push_attribute(("max", col_num.as_str()));
                col_elem.push_attribute(("width", width_str.as_str()));
                col_elem.push_attribute(("customWidth", "1"));
                w.write_event(Event::Empty(col_elem))?;
            }
            w.write_event(Event::End(BytesEnd::new("cols")))?;
        }

        w.write_event(Event::Start(BytesStart::new("sheetData")))?;

        for (row_idx, row) in sheet.rows.iter().enumerate() {
            if row.is_empty() || row.iter().all(|c| c.is_none()) {
                continue;
            }

            let row_num = (row_idx + 1).to_string();
            let mut row_elem = BytesStart::new("row");
            row_elem.push_attribute(("r", row_num.as_str()));
            w.write_event(Event::Start(row_elem))?;

            for (col_idx, cell) in row.iter().enumerate() {
                if let Some(stored) = cell {
                    let style_idx = style_table.get_idx(sheet, row_idx, col_idx);
                    Self::write_cell(&mut w, row_idx, col_idx, &stored.value, style_idx)?;
                }
            }

            w.write_event(Event::End(BytesEnd::new("row")))?;
        }

        w.write_event(Event::End(BytesEnd::new("sheetData")))?;
        w.write_event(Event::End(BytesEnd::new("worksheet")))?;

        Ok(w.into_inner())
    }

    fn write_cell(
        w: &mut Writer<Vec<u8>>,
        row: usize,
        col: usize,
        cell: &CellData,
        style_idx: Option<u32>,
    ) -> crate::core::Result<()> {
        let cell_ref = format!("{}{}", col_name(col as u32), row + 1);
        let s_attr = style_idx.map(|i| i.to_string());

        match cell {
            CellData::Empty => {},
            CellData::String(s) => {
                let mut c = BytesStart::new("c");
                c.push_attribute(("r", cell_ref.as_str()));
                c.push_attribute(("t", "inlineStr"));
                if let Some(ref s_val) = s_attr {
                    c.push_attribute(("s", s_val.as_str()));
                }
                w.write_event(Event::Start(c))?;
                w.write_event(Event::Start(BytesStart::new("is")))?;
                w.write_event(Event::Start(BytesStart::new("t")))?;
                w.write_event(Event::Text(BytesText::new(s)))?;
                w.write_event(Event::End(BytesEnd::new("t")))?;
                w.write_event(Event::End(BytesEnd::new("is")))?;
                w.write_event(Event::End(BytesEnd::new("c")))?;
            },
            CellData::Number(n) => {
                let mut c = BytesStart::new("c");
                c.push_attribute(("r", cell_ref.as_str()));
                if let Some(ref s_val) = s_attr {
                    c.push_attribute(("s", s_val.as_str()));
                }
                w.write_event(Event::Start(c))?;
                w.write_event(Event::Start(BytesStart::new("v")))?;
                w.write_event(Event::Text(BytesText::new(&format_number(*n))))?;
                w.write_event(Event::End(BytesEnd::new("v")))?;
                w.write_event(Event::End(BytesEnd::new("c")))?;
            },
            CellData::Boolean(b) => {
                let mut c = BytesStart::new("c");
                c.push_attribute(("r", cell_ref.as_str()));
                c.push_attribute(("t", "b"));
                if let Some(ref s_val) = s_attr {
                    c.push_attribute(("s", s_val.as_str()));
                }
                w.write_event(Event::Start(c))?;
                w.write_event(Event::Start(BytesStart::new("v")))?;
                w.write_event(Event::Text(BytesText::new(if *b { "1" } else { "0" })))?;
                w.write_event(Event::End(BytesEnd::new("v")))?;
                w.write_event(Event::End(BytesEnd::new("c")))?;
            },
            CellData::Formula(f) => {
                let mut c = BytesStart::new("c");
                c.push_attribute(("r", cell_ref.as_str()));
                if let Some(ref s_val) = s_attr {
                    c.push_attribute(("s", s_val.as_str()));
                }
                w.write_event(Event::Start(c))?;
                w.write_event(Event::Start(BytesStart::new("f")))?;
                w.write_event(Event::Text(BytesText::new(f)))?;
                w.write_event(Event::End(BytesEnd::new("f")))?;
                w.write_event(Event::End(BytesEnd::new("c")))?;
            },
        }

        Ok(())
    }
}

// ---------------------------------------------------------------------------
// StyleTable — collects unique CellStyle objects, assigns xfIds, builds
// styles.xml dynamically.
// ---------------------------------------------------------------------------

/// A unique font specification key (for deduplication).
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct FontKey {
    bold: bool,
    italic: bool,
    underline: bool,
    color: Option<String>,
    size_half_pt: Option<u32>, // size_pt * 2, to use as hash key
    name: Option<String>,
}

/// A unique fill specification key.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct FillKey(Option<String>); // background RGB hex

/// An xf (cell format) record.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct XfKey {
    font_idx: u32,
    fill_idx: u32,
    num_fmt_id: u32,
    h_align: Option<String>,
    wrap_text: bool,
}

struct StyleTable {
    /// Map from (sheet_ptr, row, col) to xf index.
    cell_xf: HashMap<(*const SheetDataInner, usize, usize), u32>,
    fonts: Vec<FontKey>,
    fills: Vec<FillKey>,
    num_fmts: Vec<(u32, String)>, // (numFmtId, formatCode) for custom formats
    xfs: Vec<XfKey>,
}

impl StyleTable {
    fn build(sheets: &[SheetDataInner]) -> Self {
        let mut table = StyleTable {
            cell_xf: HashMap::new(),
            fonts: Vec::new(),
            fills: Vec::new(),
            num_fmts: Vec::new(),
            xfs: Vec::new(),
        };

        // Built-in fill indices: 0=none, 1=gray125 (required by Excel)
        // We pre-populate to match the required structure.
        table.fills.push(FillKey(None)); // idx 0: none
        table.fills.push(FillKey(None)); // idx 1: gray125

        // Default font (idx 0)
        table.fonts.push(FontKey {
            bold: false,
            italic: false,
            underline: false,
            color: None,
            size_half_pt: None,
            name: None,
        });

        // Default xf (idx 0) — no style
        table.xfs.push(XfKey {
            font_idx: 0,
            fill_idx: 0,
            num_fmt_id: 0,
            h_align: None,
            wrap_text: false,
        });

        let mut next_custom_fmt_id: u32 = 164; // custom numFmtIds start at 164

        for sheet in sheets {
            let sheet_ptr = sheet as *const SheetDataInner;
            for ((row, col), style) in &sheet.cell_styles {
                // Resolve font index
                let font_key = FontKey {
                    bold: style.bold,
                    italic: style.italic,
                    underline: style.underline,
                    color: style.font_color.clone(),
                    size_half_pt: style.font_size_pt.map(|s| (s * 2.0).round() as u32),
                    name: style.font_name.clone(),
                };
                let font_idx = if font_key
                    == (FontKey {
                        bold: false,
                        italic: false,
                        underline: false,
                        color: None,
                        size_half_pt: None,
                        name: None,
                    }) {
                    0
                } else {
                    match table.fonts.iter().position(|f| f == &font_key) {
                        Some(i) => i as u32,
                        None => {
                            table.fonts.push(font_key);
                            (table.fonts.len() - 1) as u32
                        },
                    }
                };

                // Resolve fill index
                let fill_key = FillKey(style.background_color.clone());
                let fill_idx = if fill_key.0.is_none() {
                    0
                } else {
                    match table.fills.iter().position(|f| f == &fill_key) {
                        Some(i) => i as u32,
                        None => {
                            table.fills.push(fill_key);
                            (table.fills.len() - 1) as u32
                        },
                    }
                };

                // Resolve number format id
                let num_fmt_id = match style.number_format.builtin_id() {
                    Some(id) => id,
                    None => {
                        // Custom format — shouldn't happen with current enum
                        let id = next_custom_fmt_id;
                        next_custom_fmt_id += 1;
                        table.num_fmts.push((id, "General".to_string()));
                        id
                    },
                };

                let h_align_str = style.h_align.as_ref().map(|a| {
                    match a {
                        HAlign::Left => "left",
                        HAlign::Center => "center",
                        HAlign::Right => "right",
                    }
                    .to_string()
                });

                let xf_key = XfKey {
                    font_idx,
                    fill_idx,
                    num_fmt_id,
                    h_align: h_align_str,
                    wrap_text: style.wrap_text,
                };

                let xf_idx = match table.xfs.iter().position(|x| x == &xf_key) {
                    Some(i) => i as u32,
                    None => {
                        table.xfs.push(xf_key);
                        (table.xfs.len() - 1) as u32
                    },
                };

                table.cell_xf.insert((sheet_ptr, *row, *col), xf_idx);
            }
        }

        table
    }

    fn get_idx(
        &self,
        sheet: &SheetDataInner,
        row: usize,
        col: usize,
    ) -> Option<u32> {
        let key = (sheet as *const SheetDataInner, row, col);
        self.cell_xf.get(&key).copied().filter(|&i| i != 0)
    }

    fn build_styles_xml(&self) -> crate::core::Result<Vec<u8>> {
        let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

        w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))?;

        let mut root = BytesStart::new("styleSheet");
        root.push_attribute(("xmlns", NS_SML));
        w.write_event(Event::Start(root))?;

        // Custom number formats (if any)
        if !self.num_fmts.is_empty() {
            let count = self.num_fmts.len().to_string();
            let mut nf_root = BytesStart::new("numFmts");
            nf_root.push_attribute(("count", count.as_str()));
            w.write_event(Event::Start(nf_root))?;
            for (id, code) in &self.num_fmts {
                let mut nf = BytesStart::new("numFmt");
                nf.push_attribute(("numFmtId", id.to_string().as_str()));
                nf.push_attribute(("formatCode", code.as_str()));
                w.write_event(Event::Empty(nf))?;
            }
            w.write_event(Event::End(BytesEnd::new("numFmts")))?;
        }

        // fonts
        let font_count = self.fonts.len().to_string();
        let mut fonts_elem = BytesStart::new("fonts");
        fonts_elem.push_attribute(("count", font_count.as_str()));
        w.write_event(Event::Start(fonts_elem))?;
        for font in &self.fonts {
            w.write_event(Event::Start(BytesStart::new("font")))?;

            if font.bold {
                w.write_event(Event::Empty(BytesStart::new("b")))?;
            }
            if font.italic {
                w.write_event(Event::Empty(BytesStart::new("i")))?;
            }
            if font.underline {
                let mut u = BytesStart::new("u");
                u.push_attribute(("val", "single"));
                w.write_event(Event::Empty(u))?;
            }
            if let Some(ref color) = font.color {
                let mut c = BytesStart::new("color");
                c.push_attribute(("rgb", format!("FF{color}").as_str()));
                w.write_event(Event::Empty(c))?;
            }
            // size: default 11pt if not specified
            let size_half = font.size_half_pt.unwrap_or(22);
            let size_val = format!("{}", size_half / 2);
            let mut sz = BytesStart::new("sz");
            sz.push_attribute(("val", size_val.as_str()));
            w.write_event(Event::Empty(sz))?;

            let name_val = font.name.as_deref().unwrap_or("Calibri");
            let mut name = BytesStart::new("name");
            name.push_attribute(("val", name_val));
            w.write_event(Event::Empty(name))?;

            w.write_event(Event::End(BytesEnd::new("font")))?;
        }
        w.write_event(Event::End(BytesEnd::new("fonts")))?;

        // fills
        let fill_count = self.fills.len().to_string();
        let mut fills_elem = BytesStart::new("fills");
        fills_elem.push_attribute(("count", fill_count.as_str()));
        w.write_event(Event::Start(fills_elem))?;
        for (i, fill) in self.fills.iter().enumerate() {
            w.write_event(Event::Start(BytesStart::new("fill")))?;
            let pattern_type = if i == 1 {
                "gray125"
            } else if fill.0.is_some() {
                "solid"
            } else {
                "none"
            };
            let mut pf = BytesStart::new("patternFill");
            pf.push_attribute(("patternType", pattern_type));
            if let Some(ref color) = fill.0 {
                w.write_event(Event::Start(pf))?;
                let mut fg = BytesStart::new("fgColor");
                fg.push_attribute(("rgb", format!("FF{color}").as_str()));
                w.write_event(Event::Empty(fg))?;
                w.write_event(Event::End(BytesEnd::new("patternFill")))?;
            } else {
                w.write_event(Event::Empty(pf))?;
            }
            w.write_event(Event::End(BytesEnd::new("fill")))?;
        }
        w.write_event(Event::End(BytesEnd::new("fills")))?;

        // borders (minimal: one empty border)
        w.write_event(Event::Start({
            let mut e = BytesStart::new("borders");
            e.push_attribute(("count", "1"));
            e
        }))?;
        w.write_event(Event::Start(BytesStart::new("border")))?;
        for edge in ["left", "right", "top", "bottom"] {
            w.write_event(Event::Empty(BytesStart::new(edge)))?;
        }
        w.write_event(Event::End(BytesEnd::new("border")))?;
        w.write_event(Event::End(BytesEnd::new("borders")))?;

        // cellStyleXfs (base xf table required by Excel)
        w.write_event(Event::Start({
            let mut e = BytesStart::new("cellStyleXfs");
            e.push_attribute(("count", "1"));
            e
        }))?;
        w.write_event(Event::Empty(BytesStart::new("xf")))?;
        w.write_event(Event::End(BytesEnd::new("cellStyleXfs")))?;

        // cellXfs
        let xf_count = self.xfs.len().to_string();
        let mut cxf = BytesStart::new("cellXfs");
        cxf.push_attribute(("count", xf_count.as_str()));
        w.write_event(Event::Start(cxf))?;

        for xf in &self.xfs {
            let apply_font = (xf.font_idx != 0).to_string();
            let apply_fill = (xf.fill_idx != 0).to_string();
            let apply_num_fmt = (xf.num_fmt_id != 0).to_string();
            let apply_align = xf.h_align.is_some() || xf.wrap_text;

            let mut xf_elem = BytesStart::new("xf");
            xf_elem.push_attribute(("xfId", "0"));
            xf_elem.push_attribute(("fontId", xf.font_idx.to_string().as_str()));
            xf_elem.push_attribute(("fillId", xf.fill_idx.to_string().as_str()));
            xf_elem.push_attribute(("borderId", "0"));
            xf_elem.push_attribute(("numFmtId", xf.num_fmt_id.to_string().as_str()));
            if xf.font_idx != 0 {
                xf_elem.push_attribute(("applyFont", apply_font.as_str()));
            }
            if xf.fill_idx != 0 {
                xf_elem.push_attribute(("applyFill", apply_fill.as_str()));
            }
            if xf.num_fmt_id != 0 {
                xf_elem.push_attribute(("applyNumberFormat", apply_num_fmt.as_str()));
            }
            if apply_align {
                xf_elem.push_attribute(("applyAlignment", "true"));
            }

            if apply_align {
                w.write_event(Event::Start(xf_elem))?;
                let mut align = BytesStart::new("alignment");
                if let Some(ref ha) = xf.h_align {
                    align.push_attribute(("horizontal", ha.as_str()));
                }
                if xf.wrap_text {
                    align.push_attribute(("wrapText", "1"));
                }
                w.write_event(Event::Empty(align))?;
                w.write_event(Event::End(BytesEnd::new("xf")))?;
            } else {
                w.write_event(Event::Empty(xf_elem))?;
            }
        }

        w.write_event(Event::End(BytesEnd::new("cellXfs")))?;
        w.write_event(Event::End(BytesEnd::new("styleSheet")))?;

        Ok(w.into_inner())
    }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn col_name(col: u32) -> String {
    let mut result = Vec::new();
    let mut n = col + 1;
    while n > 0 {
        n -= 1;
        result.push(b'A' + (n % 26) as u8);
        n /= 26;
    }
    result.reverse();
    String::from_utf8(result).unwrap()
}

fn format_number(n: f64) -> String {
    if n.fract() == 0.0 && n.abs() < (i64::MAX as f64) {
        format!("{}", n as i64)
    } else {
        format!("{n}")
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn col_name_basic() {
        assert_eq!(col_name(0), "A");
        assert_eq!(col_name(25), "Z");
        assert_eq!(col_name(26), "AA");
        assert_eq!(col_name(701), "ZZ");
    }

    #[test]
    fn formula_cell_roundtrip() {
        let mut wb = XlsxWriter::new();
        let mut sheet = wb.add_sheet("Test");
        sheet.set_cell(0, 0, CellData::Number(10.0));
        sheet.set_cell(1, 0, CellData::Number(20.0));
        sheet.set_cell(2, 0, CellData::Formula("SUM(A1:A2)".into()));

        let mut buf = std::io::Cursor::new(Vec::new());
        wb.write_to(&mut buf).expect("write xlsx");
        assert!(!buf.get_ref().is_empty());
    }

    #[test]
    fn styled_cells() {
        let mut wb = XlsxWriter::new();
        let mut sheet = wb.add_sheet("Styled");
        sheet.set_cell_styled(
            0,
            0,
            CellData::String("Header".into()),
            CellStyle::new().bold().background("4472C4").font_color("FFFFFF"),
        );
        sheet.set_cell_styled(
            1,
            0,
            CellData::Formula("SUM(A3:A10)".into()),
            CellStyle::new().bold().number_format(NumberFormat::Currency),
        );
        sheet.set_column_width(0, 20.0);

        let mut buf = std::io::Cursor::new(Vec::new());
        wb.write_to(&mut buf).expect("write xlsx");
        assert!(!buf.get_ref().is_empty());
    }

    #[test]
    fn all_number_formats() {
        let mut wb = XlsxWriter::new();
        let mut sheet = wb.add_sheet("Fmts");
        let formats = [
            NumberFormat::General,
            NumberFormat::Integer,
            NumberFormat::Decimal2,
            NumberFormat::Currency,
            NumberFormat::Percent,
            NumberFormat::Percent2,
            NumberFormat::Date,
            NumberFormat::DateTime,
        ];
        for (i, fmt) in formats.iter().enumerate() {
            sheet.set_cell_styled(
                i,
                0,
                CellData::Number(42.0),
                CellStyle::new().number_format(fmt.clone()),
            );
        }
        let mut buf = std::io::Cursor::new(Vec::new());
        wb.write_to(&mut buf).expect("write xlsx");
        assert!(!buf.get_ref().is_empty());
    }
}
