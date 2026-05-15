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
            Self::Currency => Some(7), // #,##0.00
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
    /// Embedded font programs to ship inside the package under `xl/fonts/`.
    /// Same layout as DOCX `word/fonts/` and PPTX `ppt/fonts/`. Excel
    /// itself doesn't honor these without `<workbookView>` / theme
    /// plumbing, but the in-process reader scans the directory so
    /// PDF↔XLSX round-trips can preserve typefaces.
    embedded_fonts: Vec<(String, Vec<u8>)>,
    /// Document metadata for `docProps/core.xml`. `None` means no
    /// core-properties part is written.
    metadata: Option<crate::ir::Metadata>,
}

/// Per-worksheet page geometry.
///
/// Maps roughly 1-to-1 onto OOXML's `<pageMargins>` and `<pageSetup>` —
/// margins are stored in inches per ECMA-376 (§18.3.1.62), the page size
/// is emitted as `paperWidth`/`paperHeight` in millimetres so arbitrary
/// PDF MediaBox dimensions round-trip without snapping to the nearest
/// `paperSize` enum.  All inputs are twips for parity with the rest of
/// the IR (`width_twips`, `margin_top_twips`, …).
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PageSetup {
    /// Page width in twips (1/1440 inch).
    pub width_twips: u32,
    /// Page height in twips.
    pub height_twips: u32,
    /// Top margin in twips.
    pub margin_top_twips: u32,
    /// Bottom margin in twips.
    pub margin_bottom_twips: u32,
    /// Left margin in twips.
    pub margin_left_twips: u32,
    /// Right margin in twips.
    pub margin_right_twips: u32,
    /// Header distance from top edge in twips.
    pub header_distance_twips: u32,
    /// Footer distance from bottom edge in twips.
    pub footer_distance_twips: u32,
    /// Whether the page is in landscape orientation.
    pub landscape: bool,
}

/// A picture anchored on a worksheet via a DrawingML drawing part.
///
/// Anchor coordinates are in EMU and absolute relative to the sheet
/// origin (top-left). Round-trips render via `<xdr:absoluteAnchor>` in
/// `xl/drawings/drawingN.xml`. The writer emits the bytes verbatim; the
/// reader resolves them back through the worksheet → drawing → image
/// relationship chain.
#[derive(Debug, Clone)]
pub struct SheetImage {
    /// Raw image bytes (PNG / JPEG / etc., as produced by the source).
    pub data: Vec<u8>,
    /// Lowercase file extension (`"png"`, `"jpeg"`, ...).
    pub format: String,
    /// X anchor in EMU, from sheet origin.
    pub x_emu: i64,
    /// Y anchor in EMU.
    pub y_emu: i64,
    /// Rendered width in EMU.
    pub cx_emu: i64,
    /// Rendered height in EMU.
    pub cy_emu: i64,
}

/// A text shape anchored on a worksheet via a DrawingML drawing part.
///
/// Used by the layout-preserving PDF→XLSX path to emit each PDF text
/// span at its exact source EMU coordinates as an `<xdr:sp>` shape
/// inside `xl/drawings/drawingN.xml`. The shape carries a single run
/// with the span's text, font, size, weight, italic, and colour.
#[derive(Debug, Clone)]
pub struct SheetTextShape {
    /// Text content of the shape (single run).
    pub text: String,
    /// Font face name (e.g. `"Times New Roman"`).
    pub font_name: String,
    /// Font size in points (full-pt scale, not half-pt).
    pub font_size_pt: f32,
    /// Bold weight.
    pub bold: bool,
    /// Italic style.
    pub italic: bool,
    /// Optional 6-char hex colour like `"FF0000"`. `None` ⇒ pure black.
    pub color_hex: Option<String>,
    /// X anchor in EMU.
    pub x_emu: i64,
    /// Y anchor in EMU.
    pub y_emu: i64,
    /// Width in EMU.
    pub cx_emu: i64,
    /// Height in EMU.
    pub cy_emu: i64,
}

/// Full internal sheet representation.
struct SheetDataInner {
    pub name: String,
    pub rows: Vec<Vec<Option<StoredCellInner>>>,
    pub col_widths: HashMap<usize, f64>,
    pub cell_styles: HashMap<(usize, usize), CellStyle>,
    /// Merged cell regions: (row, col, row_span, col_span).
    pub merge_regions: Vec<(usize, usize, usize, usize)>,
    /// Per-sheet page geometry (`<pageMargins>` + `<pageSetup>`).
    pub page_setup: Option<PageSetup>,
    /// Pictures anchored on this sheet via a DrawingML drawing part.
    pub images: Vec<SheetImage>,
    /// Text shapes anchored on this sheet via a DrawingML drawing part.
    /// Used by the layout-preserving PDF→XLSX path.
    pub text_shapes: Vec<SheetTextShape>,
}

impl SheetDataInner {
    fn new(name: &str) -> Self {
        Self {
            name: name.to_string(),
            rows: Vec::new(),
            col_widths: HashMap::new(),
            cell_styles: HashMap::new(),
            merge_regions: Vec::new(),
            page_setup: None,
            images: Vec::new(),
            text_shapes: Vec::new(),
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

    pub fn merge_cells(
        &mut self,
        row: usize,
        col: usize,
        row_span: usize,
        col_span: usize,
    ) -> &mut Self {
        if row_span == 0 || col_span == 0 {
            return self;
        }
        if row_span > 1 || col_span > 1 {
            self.merge_regions.push((row, col, row_span, col_span));
        }
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

    /// Merge a rectangular range of cells.
    ///
    /// The top-left cell retains its value; the rest are blanked by the reader.
    /// A span of 1×1 is a no-op.
    pub fn merge_cells(
        &mut self,
        row: usize,
        col: usize,
        row_span: usize,
        col_span: usize,
    ) -> &mut Self {
        self.0.merge_cells(row, col, row_span, col_span);
        self
    }

    /// Set per-sheet page geometry. Emits `<pageMargins>` and `<pageSetup>`
    /// inside the worksheet XML so PDF→XLSX→PDF round-trips preserve the
    /// source MediaBox and margins instead of snapping back to default
    /// Letter-portrait. Pass `None` (the default) to omit both elements.
    pub fn set_page_setup(&mut self, ps: PageSetup) -> &mut Self {
        self.0.page_setup = Some(ps);
        self
    }

    /// Anchor a styled text run on this worksheet at absolute EMU
    /// coordinates. Used by the PDF→XLSX layout-preserving path: each
    /// PDF text span becomes one `<xdr:sp>` shape with a single run.
    #[allow(clippy::too_many_arguments)]
    pub fn add_text_shape(
        &mut self,
        text: impl Into<String>,
        font_name: impl Into<String>,
        font_size_pt: f32,
        bold: bool,
        italic: bool,
        color_hex: Option<String>,
        x_emu: i64,
        y_emu: i64,
        cx_emu: i64,
        cy_emu: i64,
    ) -> &mut Self {
        self.0.text_shapes.push(SheetTextShape {
            text: text.into(),
            font_name: font_name.into(),
            font_size_pt,
            bold,
            italic,
            color_hex,
            x_emu,
            y_emu,
            cx_emu,
            cy_emu,
        });
        self
    }

    /// Anchor a picture on this worksheet at absolute EMU coordinates.
    ///
    /// On write the writer materialises a `xl/drawings/drawingN.xml`
    /// part for this sheet, registers an IMAGE relationship per
    /// picture, and writes the bytes under `xl/media/image_<sheet>_<n>.<ext>`.
    /// `format` is the lowercase file extension (`"png"`, `"jpeg"`, ...).
    pub fn add_image(
        &mut self,
        data: Vec<u8>,
        format: impl Into<String>,
        x_emu: i64,
        y_emu: i64,
        cx_emu: i64,
        cy_emu: i64,
    ) -> &mut Self {
        self.0.images.push(SheetImage {
            data,
            format: format.into(),
            x_emu,
            y_emu,
            cx_emu,
            cy_emu,
        });
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
        Self {
            sheets: Vec::new(),
            embedded_fonts: Vec::new(),
            metadata: None,
        }
    }

    /// Set document metadata (written to `docProps/core.xml`).
    pub fn set_metadata(&mut self, meta: &crate::ir::Metadata) -> &mut Self {
        self.metadata = Some(meta.clone());
        self
    }

    /// Embed a font program (TrueType / OpenType bytes) under `xl/fonts/`.
    /// `name` is used for the file name and as the human-readable font name.
    /// Subsequent calls with the same name are deduplicated.
    pub fn embed_font(&mut self, name: impl Into<String>, data: Vec<u8>) -> &mut Self {
        let name = name.into();
        if !self.embedded_fonts.iter().any(|(n, _)| n == &name) {
            self.embedded_fonts.push((name, data));
        }
        self
    }

    /// Add a worksheet and return a mutable handle to it.
    pub fn add_sheet(&mut self, name: &str) -> SheetData<'_> {
        self.sheets.push(SheetDataInner::new(name));
        SheetData(self.sheets.last_mut().unwrap())
    }

    /// Add a sheet and return its 0-based index (for use with index-based API).
    pub fn add_sheet_get_index(&mut self, name: &str) -> usize {
        self.sheets.push(SheetDataInner::new(name));
        self.sheets.len() - 1
    }

    /// Set a cell value by sheet index.
    pub fn sheet_set_cell(&mut self, sheet: usize, row: usize, col: usize, value: CellData) {
        if let Some(s) = self.sheets.get_mut(sheet) {
            s.set_cell(row, col, value);
        }
    }

    /// Set a cell value with styling by sheet index.
    pub fn sheet_set_cell_styled(
        &mut self,
        sheet: usize,
        row: usize,
        col: usize,
        value: CellData,
        style: CellStyle,
    ) {
        if let Some(s) = self.sheets.get_mut(sheet) {
            s.set_cell_styled(row, col, value, style);
        }
    }

    /// Merge a rectangular range of cells by sheet index.
    pub fn sheet_merge_cells(
        &mut self,
        sheet: usize,
        row: usize,
        col: usize,
        row_span: usize,
        col_span: usize,
    ) {
        if let Some(s) = self.sheets.get_mut(sheet) {
            s.merge_cells(row, col, row_span, col_span);
        }
    }

    /// Set column width by sheet index.
    pub fn sheet_set_column_width(&mut self, sheet: usize, col: usize, width: f64) {
        if let Some(s) = self.sheets.get_mut(sheet) {
            s.set_column_width(col, width);
        }
    }

    /// Set per-sheet page geometry by sheet index. See `SheetData::set_page_setup`.
    pub fn sheet_set_page_setup(&mut self, sheet: usize, ps: PageSetup) {
        if let Some(s) = self.sheets.get_mut(sheet) {
            s.page_setup = Some(ps);
        }
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

        // Core properties (docProps/core.xml). Optional; written only
        // when caller supplied metadata via `set_metadata`. Surfaces
        // PDF /Title /Author etc. in Excel's "Properties" dialog after
        // a PDF→XLSX→Excel round trip.
        if let Some(ref meta) = self.metadata {
            let core_part = PartName::new("/docProps/core.xml")?;
            opc.add_package_rel(rel_types::CORE_PROPERTIES, "docProps/core.xml");
            let core_xml = crate::core::core_properties::generate_xml(meta);
            opc.add_part(&core_part, crate::core::core_properties::CONTENT_TYPE, &core_xml)?;
        }

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

            // Emit drawing + media parts up-front so we have the rId
            // for the `<drawing r:id="…"/>` element inside the
            // worksheet XML below. Sheets without pictures or text
            // shapes get no drawing part at all.
            let drawing_rid = if !sheet.images.is_empty() || !sheet.text_shapes.is_empty() {
                Some(Self::write_drawing_for_sheet(
                    opc,
                    &part_name,
                    i + 1,
                    &sheet.images,
                    &sheet.text_shapes,
                )?)
            } else {
                None
            };

            let ws_xml = Self::build_worksheet_xml(sheet, &style_table, drawing_rid.as_deref())?;
            opc.add_part(&part_name, CT_WORKSHEET, &ws_xml)?;
        }

        let styles_part = PartName::new("/xl/styles.xml")?;
        let styles_xml = style_table.build_styles_xml()?;
        opc.add_part(&styles_part, CT_STYLES, &styles_xml)?;

        // Embed fonts under `xl/fonts/font_<n>_<safe_name>.ttf`. Same
        // layout as DOCX/PPTX. Excel itself doesn't auto-discover the
        // fonts without `<workbookView>` plumbing, but the in-process
        // reader scans the directory so PDF↔XLSX round-trips can reuse
        // the source typeface.
        crate::core::embedded_fonts::write_embedded_fonts(opc, "/xl/fonts/", &self.embedded_fonts)?;

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
        drawing_rid: Option<&str>,
    ) -> crate::core::Result<Vec<u8>> {
        let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

        w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))?;

        let mut root = BytesStart::new("worksheet");
        root.push_attribute(("xmlns", NS_SML));
        // Worksheets that anchor drawings need the relationship
        // namespace so the `<drawing r:id="…"/>` element below
        // resolves. Declaring it unconditionally is harmless for
        // plain-data sheets and keeps the writer code simple.
        root.push_attribute(("xmlns:r", NS_REL));
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

        if !sheet.merge_regions.is_empty() {
            let count_str = sheet.merge_regions.len().to_string();
            let mut mc_elem = BytesStart::new("mergeCells");
            mc_elem.push_attribute(("count", count_str.as_str()));
            w.write_event(Event::Start(mc_elem))?;
            for &(row, col, row_span, col_span) in &sheet.merge_regions {
                let tl = format!("{}{}", col_name(col as u32), row + 1);
                let br = format!("{}{}", col_name((col + col_span - 1) as u32), row + row_span);
                let ref_str = format!("{tl}:{br}");
                let mut mc = BytesStart::new("mergeCell");
                mc.push_attribute(("ref", ref_str.as_str()));
                w.write_event(Event::Empty(mc))?;
            }
            w.write_event(Event::End(BytesEnd::new("mergeCells")))?;
        }

        // <pageMargins> + <pageSetup>. ECMA-376 §18.3.1.62 / §18.3.1.63 —
        // pageMargins values are in inches (f64), pageSetup carries the
        // physical paper dimensions and orientation. We emit `paperWidth`
        // and `paperHeight` in mm so arbitrary PDF MediaBoxes round-trip
        // verbatim instead of snapping to the closest `paperSize` enum
        // (which only covers a fixed set of standard sizes — Letter,
        // Legal, A4, A3, …).
        if let Some(ps) = sheet.page_setup {
            // twips → inches, twips → mm (1 inch = 1440 twips = 25.4 mm).
            let to_in = |t: u32| t as f64 / 1440.0;
            let to_mm = |t: u32| t as f64 / 1440.0 * 25.4;

            let left = format!("{:.4}", to_in(ps.margin_left_twips));
            let right = format!("{:.4}", to_in(ps.margin_right_twips));
            let top = format!("{:.4}", to_in(ps.margin_top_twips));
            let bottom = format!("{:.4}", to_in(ps.margin_bottom_twips));
            let header = format!("{:.4}", to_in(ps.header_distance_twips));
            let footer = format!("{:.4}", to_in(ps.footer_distance_twips));

            let mut pm = BytesStart::new("pageMargins");
            pm.push_attribute(("left", left.as_str()));
            pm.push_attribute(("right", right.as_str()));
            pm.push_attribute(("top", top.as_str()));
            pm.push_attribute(("bottom", bottom.as_str()));
            pm.push_attribute(("header", header.as_str()));
            pm.push_attribute(("footer", footer.as_str()));
            w.write_event(Event::Empty(pm))?;

            let pw_mm = format!("{:.2}mm", to_mm(ps.width_twips));
            let ph_mm = format!("{:.2}mm", to_mm(ps.height_twips));
            let orientation = if ps.landscape {
                "landscape"
            } else {
                "portrait"
            };
            let mut psu = BytesStart::new("pageSetup");
            psu.push_attribute(("paperWidth", pw_mm.as_str()));
            psu.push_attribute(("paperHeight", ph_mm.as_str()));
            psu.push_attribute(("orientation", orientation));
            w.write_event(Event::Empty(psu))?;
        }

        // `<drawing>` MUST appear after `<pageSetup>` per the
        // worksheet child-order schema (CT_Worksheet, ECMA-376
        // §18.3.1.99). Excel rejects the file with "We found a problem
        // with some content" otherwise.
        if let Some(rid) = drawing_rid {
            let mut d = BytesStart::new("drawing");
            d.push_attribute(("r:id", rid));
            w.write_event(Event::Empty(d))?;
        }

        w.write_event(Event::End(BytesEnd::new("worksheet")))?;

        Ok(w.into_inner())
    }

    /// Materialise `xl/drawings/drawing<sheet_n>.xml`, write each
    /// picture's bytes under `xl/media/image_<sheet_n>_<pic_n>.<ext>`,
    /// wire the worksheet→drawing and drawing→image relationships, and
    /// register PNG/JPEG default content types.
    ///
    /// Returns the relationship ID added to the worksheet — the caller
    /// places it on the `<drawing r:id="…"/>` element inside the
    /// worksheet XML.
    fn write_drawing_for_sheet<W: Write + Seek>(
        opc: &mut OpcWriter<W>,
        worksheet_part: &PartName,
        sheet_n: usize,
        images: &[SheetImage],
        text_shapes: &[SheetTextShape],
    ) -> Result<String> {
        let drawing_target = format!("../drawings/drawing{}.xml", sheet_n);
        let drawing_rid = opc.add_part_rel(worksheet_part, rel_types::DRAWING, &drawing_target);

        let drawing_part_str = format!("/xl/drawings/drawing{}.xml", sheet_n);
        let drawing_part = PartName::new(&drawing_part_str)?;

        // Add IMAGE rels off the drawing part. Targets are relative to
        // the drawing part itself (`../media/imageX.ext`). Track the
        // rIds so each `<xdr:pic>` in the drawing XML can reference
        // them via `<a:blip r:embed="rIdN"/>`.
        let mut blip_rids: Vec<String> = Vec::with_capacity(images.len());
        for (i, img) in images.iter().enumerate() {
            let ext = if img.format.is_empty() {
                "png"
            } else {
                img.format.as_str()
            };
            let media_path_str = format!("/xl/media/image_{}_{}.{}", sheet_n, i + 1, ext);
            let media_part = PartName::new(&media_path_str)?;

            // Default Content-Type by extension (Default Extension="png")
            // satisfies SDK validators that flag overrides without a
            // matching Default. Re-registering the same default is a
            // no-op inside ContentTypesBuilder.
            let mime = match ext {
                "jpg" | "jpeg" => "image/jpeg",
                "gif" => "image/gif",
                "tiff" | "tif" => "image/tiff",
                "bmp" => "image/bmp",
                "emf" => "image/x-emf",
                "wmf" => "image/x-wmf",
                _ => "image/png",
            };
            opc.register_default_content_type(ext, mime);

            // Write image bytes raw (no Content-Type override needed
            // since we registered the Default above; passing the same
            // mime to add_part is harmless).
            opc.add_part(&media_part, mime, &img.data)?;

            // Drawing-relative target: `../media/image_..._N.ext`.
            let rel_target = format!("../media/image_{}_{}.{}", sheet_n, i + 1, ext);
            let rid = opc.add_part_rel(&drawing_part, rel_types::IMAGE, &rel_target);
            blip_rids.push(rid);
        }

        // Now the drawing XML itself. One `<xdr:absoluteAnchor>` per
        // picture and per text shape; anchor in EMU from the sheet
        // origin, with the picture's `<a:blip r:embed="rIdN"/>`
        // referring back to the image rels we just added.
        let drawing_xml = build_drawing_xml(images, &blip_rids, text_shapes)?;
        const CT_DRAWING: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
        opc.add_part(&drawing_part, CT_DRAWING, &drawing_xml)?;

        Ok(drawing_rid)
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

/// Generate `xl/drawings/drawing<n>.xml` for a sheet's pictures.
///
/// Each picture becomes one `<xdr:absoluteAnchor>` containing an
/// `<xdr:pic>` with an `<a:blip r:embed="…"/>` referring back to the
/// IMAGE rel registered on the drawing part. EMU coordinates flow
/// through verbatim from the caller's `SheetImage`, which preserves
/// source-PDF anchor positions on a PDF→XLSX→PDF round-trip when the
/// upstream IR carries them.
fn build_drawing_xml(
    images: &[SheetImage],
    blip_rids: &[String],
    text_shapes: &[SheetTextShape],
) -> crate::core::Result<Vec<u8>> {
    const NS_XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    const NS_A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);
    w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))?;

    let mut root = BytesStart::new("xdr:wsDr");
    root.push_attribute(("xmlns:xdr", NS_XDR));
    root.push_attribute(("xmlns:a", NS_A));
    root.push_attribute(("xmlns:r", NS_REL));
    w.write_event(Event::Start(root))?;

    for (i, img) in images.iter().enumerate() {
        let rid = blip_rids.get(i).map(String::as_str).unwrap_or("rId1");

        // <xdr:absoluteAnchor>
        w.write_event(Event::Start(BytesStart::new("xdr:absoluteAnchor")))?;

        // <xdr:pos x=".." y=".."/>
        let pos_x = img.x_emu.to_string();
        let pos_y = img.y_emu.to_string();
        let mut pos = BytesStart::new("xdr:pos");
        pos.push_attribute(("x", pos_x.as_str()));
        pos.push_attribute(("y", pos_y.as_str()));
        w.write_event(Event::Empty(pos))?;

        // <xdr:ext cx=".." cy=".."/>
        let ext_cx = img.cx_emu.max(1).to_string();
        let ext_cy = img.cy_emu.max(1).to_string();
        let mut ext = BytesStart::new("xdr:ext");
        ext.push_attribute(("cx", ext_cx.as_str()));
        ext.push_attribute(("cy", ext_cy.as_str()));
        w.write_event(Event::Empty(ext))?;

        // <xdr:pic>
        w.write_event(Event::Start(BytesStart::new("xdr:pic")))?;

        // <xdr:nvPicPr>
        w.write_event(Event::Start(BytesStart::new("xdr:nvPicPr")))?;
        let pic_id = (i + 1).to_string();
        let pic_name = format!("Picture {}", i + 1);
        let mut cnv_pr = BytesStart::new("xdr:cNvPr");
        cnv_pr.push_attribute(("id", pic_id.as_str()));
        cnv_pr.push_attribute(("name", pic_name.as_str()));
        w.write_event(Event::Empty(cnv_pr))?;
        w.write_event(Event::Empty(BytesStart::new("xdr:cNvPicPr")))?;
        w.write_event(Event::End(BytesEnd::new("xdr:nvPicPr")))?;

        // <xdr:blipFill>
        w.write_event(Event::Start(BytesStart::new("xdr:blipFill")))?;
        let mut blip = BytesStart::new("a:blip");
        blip.push_attribute(("r:embed", rid));
        w.write_event(Event::Empty(blip))?;
        w.write_event(Event::Start(BytesStart::new("a:stretch")))?;
        w.write_event(Event::Empty(BytesStart::new("a:fillRect")))?;
        w.write_event(Event::End(BytesEnd::new("a:stretch")))?;
        w.write_event(Event::End(BytesEnd::new("xdr:blipFill")))?;

        // <xdr:spPr>
        w.write_event(Event::Start(BytesStart::new("xdr:spPr")))?;
        w.write_event(Event::Start(BytesStart::new("a:xfrm")))?;
        let mut off = BytesStart::new("a:off");
        off.push_attribute(("x", pos_x.as_str()));
        off.push_attribute(("y", pos_y.as_str()));
        w.write_event(Event::Empty(off))?;
        let mut ext2 = BytesStart::new("a:ext");
        ext2.push_attribute(("cx", ext_cx.as_str()));
        ext2.push_attribute(("cy", ext_cy.as_str()));
        w.write_event(Event::Empty(ext2))?;
        w.write_event(Event::End(BytesEnd::new("a:xfrm")))?;
        let mut prst = BytesStart::new("a:prstGeom");
        prst.push_attribute(("prst", "rect"));
        w.write_event(Event::Start(prst))?;
        w.write_event(Event::Empty(BytesStart::new("a:avLst")))?;
        w.write_event(Event::End(BytesEnd::new("a:prstGeom")))?;
        w.write_event(Event::End(BytesEnd::new("xdr:spPr")))?;

        w.write_event(Event::End(BytesEnd::new("xdr:pic")))?;

        // <xdr:clientData/>
        w.write_event(Event::Empty(BytesStart::new("xdr:clientData")))?;

        w.write_event(Event::End(BytesEnd::new("xdr:absoluteAnchor")))?;
    }

    // ── Text shapes (one `<xdr:sp>` per layout-mode PDF span) ───────────
    let pic_count = images.len();
    for (j, ts) in text_shapes.iter().enumerate() {
        // Skip empty-text shapes — Excel rejects shape XML with
        // an empty `<a:t/>` even though OOXML allows it.
        let trimmed = ts.text.trim_matches('\u{0000}');
        if trimmed.is_empty() {
            continue;
        }

        w.write_event(Event::Start(BytesStart::new("xdr:absoluteAnchor")))?;

        let pos_x = ts.x_emu.to_string();
        let pos_y = ts.y_emu.to_string();
        let mut pos = BytesStart::new("xdr:pos");
        pos.push_attribute(("x", pos_x.as_str()));
        pos.push_attribute(("y", pos_y.as_str()));
        w.write_event(Event::Empty(pos))?;

        let ext_cx = ts.cx_emu.max(1).to_string();
        let ext_cy = ts.cy_emu.max(1).to_string();
        let mut ext = BytesStart::new("xdr:ext");
        ext.push_attribute(("cx", ext_cx.as_str()));
        ext.push_attribute(("cy", ext_cy.as_str()));
        w.write_event(Event::Empty(ext))?;

        w.write_event(Event::Start(BytesStart::new("xdr:sp")))?;

        // <xdr:nvSpPr>
        w.write_event(Event::Start(BytesStart::new("xdr:nvSpPr")))?;
        let sp_id = (pic_count + j + 1).to_string();
        let sp_name = format!("TextShape {}", pic_count + j + 1);
        let mut cnv_pr = BytesStart::new("xdr:cNvPr");
        cnv_pr.push_attribute(("id", sp_id.as_str()));
        cnv_pr.push_attribute(("name", sp_name.as_str()));
        w.write_event(Event::Empty(cnv_pr))?;
        let mut cnv_sp_pr = BytesStart::new("xdr:cNvSpPr");
        cnv_sp_pr.push_attribute(("txBox", "1"));
        w.write_event(Event::Empty(cnv_sp_pr))?;
        w.write_event(Event::End(BytesEnd::new("xdr:nvSpPr")))?;

        // <xdr:spPr>
        w.write_event(Event::Start(BytesStart::new("xdr:spPr")))?;
        w.write_event(Event::Start(BytesStart::new("a:xfrm")))?;
        let mut off = BytesStart::new("a:off");
        off.push_attribute(("x", pos_x.as_str()));
        off.push_attribute(("y", pos_y.as_str()));
        w.write_event(Event::Empty(off))?;
        let mut ext2 = BytesStart::new("a:ext");
        ext2.push_attribute(("cx", ext_cx.as_str()));
        ext2.push_attribute(("cy", ext_cy.as_str()));
        w.write_event(Event::Empty(ext2))?;
        w.write_event(Event::End(BytesEnd::new("a:xfrm")))?;
        let mut prst = BytesStart::new("a:prstGeom");
        prst.push_attribute(("prst", "rect"));
        w.write_event(Event::Start(prst))?;
        w.write_event(Event::Empty(BytesStart::new("a:avLst")))?;
        w.write_event(Event::End(BytesEnd::new("a:prstGeom")))?;
        // Transparent fill so the text shape doesn't paint a
        // white rectangle over neighbouring content.
        w.write_event(Event::Empty(BytesStart::new("a:noFill")))?;
        w.write_event(Event::End(BytesEnd::new("xdr:spPr")))?;

        // <xdr:txBody> — inline a single run with the span's run
        // properties. PPTX/PRESENT and SpreadsheetML share the same
        // DrawingML run model, so the structure mirrors PPTX text
        // bodies elsewhere in this crate.
        w.write_event(Event::Start(BytesStart::new("xdr:txBody")))?;
        // <a:bodyPr wrap="none"> so a single span doesn't wrap mid-line.
        let mut body_pr = BytesStart::new("a:bodyPr");
        body_pr.push_attribute(("wrap", "none"));
        body_pr.push_attribute(("rtlCol", "0"));
        body_pr.push_attribute(("lIns", "0"));
        body_pr.push_attribute(("tIns", "0"));
        body_pr.push_attribute(("rIns", "0"));
        body_pr.push_attribute(("bIns", "0"));
        w.write_event(Event::Empty(body_pr))?;
        w.write_event(Event::Empty(BytesStart::new("a:lstStyle")))?;
        w.write_event(Event::Start(BytesStart::new("a:p")))?;
        // <a:pPr marL="0" indent="0"/>
        let mut p_pr = BytesStart::new("a:pPr");
        p_pr.push_attribute(("marL", "0"));
        p_pr.push_attribute(("indent", "0"));
        w.write_event(Event::Empty(p_pr))?;
        // <a:r>
        w.write_event(Event::Start(BytesStart::new("a:r")))?;
        // <a:rPr lang="en-US" sz=".." b=".." i=".."> with optional <a:solidFill> and <a:latin>
        let sz_hp = (ts.font_size_pt * 100.0).round() as i32;
        let sz_str = sz_hp.to_string();
        let mut r_pr = BytesStart::new("a:rPr");
        r_pr.push_attribute(("lang", "en-US"));
        r_pr.push_attribute(("sz", sz_str.as_str()));
        if ts.bold {
            r_pr.push_attribute(("b", "1"));
        }
        if ts.italic {
            r_pr.push_attribute(("i", "1"));
        }
        let want_color_or_font = ts.color_hex.is_some() || !ts.font_name.is_empty();
        if want_color_or_font {
            w.write_event(Event::Start(r_pr))?;
            if let Some(ref hex) = ts.color_hex {
                w.write_event(Event::Start(BytesStart::new("a:solidFill")))?;
                let mut srgb = BytesStart::new("a:srgbClr");
                srgb.push_attribute(("val", hex.as_str()));
                w.write_event(Event::Empty(srgb))?;
                w.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
            }
            if !ts.font_name.is_empty() {
                let mut latin = BytesStart::new("a:latin");
                latin.push_attribute(("typeface", ts.font_name.as_str()));
                w.write_event(Event::Empty(latin))?;
            }
            w.write_event(Event::End(BytesEnd::new("a:rPr")))?;
        } else {
            w.write_event(Event::Empty(r_pr))?;
        }
        // <a:t>text</a:t>
        w.write_event(Event::Start(BytesStart::new("a:t")))?;
        w.write_event(Event::Text(quick_xml::events::BytesText::new(trimmed)))?;
        w.write_event(Event::End(BytesEnd::new("a:t")))?;
        w.write_event(Event::End(BytesEnd::new("a:r")))?;
        w.write_event(Event::End(BytesEnd::new("a:p")))?;
        w.write_event(Event::End(BytesEnd::new("xdr:txBody")))?;

        w.write_event(Event::End(BytesEnd::new("xdr:sp")))?;
        w.write_event(Event::Empty(BytesStart::new("xdr:clientData")))?;
        w.write_event(Event::End(BytesEnd::new("xdr:absoluteAnchor")))?;
    }

    w.write_event(Event::End(BytesEnd::new("xdr:wsDr")))?;
    Ok(w.into_inner())
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
    /// Ordered list of fonts for XML serialization.
    fonts: Vec<FontKey>,
    /// Ordered list of fills for XML serialization.
    fills: Vec<FillKey>,
    num_fmts: Vec<(u32, String)>, // (numFmtId, formatCode) for custom formats
    /// Ordered list of xf records for XML serialization.
    xfs: Vec<XfKey>,
    // Lookup maps for O(1) deduplication during build.
    font_map: HashMap<FontKey, u32>,
    fill_map: HashMap<FillKey, u32>,
    xf_map: HashMap<XfKey, u32>,
}

impl StyleTable {
    fn build(sheets: &[SheetDataInner]) -> Self {
        let default_font = FontKey {
            bold: false,
            italic: false,
            underline: false,
            color: None,
            size_half_pt: None,
            name: None,
        };
        let default_xf = XfKey {
            font_idx: 0,
            fill_idx: 0,
            num_fmt_id: 0,
            h_align: None,
            wrap_text: false,
        };

        let mut font_map = HashMap::new();
        font_map.insert(default_font.clone(), 0u32);
        let mut fill_map: HashMap<FillKey, u32> = HashMap::new();
        fill_map.insert(FillKey(None), 0u32); // idx 0 = none; idx 1 = gray125 (pre-populated below)
        let mut xf_map = HashMap::new();
        xf_map.insert(default_xf.clone(), 0u32);

        let mut table = StyleTable {
            cell_xf: HashMap::new(),
            fonts: vec![default_font],
            fills: vec![FillKey(None), FillKey(None)], // idx 0: none, idx 1: gray125
            num_fmts: Vec::new(),
            xfs: vec![default_xf],
            font_map,
            fill_map,
            xf_map,
        };

        let mut next_custom_fmt_id: u32 = 164; // custom numFmtIds start at 164

        for sheet in sheets {
            let sheet_ptr = sheet as *const SheetDataInner;
            for ((row, col), style) in &sheet.cell_styles {
                // Resolve font index — O(1) via HashMap.
                let font_key = FontKey {
                    bold: style.bold,
                    italic: style.italic,
                    underline: style.underline,
                    color: style.font_color.clone(),
                    size_half_pt: style.font_size_pt.map(|s| (s * 2.0).round() as u32),
                    name: style.font_name.clone(),
                };
                let font_idx = if let Some(&i) = table.font_map.get(&font_key) {
                    i
                } else {
                    let idx = table.fonts.len() as u32;
                    table.fonts.push(font_key.clone());
                    table.font_map.insert(font_key, idx);
                    idx
                };

                // Resolve fill index — O(1) via HashMap.
                let fill_key = FillKey(style.background_color.clone());
                let fill_idx = if fill_key.0.is_none() {
                    0
                } else if let Some(&i) = table.fill_map.get(&fill_key) {
                    i
                } else {
                    let idx = table.fills.len() as u32;
                    table.fills.push(fill_key.clone());
                    table.fill_map.insert(fill_key, idx);
                    idx
                };

                // Resolve number format id.
                let num_fmt_id = match style.number_format.builtin_id() {
                    Some(id) => id,
                    None => {
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

                // Resolve xf index — O(1) via HashMap.
                let xf_idx = if let Some(&i) = table.xf_map.get(&xf_key) {
                    i
                } else {
                    let idx = table.xfs.len() as u32;
                    table.xfs.push(xf_key.clone());
                    table.xf_map.insert(xf_key, idx);
                    idx
                };

                table.cell_xf.insert((sheet_ptr, *row, *col), xf_idx);
            }
        }

        table
    }

    fn get_idx(&self, sheet: &SheetDataInner, row: usize, col: usize) -> Option<u32> {
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
            CellStyle::new()
                .bold()
                .background("4472C4")
                .font_color("FFFFFF"),
        );
        sheet.set_cell_styled(
            1,
            0,
            CellData::Formula("SUM(A3:A10)".into()),
            CellStyle::new()
                .bold()
                .number_format(NumberFormat::Currency),
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

    #[test]
    fn page_setup_round_trip() {
        // Letter portrait, 0.5" margins. The on-wire format is mm in
        // <pageSetup paperWidth/paperHeight> + inches in <pageMargins>;
        // verify both elements appear and that the parser recovers
        // values within rounding tolerance.
        let mut wb = XlsxWriter::new();
        let mut sheet = wb.add_sheet("Geom");
        sheet.set_cell(0, 0, CellData::String("hi".into()));
        sheet.set_page_setup(PageSetup {
            width_twips: 12240,    // 8.5"
            height_twips: 15840,   // 11"
            margin_top_twips: 720, // 0.5"
            margin_bottom_twips: 720,
            margin_left_twips: 720,
            margin_right_twips: 720,
            header_distance_twips: 432, // 0.3"
            footer_distance_twips: 432,
            landscape: false,
        });
        let mut buf = std::io::Cursor::new(Vec::new());
        wb.write_to(&mut buf).expect("write");

        // Pull sheet1.xml out and check the attributes.
        buf.set_position(0);
        let mut zip = zip::ZipArchive::new(buf).expect("zip");
        let mut xml = String::new();
        {
            let mut entry = zip.by_name("xl/worksheets/sheet1.xml").expect("sheet");
            std::io::Read::read_to_string(&mut entry, &mut xml).expect("read");
        }
        assert!(xml.contains("<pageMargins"), "missing pageMargins: {xml}");
        assert!(xml.contains("<pageSetup"), "missing pageSetup: {xml}");
        assert!(xml.contains(r#"orientation="portrait""#));
        // 8.5" = 215.90mm, 11" = 279.40mm
        assert!(xml.contains(r#"paperWidth="215.90mm""#), "width attr: {xml}");
        assert!(xml.contains(r#"paperHeight="279.40mm""#), "height attr: {xml}");
        assert!(xml.contains(r#"left="0.5000""#));
    }

    #[test]
    fn merge_cells_xml() {
        let mut wb = XlsxWriter::new();
        let mut sheet = wb.add_sheet("MergeTest");
        sheet.set_cell(0, 0, CellData::String("Merged".into()));
        sheet.merge_cells(0, 0, 1, 2);

        let mut buf = std::io::Cursor::new(Vec::new());
        wb.write_to(&mut buf).expect("write xlsx");

        // Extract sheet1.xml from the zip and verify mergeCells element
        buf.set_position(0);
        let mut zip = zip::ZipArchive::new(buf).expect("open zip");
        let mut sheet_xml = String::new();
        {
            let mut entry = zip.by_name("xl/worksheets/sheet1.xml").expect("find sheet");
            std::io::Read::read_to_string(&mut entry, &mut sheet_xml).expect("read");
        }
        assert!(sheet_xml.contains("<mergeCells"), "missing mergeCells");
        assert!(sheet_xml.contains(r#"ref="A1:B1""#), "wrong ref");
    }
}
