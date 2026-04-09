//! XLSX creation (write) module.
//!
//! Provides a builder API for creating XLSX files from scratch using
//! inline strings (no shared string table).

use std::io::{Seek, Write};
use std::path::Path;

use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};

use crate::core::opc::{OpcWriter, PartName};
use crate::core::relationships::rel_types;

use super::Result;

// ---------------------------------------------------------------------------
// Content type constants
// ---------------------------------------------------------------------------

const CT_WORKBOOK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_STYLES: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";

// ---------------------------------------------------------------------------
// SML namespace constants
// ---------------------------------------------------------------------------

use crate::core::xml::ns::{R_STR as NS_REL, SML_STR as NS_SML};

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/// A value to write into a cell.
#[derive(Debug, Clone)]
pub enum CellData {
    /// An empty cell (no value written).
    Empty,
    /// A string value (written as an inline string).
    String(String),
    /// A numeric value.
    Number(f64),
    /// A boolean value.
    Boolean(bool),
}

/// Data for a single worksheet.
#[derive(Debug, Clone)]
pub struct SheetData {
    name: String,
    rows: Vec<Vec<CellData>>,
}

impl SheetData {
    /// Append a row of cells and return `&mut Self` for chaining.
    pub fn add_row(&mut self, cells: Vec<CellData>) -> &mut Self {
        self.rows.push(cells);
        self
    }

    /// Set a specific cell value, expanding the grid as needed.
    ///
    /// Both `row` and `col` are 0-based indices.
    pub fn set_cell(&mut self, row: usize, col: usize, value: CellData) -> &mut Self {
        // Ensure enough rows exist
        if self.rows.len() <= row {
            self.rows.resize_with(row + 1, Vec::new);
        }
        // Ensure the target row has enough columns
        let row_data = &mut self.rows[row];
        if row_data.len() <= col {
            row_data.resize_with(col + 1, || CellData::Empty);
        }
        row_data[col] = value;
        self
    }
}

/// Builder for creating XLSX files.
#[derive(Debug, Clone)]
pub struct XlsxWriter {
    sheets: Vec<SheetData>,
}

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

    /// Add a worksheet with the given name and return a mutable reference to it.
    pub fn add_sheet(&mut self, name: &str) -> &mut SheetData {
        self.sheets.push(SheetData {
            name: name.to_string(),
            rows: Vec::new(),
        });
        self.sheets.last_mut().unwrap()
    }

    /// Save the workbook to a file path.
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
    // Internal: write all OPC parts
    // -----------------------------------------------------------------------

    fn write_parts<W: Write + Seek>(&self, opc: &mut OpcWriter<W>) -> Result<()> {
        let wb_part = PartName::new("/xl/workbook.xml")?;

        // Package-level relationship: workbook
        opc.add_package_rel(rel_types::OFFICE_DOCUMENT, "xl/workbook.xml");

        // Part-level relationships for workbook.
        // IMPORTANT: add worksheet rels BEFORE styles so rIds match workbook XML.
        let mut sheet_rids = Vec::with_capacity(self.sheets.len());
        for (i, _) in self.sheets.iter().enumerate() {
            let target = format!("worksheets/sheet{}.xml", i + 1);
            let rid = opc.add_part_rel(&wb_part, rel_types::WORKSHEET, &target);
            sheet_rids.push(rid);
        }
        opc.add_part_rel(&wb_part, rel_types::STYLES, "styles.xml");

        // Write workbook part
        let wb_xml = self.build_workbook_xml(&sheet_rids)?;
        opc.add_part(&wb_part, CT_WORKBOOK, &wb_xml)?;

        // Write worksheet parts
        for (i, sheet) in self.sheets.iter().enumerate() {
            let part_name_str = format!("/xl/worksheets/sheet{}.xml", i + 1);
            let part_name = PartName::new(&part_name_str)?;
            let ws_xml = Self::build_worksheet_xml(sheet)?;
            opc.add_part(&part_name, CT_WORKSHEET, &ws_xml)?;
        }

        // Write styles part
        let styles_part = PartName::new("/xl/styles.xml")?;
        let styles_xml = Self::build_styles_xml()?;
        opc.add_part(&styles_part, CT_STYLES, &styles_xml)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // XML generation: workbook.xml
    // -----------------------------------------------------------------------

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

    // -----------------------------------------------------------------------
    // XML generation: worksheet
    // -----------------------------------------------------------------------

    fn build_worksheet_xml(sheet: &SheetData) -> crate::core::Result<Vec<u8>> {
        let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

        w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))?;

        let mut root = BytesStart::new("worksheet");
        root.push_attribute(("xmlns", NS_SML));
        w.write_event(Event::Start(root))?;

        w.write_event(Event::Start(BytesStart::new("sheetData")))?;

        for (row_idx, row) in sheet.rows.iter().enumerate() {
            if row.is_empty() {
                continue;
            }

            let row_num = (row_idx + 1).to_string();
            let mut row_elem = BytesStart::new("row");
            row_elem.push_attribute(("r", row_num.as_str()));
            w.write_event(Event::Start(row_elem))?;

            for (col_idx, cell) in row.iter().enumerate() {
                Self::write_cell(&mut w, row_idx, col_idx, cell)?;
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
    ) -> crate::core::Result<()> {
        let cell_ref = format!("{}{}", col_name(col as u32), row + 1);

        match cell {
            CellData::Empty => {
                // Skip empty cells entirely
            },
            CellData::String(s) => {
                let mut c = BytesStart::new("c");
                c.push_attribute(("r", cell_ref.as_str()));
                c.push_attribute(("t", "inlineStr"));
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
                c.push_attribute(("t", "n"));
                w.write_event(Event::Start(c))?;

                w.write_event(Event::Start(BytesStart::new("v")))?;
                let text = format_number(*n);
                w.write_event(Event::Text(BytesText::new(&text)))?;
                w.write_event(Event::End(BytesEnd::new("v")))?;

                w.write_event(Event::End(BytesEnd::new("c")))?;
            },
            CellData::Boolean(b) => {
                let mut c = BytesStart::new("c");
                c.push_attribute(("r", cell_ref.as_str()));
                c.push_attribute(("t", "b"));
                w.write_event(Event::Start(c))?;

                w.write_event(Event::Start(BytesStart::new("v")))?;
                let text = if *b { "1" } else { "0" };
                w.write_event(Event::Text(BytesText::new(text)))?;
                w.write_event(Event::End(BytesEnd::new("v")))?;

                w.write_event(Event::End(BytesEnd::new("c")))?;
            },
        }

        Ok(())
    }

    // -----------------------------------------------------------------------
    // XML generation: styles.xml (minimal)
    // -----------------------------------------------------------------------

    fn build_styles_xml() -> crate::core::Result<Vec<u8>> {
        let mut w = Writer::new_with_indent(Vec::new(), b' ', 2);

        w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))?;

        let mut root = BytesStart::new("styleSheet");
        root.push_attribute(("xmlns", NS_SML));
        w.write_event(Event::Start(root))?;

        // fonts
        let mut fonts = BytesStart::new("fonts");
        fonts.push_attribute(("count", "1"));
        w.write_event(Event::Start(fonts))?;
        w.write_event(Event::Start(BytesStart::new("font")))?;
        let mut sz = BytesStart::new("sz");
        sz.push_attribute(("val", "11"));
        w.write_event(Event::Empty(sz))?;
        let mut name = BytesStart::new("name");
        name.push_attribute(("val", "Calibri"));
        w.write_event(Event::Empty(name))?;
        w.write_event(Event::End(BytesEnd::new("font")))?;
        w.write_event(Event::End(BytesEnd::new("fonts")))?;

        // fills
        let mut fills = BytesStart::new("fills");
        fills.push_attribute(("count", "2"));
        w.write_event(Event::Start(fills))?;
        // fill 1: none
        w.write_event(Event::Start(BytesStart::new("fill")))?;
        let mut pf_none = BytesStart::new("patternFill");
        pf_none.push_attribute(("patternType", "none"));
        w.write_event(Event::Empty(pf_none))?;
        w.write_event(Event::End(BytesEnd::new("fill")))?;
        // fill 2: gray125
        w.write_event(Event::Start(BytesStart::new("fill")))?;
        let mut pf_gray = BytesStart::new("patternFill");
        pf_gray.push_attribute(("patternType", "gray125"));
        w.write_event(Event::Empty(pf_gray))?;
        w.write_event(Event::End(BytesEnd::new("fill")))?;
        w.write_event(Event::End(BytesEnd::new("fills")))?;

        // borders
        let mut borders = BytesStart::new("borders");
        borders.push_attribute(("count", "1"));
        w.write_event(Event::Start(borders))?;
        w.write_event(Event::Start(BytesStart::new("border")))?;
        w.write_event(Event::Empty(BytesStart::new("left")))?;
        w.write_event(Event::Empty(BytesStart::new("right")))?;
        w.write_event(Event::Empty(BytesStart::new("top")))?;
        w.write_event(Event::Empty(BytesStart::new("bottom")))?;
        w.write_event(Event::End(BytesEnd::new("border")))?;
        w.write_event(Event::End(BytesEnd::new("borders")))?;

        // cellStyleXfs
        let mut csxf = BytesStart::new("cellStyleXfs");
        csxf.push_attribute(("count", "1"));
        w.write_event(Event::Start(csxf))?;
        w.write_event(Event::Empty(BytesStart::new("xf")))?;
        w.write_event(Event::End(BytesEnd::new("cellStyleXfs")))?;

        // cellXfs
        let mut cxf = BytesStart::new("cellXfs");
        cxf.push_attribute(("count", "1"));
        w.write_event(Event::Start(cxf))?;
        let mut xf = BytesStart::new("xf");
        xf.push_attribute(("xfId", "0"));
        w.write_event(Event::Empty(xf))?;
        w.write_event(Event::End(BytesEnd::new("cellXfs")))?;

        w.write_event(Event::End(BytesEnd::new("styleSheet")))?;

        Ok(w.into_inner())
    }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// Convert a 0-based column index to a column letter string: 0 -> "A", 25 -> "Z", 26 -> "AA".
fn col_name(col: u32) -> String {
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

/// Format a number for XML output. Integers are written without a decimal point.
fn format_number(n: f64) -> String {
    if n.fract() == 0.0 && n.abs() < (i64::MAX as f64) {
        format!("{}", n as i64)
    } else {
        format!("{}", n)
    }
}
