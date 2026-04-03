//! Workbook-level parsing: sheets, SST, cell grid construction.

use std::io::{Read, Seek};

use cfb_oxide::CfbReader;

use crate::cell::{parse_cell_record, Cell, CellValue};
use crate::error::{Result, XlsError};
use crate::images::{extract_images, XlsImage};
use crate::records::*;
use crate::sst::{parse_sst, read_short_unicode_string, read_unicode_string};

/// A parsed legacy XLS document.
#[derive(Debug)]
pub struct XlsDocument {
    pub sheets: Vec<Sheet>,
    images: Vec<XlsImage>,
}

/// A worksheet.
#[derive(Debug)]
pub struct Sheet {
    pub name: String,
    pub rows: Vec<Vec<CellValue>>,
}

/// Sheet metadata from BOUNDSHEET records.
#[derive(Debug)]
struct SheetInfo {
    name: String,
    #[allow(dead_code)]
    offset: u32,
    hidden: bool,
}

impl XlsDocument {
    /// Open an XLS file from a reader (any `Read + Seek`).
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let mut cfb = CfbReader::new(reader)?;

        // Try "Workbook" (BIFF8) first, then "Book" (BIFF5).
        let stream_data = if cfb.has_stream("Workbook") {
            cfb.open_stream("Workbook")?
        } else if cfb.has_stream("Book") {
            cfb.open_stream("Book")?
        } else {
            return Err(XlsError::MissingStream(
                "neither Workbook nor Book stream found".into(),
            ));
        };
        // Drop CFB early to free file handle and memory.
        drop(cfb);

        let mut doc = Self::parse_workbook_stream(&stream_data)?;
        doc.images = extract_images(&stream_data);
        Ok(doc)
    }

    /// Open an XLS file from a path.
    pub fn open<P: AsRef<std::path::Path>>(path: P) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    fn parse_workbook_stream(data: &[u8]) -> Result<Self> {
        let mut sheet_infos: Vec<SheetInfo> = Vec::new();
        let mut sst: Vec<String> = Vec::new();
        let mut sheets = Vec::new();

        // Quick check: if the first BOF indicates BIFF5 or earlier, limit processing.
        let biff8 = data.len() >= 6 && {
            let rt = u16::from_le_bytes([data[0], data[1]]);
            let ver = if rt == RT_BOF && data.len() >= 6 {
                u16::from_le_bytes([data[4], data[5]])
            } else {
                0
            };
            ver == 0x0600 // BIFF8
        };

        // Single-pass parsing: globals then sheets sequentially.
        let mut phase = Phase::Globals;
        let mut cells: Vec<Cell> = Vec::new();
        let mut sheet_idx = 0usize;
        let mut pending_formula_string: Option<(u16, u16)> = None;
        let mut record_budget = 500_000u32; // Safety cap to prevent pathological files

        for rec in RecordIter::new(data) {
            if record_budget == 0 {
                break;
            }
            record_budget -= 1;
            let rec = rec?;
            match phase {
                Phase::Globals => match rec.record_type {
                    RT_FILEPASS => {
                        // File is encrypted — we can't read it.
                        return Ok(Self { sheets: Vec::new(), images: Vec::new() });
                    }
                    RT_BOUNDSHEET => {
                        if let Ok(info) = parse_boundsheet(&rec.data) {
                            sheet_infos.push(info);
                        }
                    }
                    RT_SST => {
                        sst = parse_sst(&rec.data)?;
                    }
                    RT_EOF => {
                        phase = Phase::BetweenSheets;
                    }
                    _ => {}
                },
                Phase::BetweenSheets => {
                    if rec.record_type == RT_BOF {
                        phase = Phase::InSheet;
                        cells.clear();
                        pending_formula_string = None;
                    }
                }
                Phase::InSheet => match rec.record_type {
                    RT_EOF => {
                        let name = if sheet_idx < sheet_infos.len() {
                            sheet_infos[sheet_idx].name.clone()
                        } else {
                            format!("Sheet{}", sheet_idx + 1)
                        };
                        let hidden =
                            sheet_idx < sheet_infos.len() && sheet_infos[sheet_idx].hidden;
                        if !hidden {
                            let rows = build_grid(&mut cells);
                            sheets.push(Sheet { name, rows });
                        }
                        sheet_idx += 1;
                        phase = Phase::BetweenSheets;
                    }
                    RT_STRING => {
                        if let Some((row, col)) = pending_formula_string.take() {
                            if rec.data.len() >= 3 {
                                if let Ok((s, _)) = read_unicode_string(&rec.data, 0) {
                                    cells.push(Cell {
                                        row,
                                        col,
                                        value: CellValue::String(s),
                                    });
                                }
                            }
                        }
                    }
                    RT_FORMULA => {
                        pending_formula_string = None;
                        if rec.data.len() >= 14 {
                            let val_bytes = &rec.data[6..14];
                            if val_bytes[6] == 0xFF && val_bytes[7] == 0xFF && val_bytes[0] == 0 {
                                let row = u16::from_le_bytes([rec.data[0], rec.data[1]]);
                                let col = u16::from_le_bytes([rec.data[2], rec.data[3]]);
                                pending_formula_string = Some((row, col));
                                continue;
                            }
                        }
                        if let Ok(parsed) = parse_cell_record(&rec, &sst) {
                            cells.extend(parsed);
                        }
                    }
                    _ => {
                        pending_formula_string = None;
                        // Skip LABEL/RSTRING parsing for non-BIFF8 (avoids slow unicode fallback).
                        if !biff8 && matches!(rec.record_type, RT_LABEL | RT_RSTRING) {
                            // BIFF5 LABEL: extract text directly.
                            if rec.data.len() >= 8 {
                                let row = u16::from_le_bytes([rec.data[0], rec.data[1]]);
                                let col = u16::from_le_bytes([rec.data[2], rec.data[3]]);
                                let str_len = u16::from_le_bytes([rec.data[6], rec.data[7]]) as usize;
                                let start = 8;
                                let end = (start + str_len).min(rec.data.len());
                                let s: String = rec.data[start..end].iter().map(|&b| b as char).collect();
                                cells.push(Cell { row, col, value: CellValue::String(s) });
                            }
                        } else if let Ok(parsed) = parse_cell_record(&rec, &sst) {
                            cells.extend(parsed);
                        }
                    }
                },
            }
        }

        Ok(Self { sheets, images: Vec::new() })
    }

    /// Get all extracted images.
    pub fn images(&self) -> &[XlsImage] {
        &self.images
    }

    /// Extract plain text from the document.
    pub fn plain_text(&self) -> String {
        let mut out = String::new();
        for (i, sheet) in self.sheets.iter().enumerate() {
            if i > 0 {
                out.push('\n');
            }
            out.push_str(&sheet.name);
            out.push('\n');
            for row in &sheet.rows {
                let line: Vec<String> = row.iter().map(|c| c.as_text()).collect();
                let trimmed = line.join("\t").trim_end().to_string();
                out.push_str(&trimmed);
                out.push('\n');
            }
        }
        out
    }

    /// Convert to markdown.
    pub fn to_markdown(&self) -> String {
        let mut out = String::new();
        for (i, sheet) in self.sheets.iter().enumerate() {
            if i > 0 {
                out.push('\n');
            }
            out.push_str("## ");
            out.push_str(&sheet.name);
            out.push_str("\n\n");

            if sheet.rows.is_empty() {
                continue;
            }

            // Header row.
            let col_count = sheet.rows.iter().map(|r| r.len()).max().unwrap_or(0);
            if col_count == 0 {
                continue;
            }

            // First row as header.
            out.push('|');
            if let Some(first_row) = sheet.rows.first() {
                for c in 0..col_count {
                    let text = first_row.get(c).map(|v| v.as_text()).unwrap_or_default();
                    out.push(' ');
                    out.push_str(&text);
                    out.push_str(" |");
                }
            }
            out.push('\n');

            // Separator.
            out.push('|');
            for _ in 0..col_count {
                out.push_str(" --- |");
            }
            out.push('\n');

            // Data rows.
            for row in sheet.rows.iter().skip(1) {
                out.push('|');
                for c in 0..col_count {
                    let text = row.get(c).map(|v| v.as_text()).unwrap_or_default();
                    out.push(' ');
                    out.push_str(&text);
                    out.push_str(" |");
                }
                out.push('\n');
            }
        }
        out
    }
}

enum Phase {
    Globals,
    BetweenSheets,
    InSheet,
}

fn parse_boundsheet(data: &[u8]) -> Result<SheetInfo> {
    if data.len() < 8 {
        return Err(XlsError::InvalidRecord("BOUNDSHEET too short".into()));
    }
    let offset = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
    let visibility = data[4]; // 0=visible, 1=hidden, 2=very hidden
    let _sheet_type = data[5]; // 0=worksheet, 2=chart, 6=VBA
    let (name, _) = read_short_unicode_string(data, 6)?;

    Ok(SheetInfo {
        name,
        offset,
        hidden: visibility != 0,
    })
}

/// Build a 2D grid from sparse cells.
///
/// Takes ownership of cell values via `std::mem::take` to avoid cloning.
fn build_grid(cells: &mut [Cell]) -> Vec<Vec<CellValue>> {
    if cells.is_empty() {
        return Vec::new();
    }

    let max_row = cells.iter().map(|c| c.row).max().unwrap_or(0) as usize;
    let max_col = cells.iter().map(|c| c.col).max().unwrap_or(0) as usize;

    // Cap to prevent OOM on pathological files. If the grid would exceed 4M cells, use
    // a compact representation: only allocate rows that have data.
    let max_row = max_row.min(65535);
    let max_col = max_col.min(255);
    let grid_size = (max_row + 1) * (max_col + 1);

    if grid_size > 4_000_000 {
        // Sparse: only create rows with actual data.
        return build_grid_sparse(cells, max_col);
    }

    let mut grid = vec![vec![CellValue::Empty; max_col + 1]; max_row + 1];
    for cell in cells.iter_mut() {
        let r = cell.row as usize;
        let c = cell.col as usize;
        if r <= max_row && c <= max_col {
            grid[r][c] = std::mem::take(&mut cell.value);
        }
    }

    // Trim trailing empty rows.
    while grid.last().is_some_and(|row| row.iter().all(|c| matches!(c, CellValue::Empty))) {
        grid.pop();
    }

    grid
}

/// Sparse grid builder for large/sparse sheets.
fn build_grid_sparse(cells: &mut [Cell], max_col: usize) -> Vec<Vec<CellValue>> {
    // Sort cells by row, then column.
    cells.sort_unstable_by(|a, b| a.row.cmp(&b.row).then(a.col.cmp(&b.col)));

    let mut grid: Vec<Vec<CellValue>> = Vec::new();
    let mut current_row = u16::MAX;

    for cell in cells.iter_mut() {
        let r = cell.row as usize;
        let c = cell.col as usize;
        if c > max_col {
            continue;
        }

        // Fill missing rows.
        while grid.len() <= r {
            grid.push(vec![CellValue::Empty; max_col + 1]);
        }

        if cell.row != current_row {
            current_row = cell.row;
        }
        grid[r][c] = std::mem::take(&mut cell.value);
    }

    // Trim trailing empty rows.
    while grid.last().is_some_and(|row| row.iter().all(|c| matches!(c, CellValue::Empty))) {
        grid.pop();
    }

    grid
}

impl office_core::OfficeDocument for XlsDocument {
    fn plain_text(&self) -> String {
        self.plain_text()
    }

    fn to_markdown(&self) -> String {
        self.to_markdown()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn build_grid_from_cells() {
        let mut cells = vec![
            Cell { row: 0, col: 0, value: CellValue::String("A1".into()) },
            Cell { row: 0, col: 1, value: CellValue::Number(42.0) },
            Cell { row: 1, col: 0, value: CellValue::String("A2".into()) },
        ];
        let grid = build_grid(&mut cells);
        assert_eq!(grid.len(), 2);
        assert_eq!(grid[0].len(), 2);
        assert_eq!(grid[0][0], CellValue::String("A1".into()));
        assert_eq!(grid[0][1], CellValue::Number(42.0));
        assert_eq!(grid[1][0], CellValue::String("A2".into()));
        assert_eq!(grid[1][1], CellValue::Empty);
    }

    #[test]
    fn build_grid_empty() {
        let grid = build_grid(&mut Vec::new());
        assert!(grid.is_empty());
    }

    #[test]
    fn parse_boundsheet_record() {
        let mut data = Vec::new();
        data.extend_from_slice(&100u32.to_le_bytes()); // offset
        data.push(0); // visible
        data.push(0); // worksheet
        // Short string "Sheet1"
        data.push(6); // char count
        data.push(0); // compressed
        data.extend_from_slice(b"Sheet1");
        let info = parse_boundsheet(&data).unwrap();
        assert_eq!(info.name, "Sheet1");
        assert_eq!(info.offset, 100);
        assert!(!info.hidden);
    }

    #[test]
    fn plain_text_output() {
        let doc = XlsDocument {
            images: Vec::new(),
            sheets: vec![Sheet {
                name: "Sheet1".into(),
                rows: vec![
                    vec![CellValue::String("Name".into()), CellValue::String("Age".into())],
                    vec![CellValue::String("Alice".into()), CellValue::Number(30.0)],
                ],
            }],
        };
        let text = doc.plain_text();
        assert!(text.contains("Sheet1"));
        assert!(text.contains("Name\tAge"));
        assert!(text.contains("Alice\t30"));
    }

    #[test]
    fn markdown_output() {
        let doc = XlsDocument {
            images: Vec::new(),
            sheets: vec![Sheet {
                name: "Data".into(),
                rows: vec![
                    vec![CellValue::String("X".into()), CellValue::String("Y".into())],
                    vec![CellValue::Number(1.0), CellValue::Number(2.0)],
                ],
            }],
        };
        let md = doc.to_markdown();
        assert!(md.contains("## Data"));
        assert!(md.contains("| X | Y |"));
        assert!(md.contains("| --- | --- |"));
        assert!(md.contains("| 1 | 2 |"));
    }
}
