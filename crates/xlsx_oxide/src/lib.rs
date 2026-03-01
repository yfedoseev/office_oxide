//! # xlsx_oxide
//!
//! High-performance Excel spreadsheet (.xlsx) processing.
//!
//! Read, convert, and extract content from XLSX files
//! (Office Open XML SpreadsheetML, ISO 29500 / ECMA-376).
//!
//! # Quick Start
//!
//! ```rust,no_run
//! use xlsx_oxide::XlsxDocument;
//!
//! let doc = XlsxDocument::open("data.xlsx").unwrap();
//! println!("{}", doc.plain_text());
//! println!("{}", doc.to_csv());
//! println!("{}", doc.to_markdown());
//! ```

pub mod cell;
pub mod date;
pub mod error;
pub mod shared_strings;
pub mod styles;
pub mod text;
pub mod workbook;
pub mod worksheet;
pub mod write;
pub mod edit;

pub use cell::{Cell, CellRef, CellValue};
pub use date::DateTimeValue;
pub use error::{Result, XlsxError};
pub use shared_strings::SharedStringTable;
pub use styles::StyleSheet;
pub use workbook::{SheetState, WorkbookInfo};
pub use worksheet::{HyperlinkInfo, HyperlinkTarget, Worksheet};

use std::io::{Read, Seek};
use std::path::Path;

use log::debug;
use office_core::opc::OpcReader;
use office_core::relationships::{rel_types, Relationships};
use office_core::theme::Theme;

/// A parsed XLSX document.
#[derive(Debug, Clone)]
pub struct XlsxDocument {
    pub workbook: WorkbookInfo,
    pub worksheets: Vec<Worksheet>,
    pub shared_strings: SharedStringTable,
    pub styles: Option<StyleSheet>,
    pub theme: Option<Theme>,
}

impl XlsxDocument {
    /// Open an XLSX file from a file path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let reader = OpcReader::open(path)?;
        Self::from_opc(reader)
    }

    /// Open an XLSX file using memory-mapped I/O for better performance on large files.
    #[cfg(feature = "mmap")]
    pub fn open_mmap(path: impl AsRef<Path>) -> Result<Self> {
        let reader = OpcReader::open_mmap(path)?;
        Self::from_opc(reader)
    }

    /// Open an XLSX document from any `Read + Seek` source.
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let opc = OpcReader::new(reader)?;
        Self::from_opc(opc)
    }

    fn from_opc<R: Read + Seek>(mut opc: OpcReader<R>) -> Result<Self> {
        debug!("XlsxDocument: parsing started");
        let main_part = opc.main_document_part()?;
        let wb_rels = opc.read_rels_for(&main_part)?;

        // Parse shared strings (must be first — cells reference by index)
        let shared_strings = if let Some(rel) = wb_rels.first_by_type(rel_types::SHARED_STRINGS) {
            let part_name = main_part.resolve_relative(&rel.target)?;
            let data = opc.read_part(&part_name)?;
            SharedStringTable::parse(&data)?
        } else {
            SharedStringTable::empty()
        };

        // Parse theme (optional — some files reference it but don't include it)
        let theme = if let Some(rel) = wb_rels.first_by_type(rel_types::THEME) {
            let part_name = main_part.resolve_relative(&rel.target)?;
            match opc.read_part(&part_name) {
                Ok(data) => Theme::parse(&data).ok(),
                Err(_) => None,
            }
        } else {
            None
        };

        // Parse styles
        let styles = if let Some(rel) = wb_rels.first_by_type(rel_types::STYLES) {
            let part_name = main_part.resolve_relative(&rel.target)?;
            let data = opc.read_part(&part_name)?;
            Some(StyleSheet::parse(&data)?)
        } else {
            None
        };

        // Parse workbook
        let wb_data = opc.read_part(&main_part)?;
        let workbook = WorkbookInfo::parse(&wb_data)?;

        // Phase 1: gather raw data sequentially (requires &mut opc)
        struct SheetBundle {
            name: String,
            data: Vec<u8>,
            rels: Relationships,
        }
        let mut bundles = Vec::with_capacity(workbook.sheets.len());
        for sheet in &workbook.sheets {
            // Skip sheets with empty r:id (virtual sheets, VBA modules, etc.)
            if sheet.rel_id.is_empty() {
                continue;
            }
            let part_name = match wb_rels.resolve_target(&sheet.rel_id, &main_part) {
                Ok(pn) => pn,
                Err(_) => {
                    // Fallback: scan ZIP entries for xl/worksheets/sheet<N>.xml
                    // This handles files with corrupted workbook.xml.rels
                    let idx = bundles.len() + 1;
                    let candidates = [
                        format!("/xl/worksheets/sheet{}.xml", idx),
                        format!("/xl/worksheets/sheet{}.xml", sheet.sheet_id),
                    ];
                    match candidates
                        .iter()
                        .find_map(|c| office_core::opc::PartName::new(c).ok().filter(|pn| opc.has_part(pn)))
                    {
                        Some(pn) => {
                            debug!("worksheet fallback: '{}' -> '{}'", sheet.rel_id, pn);
                            pn
                        }
                        None => continue,
                    }
                }
            };
            let ws_rels = opc
                .read_rels_for(&part_name)
                .unwrap_or_else(|_| Relationships::empty());
            let ws_data = match opc.read_part(&part_name) {
                Ok(data) => data,
                Err(_) => continue, // Skip unreadable sheets
            };
            bundles.push(SheetBundle {
                name: sheet.name.clone(),
                data: ws_data,
                rels: ws_rels,
            });
        }

        // Phase 2: parse worksheets (parallel when feature enabled)
        #[cfg(feature = "parallel")]
        let worksheets: Result<Vec<Worksheet>> = {
            use rayon::prelude::*;
            bundles
                .into_par_iter()
                .map(|b| {
                    let mut ws = Worksheet::parse(&b.data, b.name, &b.rels)?;
                    dereference_shared_strings(&mut ws, &shared_strings);
                    Ok(ws)
                })
                .collect()
        };
        #[cfg(not(feature = "parallel"))]
        let worksheets: Result<Vec<Worksheet>> = bundles
            .into_iter()
            .map(|b| {
                let mut ws = Worksheet::parse(&b.data, b.name, &b.rels)?;
                dereference_shared_strings(&mut ws, &shared_strings);
                Ok(ws)
            })
            .collect();
        let worksheets = worksheets?;

        debug!("XlsxDocument: {} worksheets parsed", worksheets.len());
        Ok(XlsxDocument {
            workbook,
            worksheets,
            shared_strings,
            styles,
            theme,
        })
    }

}

/// Maximum string length per cell during shared string dereference.
/// Prevents excessive memory allocation from DoS test files (e.g., 1MB string × 12,000 cells).
const MAX_CELL_STRING_LEN: usize = 32_768;

/// Dereference `CellValue::SharedString(idx)` into `CellValue::String(text)`.
fn dereference_shared_strings(ws: &mut Worksheet, sst: &SharedStringTable) {
    for row in &mut ws.rows {
        for cell in &mut row.cells {
            if let CellValue::SharedString(idx) = &cell.value {
                let full_text = sst.get(*idx).unwrap_or("");
                let text = if full_text.len() > MAX_CELL_STRING_LEN {
                    // Truncate at a char boundary
                    let mut end = MAX_CELL_STRING_LEN;
                    while !full_text.is_char_boundary(end) && end > 0 {
                        end -= 1;
                    }
                    full_text[..end].to_string()
                } else {
                    full_text.to_string()
                };
                cell.value = CellValue::String(text);
            }
        }
    }
}
