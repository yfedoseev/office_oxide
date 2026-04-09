//! # office_oxide::xlsx
//!
//! High-performance Excel spreadsheet (.xlsx) processing.
//!
//! Read, convert, and extract content from XLSX files
//! (Office Open XML SpreadsheetML, ISO 29500 / ECMA-376).
//!
//! # Quick Start
//!
//! ```rust,no_run
//! use office_oxide::xlsx::XlsxDocument;
//!
//! let doc = XlsxDocument::open("data.xlsx").unwrap();
//! println!("{}", doc.plain_text());
//! println!("{}", doc.to_csv());
//! println!("{}", doc.to_markdown());
//! ```

pub mod cell;
pub mod date;
pub mod edit;
pub mod error;
pub mod shared_strings;
pub mod styles;
pub mod text;
pub mod workbook;
pub mod worksheet;
pub mod write;

pub use cell::{Cell, CellRef, CellValue};
pub use date::DateTimeValue;
pub use error::{Result, XlsxError};
pub use shared_strings::SharedStringTable;
pub use styles::StyleSheet;
pub use workbook::{SheetState, WorkbookInfo};
pub use worksheet::{HyperlinkInfo, HyperlinkTarget, Worksheet};

use std::fs::File;
use std::io::{Read, Seek};
use std::path::Path;

use log::debug;
use zip::read::ZipArchive;

use crate::core::opc::{self, OpcReader};
use crate::core::relationships::{Relationships, rel_types};
use crate::core::theme::Theme;

/// A parsed XLSX document.
#[derive(Debug, Clone)]
pub struct XlsxDocument {
    pub workbook: WorkbookInfo,
    pub worksheets: Vec<Worksheet>,
    pub shared_strings: SharedStringTable,
    pub styles: Option<StyleSheet>,
    pub theme: Option<Theme>,
    // Raw bytes for lazy parsing (None after parsing or if not present)
    styles_data: Option<Vec<u8>>,
    theme_data: Option<Vec<u8>>,
}

impl XlsxDocument {
    /// Parse and cache styles on demand. Returns the stylesheet if available.
    pub fn ensure_styles(&mut self) -> Option<&StyleSheet> {
        if self.styles.is_none() {
            if let Some(data) = self.styles_data.take() {
                self.styles = StyleSheet::parse(&data).ok();
            }
        }
        self.styles.as_ref()
    }

    /// Parse and cache theme on demand. Returns the theme if available.
    pub fn ensure_theme(&mut self) -> Option<&Theme> {
        if self.theme.is_none() {
            if let Some(data) = self.theme_data.take() {
                self.theme = Theme::parse(&data).ok();
            }
        }
        self.theme.as_ref()
    }
}

impl XlsxDocument {
    /// Open an XLSX file from a file path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = File::open(path).map_err(crate::core::Error::from)?;
        let archive = ZipArchive::new(file).map_err(crate::core::Error::from)?;
        Self::from_zip(archive)
    }

    /// Open an XLSX file using memory-mapped I/O for better performance on large files.
    #[cfg(feature = "mmap")]
    pub fn open_mmap(path: impl AsRef<Path>) -> Result<Self> {
        let file = File::open(path).map_err(crate::core::Error::from)?;
        let mmap = unsafe { memmap2::Mmap::map(&file).map_err(crate::core::Error::from)? };
        debug!("XLSX fast path: mmap opened ({} bytes)", mmap.len());
        let archive =
            ZipArchive::new(std::io::Cursor::new(mmap)).map_err(crate::core::Error::from)?;
        Self::from_zip(archive)
    }

    /// Open an XLSX document from any `Read + Seek` source.
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let archive = ZipArchive::new(reader).map_err(crate::core::Error::from)?;
        Self::from_zip(archive)
    }

    /// Read a ZIP entry by name with UTF-8 transcoding for XML parts.
    fn read_xml_entry<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
        name: &str,
    ) -> std::result::Result<Vec<u8>, crate::core::Error> {
        let data = opc::read_zip_entry(archive, name)?;
        if name.ends_with(".xml") || name.ends_with(".rels") {
            if let Some(utf8_data) = crate::core::xml::ensure_utf8(&data) {
                return Ok(utf8_data);
            }
        }
        Ok(data)
    }

    /// Fast path: open ZIP directly and read XLSX parts by known paths,
    /// bypassing OPC content-types and package-level relationships.
    fn from_zip<R: Read + Seek>(mut archive: ZipArchive<R>) -> Result<Self> {
        debug!("XlsxDocument: fast path parsing started ({} ZIP entries)", archive.len());

        // Read workbook relationships to resolve sheet targets
        let wb_rels = match Self::read_xml_entry(&mut archive, "xl/_rels/workbook.xml.rels") {
            Ok(data) => Relationships::parse(&data)?,
            Err(_) => Relationships::empty(),
        };

        // Parse shared strings (must be first — cells reference by index)
        let shared_strings = match Self::read_xml_entry(&mut archive, "xl/sharedStrings.xml") {
            Ok(data) => SharedStringTable::parse(&data)?,
            Err(_) => SharedStringTable::empty(),
        };

        // Parse styles eagerly — needed for date detection in format_cell_value().
        let styles = match Self::read_xml_entry(&mut archive, "xl/styles.xml") {
            Ok(data) => StyleSheet::parse(&data).ok(),
            Err(_) => None,
        };

        // Read theme data lazily
        let theme_data = Self::read_xml_entry(&mut archive, "xl/theme/theme1.xml").ok();

        // Parse workbook
        let wb_data = Self::read_xml_entry(&mut archive, "xl/workbook.xml")?;
        let workbook = WorkbookInfo::parse(&wb_data)?;

        // Phase 1: gather raw data sequentially (requires &mut archive)
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

            // Resolve sheet path from relationships, or fall back to convention
            let sheet_path = if let Some(rel) = wb_rels.get_by_id(&sheet.rel_id) {
                // rel.target is typically "worksheets/sheet1.xml" (relative to xl/)
                let target = &rel.target;
                if let Some(stripped) = target.strip_prefix('/') {
                    // Absolute path — strip leading slash for ZIP entry name
                    stripped.to_string()
                } else {
                    format!("xl/{}", target)
                }
            } else {
                // Fallback: guess by index
                let idx = bundles.len() + 1;
                format!("xl/worksheets/sheet{}.xml", idx)
            };

            let ws_data = match Self::read_xml_entry(&mut archive, &sheet_path) {
                Ok(data) => data,
                Err(_) => {
                    // Try alternate index-based name
                    let idx = bundles.len() + 1;
                    let alt = format!("xl/worksheets/sheet{}.xml", idx);
                    match Self::read_xml_entry(&mut archive, &alt) {
                        Ok(data) => data,
                        Err(_) => continue,
                    }
                },
            };

            // Read worksheet relationships (for hyperlinks)
            let rels_path = sheet_rels_path(&sheet_path);
            let ws_rels = match Self::read_xml_entry(&mut archive, &rels_path) {
                Ok(data) => Relationships::parse(&data).unwrap_or_else(|_| Relationships::empty()),
                Err(_) => Relationships::empty(),
            };

            bundles.push(SheetBundle {
                name: sheet.name.clone(),
                data: ws_data,
                rels: ws_rels,
            });
        }

        // Phase 2: parse worksheets (parallel when feature enabled)
        let worksheets = crate::core::parallel::map_collect(bundles, |b| -> Result<Worksheet> {
            let ws = Worksheet::parse(&b.data, b.name, &b.rels)?;
            Ok(ws)
        })?;

        debug!("XlsxDocument: {} worksheets parsed", worksheets.len());
        Ok(XlsxDocument {
            workbook,
            worksheets,
            shared_strings,
            styles,
            theme: None,
            styles_data: None,
            theme_data,
        })
    }

    /// OPC-based fallback path (used by the unified office_oxide crate when OPC is needed).
    #[allow(dead_code)]
    pub(crate) fn from_opc<R: Read + Seek>(mut opc: OpcReader<R>) -> Result<Self> {
        debug!("XlsxDocument: OPC parsing started");
        let main_part = opc.main_document_part()?;
        let wb_rels = opc.read_rels_for(&main_part)?;

        let shared_strings = if let Some(rel) = wb_rels.first_by_type(rel_types::SHARED_STRINGS) {
            let part_name = main_part.resolve_relative(&rel.target)?;
            let data = opc.read_part(&part_name)?;
            SharedStringTable::parse(&data)?
        } else {
            SharedStringTable::empty()
        };

        let theme_data = if let Some(rel) = wb_rels.first_by_type(rel_types::THEME) {
            let part_name = main_part.resolve_relative(&rel.target)?;
            opc.read_part(&part_name).ok()
        } else {
            None
        };

        let styles = if let Some(rel) = wb_rels.first_by_type(rel_types::STYLES) {
            let part_name = main_part.resolve_relative(&rel.target)?;
            let data = opc.read_part(&part_name)?;
            Some(StyleSheet::parse(&data)?)
        } else {
            None
        };

        let wb_data = opc.read_part(&main_part)?;
        let workbook = WorkbookInfo::parse(&wb_data)?;

        struct SheetBundle {
            name: String,
            data: Vec<u8>,
            rels: Relationships,
        }
        let mut bundles = Vec::with_capacity(workbook.sheets.len());
        for sheet in &workbook.sheets {
            if sheet.rel_id.is_empty() {
                continue;
            }
            let part_name = match wb_rels.resolve_target(&sheet.rel_id, &main_part) {
                Ok(pn) => pn,
                Err(_) => {
                    let idx = bundles.len() + 1;
                    let candidates = [
                        format!("/xl/worksheets/sheet{}.xml", idx),
                        format!("/xl/worksheets/sheet{}.xml", sheet.sheet_id),
                    ];
                    match candidates.iter().find_map(|c| {
                        crate::core::opc::PartName::new(c)
                            .ok()
                            .filter(|pn| opc.has_part(pn))
                    }) {
                        Some(pn) => {
                            debug!("worksheet fallback: '{}' -> '{}'", sheet.rel_id, pn);
                            pn
                        },
                        None => continue,
                    }
                },
            };
            let ws_rels = opc
                .read_rels_for(&part_name)
                .unwrap_or_else(|_| Relationships::empty());
            let ws_data = match opc.read_part(&part_name) {
                Ok(data) => data,
                Err(_) => continue,
            };
            bundles.push(SheetBundle {
                name: sheet.name.clone(),
                data: ws_data,
                rels: ws_rels,
            });
        }

        #[cfg(feature = "parallel")]
        let worksheets: Result<Vec<Worksheet>> = {
            use rayon::prelude::*;
            bundles
                .into_par_iter()
                .map(|b| {
                    let ws = Worksheet::parse(&b.data, b.name, &b.rels)?;
                    Ok(ws)
                })
                .collect()
        };
        #[cfg(not(feature = "parallel"))]
        let worksheets: Result<Vec<Worksheet>> = bundles
            .into_iter()
            .map(|b| {
                let ws = Worksheet::parse(&b.data, b.name, &b.rels)?;
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
            theme: None,
            styles_data: None,
            theme_data,
        })
    }
}

/// Compute the .rels path for a worksheet ZIP entry.
/// e.g. "xl/worksheets/sheet1.xml" → "xl/worksheets/_rels/sheet1.xml.rels"
fn sheet_rels_path(sheet_path: &str) -> String {
    if let Some(pos) = sheet_path.rfind('/') {
        let dir = &sheet_path[..pos];
        let file = &sheet_path[pos + 1..];
        format!("{}/_rels/{}.rels", dir, file)
    } else {
        format!("_rels/{}.rels", sheet_path)
    }
}

impl crate::core::OfficeDocument for XlsxDocument {
    fn plain_text(&self) -> String {
        self.plain_text()
    }

    fn to_markdown(&self) -> String {
        self.to_markdown()
    }
}
