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

/// Cell value types and cell reference types.
pub mod cell;
/// Date/time serial number conversion for XLSX dates.
pub mod date;
/// In-place editing of existing XLSX files.
pub mod edit;
/// XLSX-specific error type.
pub mod error;
/// Number format rendering: apply Excel format strings to numeric values.
pub mod numfmt;
/// Shared string table (SST) parsing and lookup.
pub mod shared_strings;
/// Spreadsheet styles: number formats, fonts, fills, borders, cell formats.
pub mod styles;
/// Text extraction and markdown/CSV rendering for XLSX.
pub mod text;
/// Workbook-level metadata and sheet list.
pub mod workbook;
/// Worksheet parsing: cells, dimensions, hyperlinks.
pub mod worksheet;
/// XLSX creation (write) API.
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
    /// Workbook-level metadata (name, sheets list, date system).
    pub workbook: WorkbookInfo,
    /// Parsed worksheets in sheet-order.
    pub worksheets: Vec<Worksheet>,
    /// Shared string table.
    pub shared_strings: SharedStringTable,
    /// Stylesheet (lazily parsed; access via `ensure_styles()`).
    pub styles: Option<StyleSheet>,
    /// DrawingML theme (lazily parsed; access via `ensure_theme()`).
    pub theme: Option<Theme>,
    /// Text content extracted from `xl/charts/chart*.xml` parts. Each entry
    /// is the flattened text (titles, axis labels, series names, category
    /// labels, values) of one chart in document order. We don't render
    /// charts as graphics but keeping their text content lets it appear in
    /// extracted text and downstream conversions.
    pub chart_text: Vec<String>,
    /// Font programs found under `xl/fonts/`. Each entry is
    /// `(font_name, ttf_or_otf_bytes)`. Mirrors `DocxDocument` and
    /// `PptxDocument`. PDF→XLSX→PDF round-trips ship source fonts
    /// here so the round-trip can re-register them with the PDF
    /// renderer; without this hop XLSX-mediated round-trips lost
    /// every typeface to the base 14 fallback.
    pub embedded_fonts: Vec<(String, Vec<u8>)>,
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
            images: Vec<crate::xlsx::worksheet::WorksheetPicture>,
            text_shapes: Vec<crate::xlsx::worksheet::WorksheetTextShape>,
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

            // Resolve the worksheet's DRAWING rel up-front (Phase 1
            // has access to &mut archive). Each entry decodes
            // `<xdr:pic>` and `<xdr:sp>` anchors and the underlying
            // media bytes so Phase 2's parallel parser doesn't need
            // the archive.
            let (images, text_shapes) = read_drawing_for_sheet(&mut archive, &sheet_path, &ws_rels);

            bundles.push(SheetBundle {
                name: sheet.name.clone(),
                data: ws_data,
                rels: ws_rels,
                images,
                text_shapes,
            });
        }

        // Phase 2: parse worksheets (parallel when feature enabled)
        let worksheets = crate::core::parallel::map_collect(bundles, |b| -> Result<Worksheet> {
            let mut ws = Worksheet::parse(&b.data, b.name, &b.rels)?;
            ws.images = b.images;
            ws.text_shapes = b.text_shapes;
            Ok(ws)
        })?;

        // Scan for chart XML parts (xl/charts/chart*.xml) and extract their
        // visible text — title, axis titles, series names, category labels,
        // cached values. We don't render charts as graphics but their words
        // belong in any text-based downstream conversion (markdown, search
        // indexes, accessibility readers, our PDF text fallback).
        let mut chart_text: Vec<String> = Vec::new();
        let chart_names: Vec<String> = (0..archive.len())
            .filter_map(|i| archive.by_index(i).ok().map(|f| f.name().to_string()))
            .filter(|n| n.starts_with("xl/charts/chart") && n.ends_with(".xml"))
            .collect();
        for name in chart_names {
            if let Ok(data) = Self::read_xml_entry(&mut archive, &name) {
                let text = extract_chart_text(&data);
                if !text.is_empty() {
                    chart_text.push(text);
                }
            }
        }

        // Scan `xl/fonts/` for embedded font programs. Mirrors the
        // DOCX (`word/fonts/`) and PPTX (`ppt/fonts/`) readers.
        let mut embedded_fonts: Vec<(String, Vec<u8>)> = Vec::new();
        let font_names: Vec<String> = (0..archive.len())
            .filter_map(|i| archive.by_index(i).ok().map(|f| f.name().to_string()))
            .filter(|n| {
                n.starts_with("xl/fonts/")
                    && (n.to_lowercase().ends_with(".ttf") || n.to_lowercase().ends_with(".otf"))
            })
            .collect();
        for name in font_names {
            if let Ok(data) = opc::read_zip_entry(&mut archive, &name) {
                let basename = name.rsplit('/').next().unwrap_or("font");
                let face = crate::docx::strip_embedded_font_filename(basename);
                let font_name = if face.is_empty() {
                    basename.to_string()
                } else {
                    face
                };
                embedded_fonts.push((font_name, data));
            }
        }

        debug!(
            "XlsxDocument: {} worksheets parsed, {} chart(s), {} embedded fonts",
            worksheets.len(),
            chart_text.len(),
            embedded_fonts.len()
        );
        Ok(XlsxDocument {
            workbook,
            worksheets,
            shared_strings,
            styles,
            theme: None,
            chart_text,
            embedded_fonts,
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

        // Mirror the zip-path embedded-fonts scan over the OPC part
        // listing. Loading via OPC is the slow path used when a
        // caller hands us a pre-built `OpcReader`, so duplicating
        // the cheap scan keeps font fidelity working there too.
        let mut embedded_fonts: Vec<(String, Vec<u8>)> = Vec::new();
        for name in opc.part_names() {
            let s = name.to_string();
            if !s.starts_with("/xl/fonts/") {
                continue;
            }
            let lower = s.to_lowercase();
            if !(lower.ends_with(".ttf") || lower.ends_with(".otf")) {
                continue;
            }
            if let Ok(data) = opc.read_part(&name) {
                let basename = s.rsplit('/').next().unwrap_or("font");
                let face = crate::docx::strip_embedded_font_filename(basename);
                let font_name = if face.is_empty() {
                    basename.to_string()
                } else {
                    face
                };
                embedded_fonts.push((font_name, data));
            }
        }

        debug!(
            "XlsxDocument: {} worksheets parsed (OPC path), {} embedded fonts",
            worksheets.len(),
            embedded_fonts.len()
        );
        Ok(XlsxDocument {
            workbook,
            worksheets,
            shared_strings,
            styles,
            theme: None,
            // OPC path doesn't extract chart text yet; the zip path is the
            // hot one used by Document::from_reader. Charts via OPC can be
            // added if a use case appears.
            chart_text: Vec::new(),
            embedded_fonts,
            styles_data: None,
            theme_data,
        })
    }
}

/// Extract structured content from a chart XML stream (DrawingML chart
/// format) into a flat textual representation.
///
/// Walks the chart's title (`<c:title>`), axis titles (`<c:catAx>` /
/// `<c:valAx>` / `<c:title>`), and each series (`<c:ser>`). For every
/// series we capture the name (`<c:tx>`), category labels (`<c:cat>`),
/// and cached numeric values (`<c:val>`). The output groups them into
/// readable lines that include the **structure** of the chart — series
/// names paired with their values per category — rather than the flat
/// soup of `<a:t>`/`<c:v>` text the previous implementation produced.
///
/// Output shape:
/// ```text
/// Title: ...
/// Categories: A, B, C, ...
/// Series Budget: 1690, 2100, 1570, ...
/// Series Projected: 1310, 3480, 510, ...
/// ```
///
/// This still travels through `to_markdown` and `convert_xlsx_to_ir` as
/// plain text (not an actual table), but the structure is now meaningful
/// for both human readers and downstream NLP / search.
fn extract_chart_text(xml: &[u8]) -> String {
    let mut reader = quick_xml::Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    // Tag-context stack — push localname on Start, pop on End.
    let mut stack: Vec<Vec<u8>> = Vec::new();
    // Most recently seen text inside a `<t>` (rich-text run) — used to
    // build the chart title and axis-title strings.
    let mut current_title: String = String::new();
    let mut titles: Vec<String> = Vec::new();
    // The chart-level title is the first `<c:title>` we close that lives
    // outside any `<c:catAx>` / `<c:valAx>` / `<c:legend>`.
    // Per-series state.
    let mut series: Vec<ChartSeries> = Vec::new();
    let mut cur_series: Option<ChartSeries> = None;
    // Current `<c:v>` text being accumulated.
    let mut cur_v: String = String::new();
    // Categories from the current series (or the first series — they are
    // typically shared across all series in the chart).
    let mut shared_categories: Vec<String> = Vec::new();
    let mut cur_cat_buf: Vec<String> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(e)) => {
                let local = e.local_name().as_ref().to_vec();
                if local == b"ser" {
                    cur_series = Some(ChartSeries::default());
                    cur_cat_buf.clear();
                }
                stack.push(local);
            },
            Ok(quick_xml::events::Event::End(e)) => {
                let local = e.local_name().as_ref().to_vec();
                let _ = stack.pop();
                match local.as_slice() {
                    b"t" => {
                        // End of a rich-text run — accumulate into current_title
                        // if we're inside a chart-level or axis title.
                    },
                    b"title" => {
                        if !current_title.trim().is_empty() {
                            titles.push(current_title.trim().to_string());
                        }
                        current_title.clear();
                    },
                    b"v" => {
                        let val = cur_v.trim().to_string();
                        cur_v.clear();
                        if val.is_empty() {
                            continue;
                        }
                        if let Some(s) = cur_series.as_mut() {
                            // Decide whether this <c:v> is series-name, category,
                            // or value based on the enclosing scope.
                            let in_tx = stack.iter().any(|t| t.as_slice() == b"tx");
                            let in_cat = stack.iter().any(|t| t.as_slice() == b"cat");
                            let in_val = stack.iter().any(|t| t.as_slice() == b"val");
                            if in_tx && s.name.is_empty() {
                                s.name = val;
                            } else if in_cat {
                                cur_cat_buf.push(val);
                            } else if in_val {
                                s.values.push(val);
                            }
                        }
                    },
                    b"ser" => {
                        if let Some(mut s) = cur_series.take() {
                            // Fold the per-series categories into shared_categories
                            // (first series wins — they are typically identical).
                            if shared_categories.is_empty() && !cur_cat_buf.is_empty() {
                                shared_categories = std::mem::take(&mut cur_cat_buf);
                            } else {
                                cur_cat_buf.clear();
                            }
                            if s.name.is_empty() {
                                s.name = format!("Series {}", series.len() + 1);
                            }
                            series.push(s);
                        }
                    },
                    _ => {},
                }
            },
            Ok(quick_xml::events::Event::Text(t)) => {
                if let Ok(s) = t.unescape() {
                    let trimmed = s.trim();
                    if trimmed.is_empty() {
                        continue;
                    }
                    let top = stack.last().map(|v| v.as_slice());
                    match top {
                        Some(b"t") => {
                            // Rich-text run — append to current_title.
                            if !current_title.is_empty() {
                                current_title.push_str("");
                            }
                            current_title.push_str(trimmed);
                        },
                        Some(b"v") => {
                            cur_v.push_str(trimmed);
                        },
                        _ => {},
                    }
                }
            },
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {},
        }
        buf.clear();
    }

    // Emit a structured representation. Each line is independent — the
    // markdown writer joins them with `\n`.
    let mut out = String::new();
    if !titles.is_empty() {
        out.push_str(&format!("Title: {}", titles.join(" — ")));
    }
    if !shared_categories.is_empty() {
        if !out.is_empty() {
            out.push('\n');
        }
        out.push_str(&format!("Categories: {}", shared_categories.join(", ")));
    }
    for s in &series {
        if !out.is_empty() {
            out.push('\n');
        }
        if s.values.is_empty() {
            out.push_str(&format!("Series: {}", s.name));
        } else {
            out.push_str(&format!("{}: {}", s.name, s.values.join(", ")));
        }
    }
    out
}

#[derive(Default)]
struct ChartSeries {
    name: String,
    values: Vec<String>,
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

/// Read the DRAWING-rel target for a worksheet, parse its `<xdr:pic>`
/// and `<xdr:sp>` anchors, and resolve each picture's underlying media
/// bytes. Returns `(pictures, text_shapes)`. Soft failures (no rel,
/// missing part, parse error) yield empty vectors — drawings are
/// best-effort extras and shouldn't fail worksheet loading.
fn read_drawing_for_sheet<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    sheet_path: &str,
    sheet_rels: &Relationships,
) -> (
    Vec<crate::xlsx::worksheet::WorksheetPicture>,
    Vec<crate::xlsx::worksheet::WorksheetTextShape>,
) {
    let drawing_rel = match sheet_rels.first_by_type(rel_types::DRAWING) {
        Some(r) => r,
        None => return (Vec::new(), Vec::new()),
    };

    let drawing_path = resolve_relative_zip_path(sheet_path, &drawing_rel.target);

    let drawing_xml = match XlsxDocument::read_xml_entry(archive, &drawing_path) {
        Ok(d) => d,
        Err(_) => return (Vec::new(), Vec::new()),
    };

    let drawing_rels_path = sheet_rels_path(&drawing_path);
    let drawing_rels = match XlsxDocument::read_xml_entry(archive, &drawing_rels_path) {
        Ok(d) => Relationships::parse(&d).unwrap_or_else(|_| Relationships::empty()),
        Err(_) => Relationships::empty(),
    };

    let parsed = match parse_drawing_anchors(&drawing_xml) {
        Ok(a) => a,
        Err(_) => return (Vec::new(), Vec::new()),
    };

    // Resolve picture anchors → bytes.
    let mut pictures = Vec::with_capacity(parsed.pictures.len());
    for a in parsed.pictures {
        let rel = match drawing_rels.get_by_id(&a.embed_rid) {
            Some(r) => r,
            None => continue,
        };
        let media_path = resolve_relative_zip_path(&drawing_path, &rel.target);
        let bytes = match opc::read_zip_entry(archive, &media_path) {
            Ok(b) => b,
            Err(_) => continue,
        };
        let ext = std::path::Path::new(&rel.target)
            .extension()
            .and_then(|s| s.to_str())
            .map(|s| s.to_ascii_lowercase())
            .unwrap_or_else(|| guess_image_format_from_bytes(&bytes).to_string());

        pictures.push(crate::xlsx::worksheet::WorksheetPicture {
            data: bytes,
            format: ext,
            x_emu: a.x_emu,
            y_emu: a.y_emu,
            cx_emu: a.cx_emu,
            cy_emu: a.cy_emu,
            alt_text: a.alt_text,
        });
    }

    let text_shapes = parsed
        .text_shapes
        .into_iter()
        .map(|t| crate::xlsx::worksheet::WorksheetTextShape {
            text: t.text,
            font_name: t.font_name,
            font_size_pt: t.font_size_pt,
            bold: t.bold,
            italic: t.italic,
            color_hex: t.color_hex,
            x_emu: t.x_emu,
            y_emu: t.y_emu,
            cx_emu: t.cx_emu,
            cy_emu: t.cy_emu,
        })
        .collect();

    (pictures, text_shapes)
}

/// Resolve a `..`-relative target inside an OPC package back to an
/// absolute ZIP-entry path. Mirrors `PartName::resolve_relative` but
/// operates on plain ZIP paths (the `from_zip` fast path doesn't use
/// `PartName`).
fn resolve_relative_zip_path(source: &str, target: &str) -> String {
    if target.starts_with('/') {
        return target.trim_start_matches('/').to_string();
    }
    let base_dir = match source.rfind('/') {
        Some(i) => &source[..i],
        None => "",
    };
    let mut parts: Vec<&str> = if base_dir.is_empty() {
        Vec::new()
    } else {
        base_dir.split('/').collect()
    };
    for seg in target.split('/') {
        match seg {
            "" | "." => {},
            ".." => {
                parts.pop();
            },
            other => parts.push(other),
        }
    }
    parts.join("/")
}

#[derive(Debug)]
struct DrawingPictureAnchor {
    embed_rid: String,
    x_emu: i64,
    y_emu: i64,
    cx_emu: i64,
    cy_emu: i64,
    alt_text: Option<String>,
}

#[derive(Debug, Default)]
struct DrawingTextAnchor {
    text: String,
    font_name: Option<String>,
    font_size_pt: Option<f32>,
    bold: bool,
    italic: bool,
    color_hex: Option<String>,
    x_emu: i64,
    y_emu: i64,
    cx_emu: i64,
    cy_emu: i64,
}

#[derive(Debug, Default)]
struct DrawingAnchors {
    pictures: Vec<DrawingPictureAnchor>,
    text_shapes: Vec<DrawingTextAnchor>,
}

/// Parse `xl/drawings/drawingN.xml` and return both `<xdr:pic>` and
/// `<xdr:sp>` anchors. Supports `<xdr:absoluteAnchor>` (direct EMU
/// pos+ext) and the cell-anchor variants — for cell anchors we
/// approximate the absolute origin from `<xdr:from>` x/y when present.
/// `<xdr:sp>` shapes carry text inside `<xdr:txBody>` runs.
fn parse_drawing_anchors(xml_data: &[u8]) -> crate::core::Result<DrawingAnchors> {
    use quick_xml::events::Event;

    let mut reader = crate::core::xml::make_fast_reader(xml_data);
    let mut out = DrawingAnchors::default();

    // Per-anchor accumulator state. We don't pre-classify the anchor
    // as picture-vs-text; we discover that mid-walk based on which
    // child element appears (`pic` vs `sp`).
    enum AnchorKind {
        Unknown,
        Picture,
        Text,
    }
    let mut in_anchor = false;
    let mut kind = AnchorKind::Unknown;
    let mut x_emu = 0i64;
    let mut y_emu = 0i64;
    let mut cx_emu = 0i64;
    let mut cy_emu = 0i64;
    let mut embed_rid: Option<String> = None;
    let mut alt_text: Option<String> = None;
    // Text-shape state.
    let mut in_txbody = false;
    let mut in_run = false;
    let mut in_a_t = false;
    let mut text_buf = String::new();
    let mut font_name: Option<String> = None;
    let mut font_size_pt: Option<f32> = None;
    let mut bold = false;
    let mut italic = false;
    let mut color_hex: Option<String> = None;
    let mut in_solid_fill = false;

    loop {
        let evt = reader.read_event()?;
        match evt {
            Event::Start(ref e) => {
                let local = e.local_name().as_ref().to_vec();
                match local.as_slice() {
                    b"absoluteAnchor" | b"oneCellAnchor" | b"twoCellAnchor" => {
                        in_anchor = true;
                        kind = AnchorKind::Unknown;
                        x_emu = 0;
                        y_emu = 0;
                        cx_emu = 0;
                        cy_emu = 0;
                        embed_rid = None;
                        alt_text = None;
                        in_txbody = false;
                        in_run = false;
                        in_a_t = false;
                        text_buf.clear();
                        font_name = None;
                        font_size_pt = None;
                        bold = false;
                        italic = false;
                        color_hex = None;
                        in_solid_fill = false;
                    },
                    b"pic" if in_anchor => {
                        kind = AnchorKind::Picture;
                    },
                    b"sp" if in_anchor => {
                        kind = AnchorKind::Text;
                    },
                    b"txBody" if in_anchor => {
                        in_txbody = true;
                    },
                    b"r" if in_txbody => {
                        in_run = true;
                    },
                    b"t" if in_run => {
                        in_a_t = true;
                    },
                    b"rPr" if in_run => {
                        for attr in e.attributes().with_checks(false) {
                            let attr = attr.map_err(crate::core::Error::from)?;
                            let key = attr.key.as_ref();
                            let raw = attr.unescape_value().map_err(crate::core::Error::from)?;
                            match key {
                                b"sz" => {
                                    // sz is in hundredths of a pt.
                                    if let Ok(n) = raw.parse::<i32>() {
                                        font_size_pt = Some(n as f32 / 100.0);
                                    }
                                },
                                b"b" => bold = raw == "1" || raw == "true",
                                b"i" => italic = raw == "1" || raw == "true",
                                _ => {},
                            }
                        }
                    },
                    b"solidFill" if in_run => {
                        in_solid_fill = true;
                    },
                    b"cNvPr" if in_anchor => {
                        if let Some(d) = crate::core::xml::optional_attr_str(e, b"descr")? {
                            alt_text = Some(d.into_owned());
                        }
                    },
                    _ => {},
                }
            },
            Event::Empty(ref e) => {
                if !in_anchor {
                    continue;
                }
                let local = e.local_name().as_ref().to_vec();
                match local.as_slice() {
                    b"pos" => {
                        if let Some(v) = crate::core::xml::optional_attr_str(e, b"x")? {
                            x_emu = v.parse().unwrap_or(0);
                        }
                        if let Some(v) = crate::core::xml::optional_attr_str(e, b"y")? {
                            y_emu = v.parse().unwrap_or(0);
                        }
                    },
                    b"ext" => {
                        if let Some(v) = crate::core::xml::optional_attr_str(e, b"cx")? {
                            cx_emu = v.parse().unwrap_or(0);
                        }
                        if let Some(v) = crate::core::xml::optional_attr_str(e, b"cy")? {
                            cy_emu = v.parse().unwrap_or(0);
                        }
                    },
                    b"off" if cx_emu == 0 && cy_emu == 0 => {
                        if let Some(v) = crate::core::xml::optional_attr_str(e, b"x")? {
                            x_emu = v.parse().unwrap_or(x_emu);
                        }
                        if let Some(v) = crate::core::xml::optional_attr_str(e, b"y")? {
                            y_emu = v.parse().unwrap_or(y_emu);
                        }
                    },
                    b"blip" => {
                        for attr in e.attributes().with_checks(false) {
                            let attr = attr.map_err(crate::core::Error::from)?;
                            let key = attr.key.as_ref();
                            if key == b"r:embed" || key.ends_with(b":embed") || key == b"embed" {
                                let raw =
                                    attr.unescape_value().map_err(crate::core::Error::from)?;
                                embed_rid = Some(raw.into_owned());
                                break;
                            }
                        }
                    },
                    b"cNvPr" => {
                        if let Some(d) = crate::core::xml::optional_attr_str(e, b"descr")? {
                            alt_text = Some(d.into_owned());
                        }
                    },
                    b"latin" if in_run => {
                        if let Some(t) = crate::core::xml::optional_attr_str(e, b"typeface")? {
                            font_name = Some(t.into_owned());
                        }
                    },
                    b"srgbClr" if in_solid_fill => {
                        if let Some(v) = crate::core::xml::optional_attr_str(e, b"val")? {
                            color_hex = Some(v.into_owned().to_uppercase());
                        }
                    },
                    b"rPr" if in_run => {
                        for attr in e.attributes().with_checks(false) {
                            let attr = attr.map_err(crate::core::Error::from)?;
                            let key = attr.key.as_ref();
                            let raw = attr.unescape_value().map_err(crate::core::Error::from)?;
                            match key {
                                b"sz" => {
                                    if let Ok(n) = raw.parse::<i32>() {
                                        font_size_pt = Some(n as f32 / 100.0);
                                    }
                                },
                                b"b" => bold = raw == "1" || raw == "true",
                                b"i" => italic = raw == "1" || raw == "true",
                                _ => {},
                            }
                        }
                    },
                    _ => {},
                }
            },
            Event::Text(ref e) if in_a_t => {
                let s = e.unescape().map_err(crate::core::Error::from)?;
                text_buf.push_str(&s);
            },
            Event::End(ref e) => {
                let local = e.local_name().as_ref().to_vec();
                match local.as_slice() {
                    b"t" => in_a_t = false,
                    b"r" => in_run = false,
                    b"txBody" => in_txbody = false,
                    b"solidFill" => in_solid_fill = false,
                    s if matches!(s, b"absoluteAnchor" | b"oneCellAnchor" | b"twoCellAnchor")
                        && in_anchor =>
                    {
                        in_anchor = false;
                        match kind {
                            AnchorKind::Picture => {
                                if let Some(rid) = embed_rid.take() {
                                    out.pictures.push(DrawingPictureAnchor {
                                        embed_rid: rid,
                                        x_emu,
                                        y_emu,
                                        cx_emu,
                                        cy_emu,
                                        alt_text: alt_text.take(),
                                    });
                                }
                            },
                            AnchorKind::Text => {
                                if !text_buf.is_empty() {
                                    out.text_shapes.push(DrawingTextAnchor {
                                        text: std::mem::take(&mut text_buf),
                                        font_name: font_name.take(),
                                        font_size_pt: font_size_pt.take(),
                                        bold,
                                        italic,
                                        color_hex: color_hex.take(),
                                        x_emu,
                                        y_emu,
                                        cx_emu,
                                        cy_emu,
                                    });
                                }
                            },
                            AnchorKind::Unknown => {},
                        }
                        kind = AnchorKind::Unknown;
                    },
                    _ => {},
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(out)
}

/// Best-effort image-format detection from raw bytes (used when the
/// drawing rel target lacks a recognisable extension). Mirrors the
/// PPTX helper.
fn guess_image_format_from_bytes(bytes: &[u8]) -> &'static str {
    if bytes.starts_with(&[0x89, b'P', b'N', b'G']) {
        "png"
    } else if bytes.starts_with(&[0xFF, 0xD8, 0xFF]) {
        "jpeg"
    } else if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        "gif"
    } else if bytes.starts_with(b"BM") {
        "bmp"
    } else if bytes.len() >= 4 && (bytes.starts_with(b"II*\0") || bytes.starts_with(b"MM\0*")) {
        "tiff"
    } else if bytes.len() >= 4 && bytes.starts_with(&[0xD7, 0xCD, 0xC6, 0x9A]) {
        "wmf"
    } else if bytes.len() >= 4 && bytes.starts_with(&[0x01, 0x00, 0x00, 0x00]) {
        "emf"
    } else {
        "png"
    }
}
