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
/// Shared-formula (`<f t="shared">`) group expansion.
pub mod shared_formula;
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
mod xlsb;

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

use crate::core::opc;
use crate::core::relationships::{Relationships, rel_types};
use crate::core::theme::Theme;
use crate::core::xml;

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
    /// Parsed `docProps/core.xml`. `None` when the package carries no
    /// core-properties part.
    pub core_properties: Option<crate::core::properties::CoreProperties>,
    /// Parsed `docProps/app.xml` (company, producing application, template,
    /// page/word/character/paragraph counts). `None` when the package
    /// carries no extended-properties part.
    pub app_properties: Option<crate::core::properties::AppProperties>,
    /// `true` when the workbook part's own relationships include a
    /// `vbaProject` entry — a cheap macro-presence signal, no VBA
    /// interpretation.
    pub has_macros: bool,
    /// Sheets whose part could not be parsed — `(sheet name, error)`. A
    /// damaged archive (a bad CRC turning one part to garbage) used to
    /// fail the whole workbook where Tika/POI return the sheets that are
    /// intact; those are returned now, and the loss is on record here,
    /// in every renderer's output and in `Metadata::text_truncated`.
    pub unreadable_sheets: Vec<(String, String)>,
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
        let archive = opc::open_zip_rejecting_duplicates(file)?;
        Self::from_zip(archive)
    }

    /// Open an XLSX file using memory-mapped I/O for better performance on large files.
    #[cfg(feature = "mmap")]
    pub fn open_mmap(path: impl AsRef<Path>) -> Result<Self> {
        let file = File::open(path).map_err(crate::core::Error::from)?;
        let mmap = unsafe { memmap2::Mmap::map(&file).map_err(crate::core::Error::from)? };
        debug!("XLSX fast path: mmap opened ({} bytes)", mmap.len());
        let archive = opc::open_zip_rejecting_duplicates(std::io::Cursor::new(mmap))?;
        Self::from_zip(archive)
    }

    /// Open an XLSX document from any `Read + Seek` source.
    pub fn from_reader<R: Read + Seek>(mut reader: R) -> Result<Self> {
        // A password-protected XLSX is a CFB container, not a zip at all.
        // See the identical check in docx::DocxDocument::from_reader.
        if crate::cfb::is_cfb_container(&mut reader).map_err(crate::core::Error::from)? {
            return Err(crate::core::Error::Unsupported(
                "the file is a password-protected (encrypted) OOXML package; \
                 decryption is not supported"
                    .into(),
            )
            .into());
        }
        let archive = opc::open_zip_rejecting_duplicates(reader)?;
        Self::from_zip(archive)
    }

    /// Read a ZIP entry by name with UTF-8 transcoding for XML parts.
    pub(super) fn read_xml_entry<R: Read + Seek>(
        archive: &mut ZipArchive<R>,
        entries: &opc::ZipEntryIndex,
        name: &str,
    ) -> std::result::Result<Vec<u8>, crate::core::Error> {
        let data = opc::read_zip_entry(archive, entries, name)?;
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
        let entries = opc::ZipEntryIndex::new(&archive);
        // An Excel Binary Workbook is the same package with BIFF12 parts
        // in place of the XML ones; it decodes into this same model.
        if xlsb::is_xlsb(&archive, &entries) {
            return xlsb::from_zip(&mut archive, &entries);
        }

        // Document metadata lives at the conventional path in every package
        // Excel writes; the fast path doesn't consult package relationships,
        // so read it by name here.
        let core_properties = Self::read_xml_entry(&mut archive, &entries, "docProps/core.xml")
            .ok()
            .and_then(|d| crate::core::properties::CoreProperties::parse(&d).ok());
        let app_properties = Self::read_xml_entry(&mut archive, &entries, "docProps/app.xml")
            .ok()
            .and_then(|d| crate::core::properties::AppProperties::parse(&d).ok());

        // Read workbook relationships to resolve sheet targets
        let wb_rels =
            match Self::read_xml_entry(&mut archive, &entries, "xl/_rels/workbook.xml.rels") {
                Ok(data) => Relationships::parse(&data)?,
                Err(_) => Relationships::empty(),
            };
        let has_macros = wb_rels.has_vba_project();

        // Parse shared strings (must be first — cells reference by index)
        let shared_strings =
            match Self::read_xml_entry(&mut archive, &entries, "xl/sharedStrings.xml") {
                Ok(data) => SharedStringTable::parse(&data)?,
                Err(_) => SharedStringTable::empty(),
            };

        // Parse styles eagerly — needed for date detection in format_cell_value().
        let styles = match Self::read_xml_entry(&mut archive, &entries, "xl/styles.xml") {
            Ok(data) => StyleSheet::parse(&data).ok(),
            Err(_) => None,
        };

        // Read theme data lazily
        let theme_data = Self::read_xml_entry(&mut archive, &entries, "xl/theme/theme1.xml").ok();

        // Parse workbook
        let wb_data = Self::read_xml_entry(&mut archive, &entries, "xl/workbook.xml")?;
        let workbook = WorkbookInfo::parse(&wb_data)?;

        // The threaded-comments person list (personId -> display name) is a
        // workbook-level part, not per-sheet.
        let persons = wb_rels
            .first_by_type(rel_types::PERSONS)
            .map(|rel| resolve_relative_zip_path("xl/workbook.xml", &rel.target))
            .or_else(|| Some("xl/persons/person.xml".to_string()))
            .and_then(|path| Self::read_xml_entry(&mut archive, &entries, &path).ok())
            .and_then(|data| worksheet::parse_persons(&data).ok())
            .unwrap_or_default();

        // Phase 1: gather raw data sequentially (requires &mut archive)
        struct SheetBundle {
            name: String,
            data: Vec<u8>,
            rels: Relationships,
            images: Vec<crate::xlsx::worksheet::WorksheetPicture>,
            text_shapes: Vec<crate::xlsx::worksheet::WorksheetTextShape>,
            comments: Vec<crate::xlsx::worksheet::SheetComment>,
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

            let ws_data = match Self::read_xml_entry(&mut archive, &entries, &sheet_path) {
                Ok(data) => data,
                Err(_) => {
                    // Try alternate index-based name
                    let idx = bundles.len() + 1;
                    let alt = format!("xl/worksheets/sheet{}.xml", idx);
                    match Self::read_xml_entry(&mut archive, &entries, &alt) {
                        Ok(data) => data,
                        Err(_) => continue,
                    }
                },
            };

            // Read worksheet relationships (for hyperlinks)
            let rels_path = sheet_rels_path(&sheet_path);
            let ws_rels = match Self::read_xml_entry(&mut archive, &entries, &rels_path) {
                Ok(data) => Relationships::parse(&data).unwrap_or_else(|_| Relationships::empty()),
                Err(_) => Relationships::empty(),
            };

            // Resolve the worksheet's DRAWING rel up-front (Phase 1
            // has access to &mut archive). Each entry decodes
            // `<xdr:pic>` and `<xdr:sp>` anchors and the underlying
            // media bytes so Phase 2's parallel parser doesn't need
            // the archive.
            let (images, text_shapes) =
                read_drawing_for_sheet(&mut archive, &entries, &sheet_path, &ws_rels);

            // Cell comments live in a separate part reached through the
            // sheet's own relationships.
            let comments = ws_rels
                .first_by_type(rel_types::COMMENTS)
                .map(|rel| resolve_relative_zip_path(&sheet_path, &rel.target))
                .and_then(|path| Self::read_xml_entry(&mut archive, &entries, &path).ok())
                .and_then(|data| worksheet::parse_comments(&data).ok())
                .unwrap_or_default();
            // Modern (Excel 2016+) threaded comments, when present,
            // replace the legacy compatibility-boilerplate text for the
            // same cell with the real thread text.
            let threaded_comments = ws_rels
                .first_by_type(rel_types::THREADED_COMMENTS)
                .map(|rel| resolve_relative_zip_path(&sheet_path, &rel.target))
                .and_then(|path| Self::read_xml_entry(&mut archive, &entries, &path).ok())
                .and_then(|data| worksheet::parse_threaded_comments(&data).ok())
                .unwrap_or_default();
            let comments =
                worksheet::merge_threaded_comments(comments, threaded_comments, &persons);

            bundles.push(SheetBundle {
                name: sheet.name.clone(),
                data: ws_data,
                rels: ws_rels,
                images,
                text_shapes,
                comments,
            });
        }

        // Phase 2: parse worksheets (parallel when feature enabled). A
        // sheet whose part does not parse is recorded, not fatal.
        let parsed = crate::core::parallel::map_collect(
            bundles,
            |b| -> Result<std::result::Result<Worksheet, (String, String)>> {
                let name = b.name.clone();
                match Worksheet::parse(&b.data, b.name, &b.rels) {
                    Ok(mut ws) => {
                        ws.images = b.images;
                        ws.text_shapes = b.text_shapes;
                        ws.comments = b.comments;
                        Ok(Ok(ws))
                    },
                    Err(e) => {
                        log::warn!("xlsx: sheet {name:?} is unreadable and skipped: {e}");
                        Ok(Err((name, e.to_string())))
                    },
                }
            },
        )?;
        let mut worksheets = Vec::with_capacity(parsed.len());
        let mut unreadable_sheets = Vec::new();
        for p in parsed {
            match p {
                Ok(ws) => worksheets.push(ws),
                Err(u) => unreadable_sheets.push(u),
            }
        }
        if !unreadable_sheets.is_empty() && worksheets.is_empty() {
            let (name, err) = &unreadable_sheets[0];
            return Err(crate::core::Error::MalformedXml(format!(
                "no readable sheet: {name:?}: {err}"
            ))
            .into());
        }

        // Resolve any in-cell rich-value images: a `vm`-
        // tagged `t="e"` cell whose `vm` maps through the workbook's
        // metadata/richData chain is a real embedded image, not a
        // genuine formula error — swap its fabricated `#VALUE!` text
        // for the actual picture.
        let rich_value_images = read_rich_value_images(&mut archive, &entries);
        if !rich_value_images.is_empty() {
            for ws in &mut worksheets {
                for row in &mut ws.rows {
                    for cell in &mut row.cells {
                        let Some(vm) = cell.vm else { continue };
                        if !matches!(cell.value, CellValue::Error(_)) {
                            continue;
                        }
                        if let Some(pic) = rich_value_images.get(&vm) {
                            ws.images.push(pic.clone());
                            cell.value = CellValue::Empty;
                        }
                    }
                }
            }
        }

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
            if let Ok(data) = Self::read_xml_entry(&mut archive, &entries, &name) {
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
            if let Ok(data) = opc::read_zip_entry(&mut archive, &entries, &name) {
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
            core_properties,
            app_properties,
            has_macros,
            unreadable_sheets,
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
    let mut stack: Vec<String> = Vec::new();
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
                let local = e.local_name().as_ref().to_string();
                if local == "ser" {
                    cur_series = Some(ChartSeries::default());
                    cur_cat_buf.clear();
                }
                stack.push(local);
            },
            Ok(quick_xml::events::Event::End(e)) => {
                let local = e.local_name().as_ref().to_string();
                let _ = stack.pop();
                match local.as_str() {
                    "t" => {
                        // End of a rich-text run — accumulate into current_title
                        // if we're inside a chart-level or axis title.
                    },
                    "title" => {
                        if !current_title.trim().is_empty() {
                            titles.push(current_title.trim().to_string());
                        }
                        current_title.clear();
                    },
                    "v" => {
                        let val = cur_v.trim().to_string();
                        cur_v.clear();
                        if val.is_empty() {
                            continue;
                        }
                        if let Some(s) = cur_series.as_mut() {
                            // Decide whether this <c:v> is series-name, category,
                            // or value based on the enclosing scope.
                            let in_tx = stack.iter().any(|t| t == "tx");
                            // Scatter charts carry their points in
                            // `<c:xVal>`/`<c:yVal>` and bubble charts add
                            // `<c:bubbleSize>` rather than the
                            // `<c:cat>`/`<c:val>` bar/line/pie/area charts
                            // use. Checking only the latter dropped every
                            // scatter/bubble data point.
                            let in_cat = stack.iter().any(|t| matches!(t.as_str(), "cat" | "xVal"));
                            let in_val = stack
                                .iter()
                                .any(|t| matches!(t.as_str(), "val" | "yVal" | "bubbleSize"));
                            if in_tx && s.name.is_empty() {
                                s.name = val;
                            } else if in_cat {
                                cur_cat_buf.push(val);
                            } else if in_val {
                                s.values.push(val);
                            }
                        }
                    },
                    "ser" => {
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
            // Entity references are separate events; folding them in here
            // keeps `&amp;` in a chart title from disappearing.
            Ok(quick_xml::events::Event::GeneralRef(ref r)) => {
                if let Ok(s) = crate::core::xml::resolve_general_ref(r) {
                    let top = stack.last().map(|v| v.as_str());
                    match top {
                        Some("t") => current_title.push_str(&s),
                        Some("v") => cur_v.push_str(&s),
                        _ => {},
                    }
                }
            },
            Ok(quick_xml::events::Event::Text(t)) => {
                if let Ok(s) = crate::core::xml::unescape_text(&t) {
                    let top = stack.last().map(|v| v.as_str());
                    // Appended verbatim. A title Excel split across runs
                    // carries the space *between* two runs as leading or
                    // trailing whitespace on one of them, so trimming each
                    // run before concatenating ran the words together —
                    // "Chart Title - with additional formatting" came back
                    // as "ChartTitle-withadditionalformatting".
                    // The assembled string is trimmed once, where it is
                    // flushed at `</c:title>` / `</c:v>`.
                    match top {
                        Some("t") => {
                            current_title.push_str(&s);
                        },
                        Some("v") => {
                            cur_v.push_str(&s);
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
    entries: &opc::ZipEntryIndex,
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

    let drawing_xml = match XlsxDocument::read_xml_entry(archive, entries, &drawing_path) {
        Ok(d) => d,
        Err(e) => {
            debug!("XlsxDocument: drawing part {} unreadable ({}); skipping", drawing_path, e);
            return (Vec::new(), Vec::new());
        },
    };

    let drawing_rels_path = sheet_rels_path(&drawing_path);
    let drawing_rels = match XlsxDocument::read_xml_entry(archive, entries, &drawing_rels_path) {
        Ok(d) => Relationships::parse(&d).unwrap_or_else(|_| Relationships::empty()),
        Err(_) => Relationships::empty(),
    };

    let parsed = match parse_drawing_anchors(&drawing_xml) {
        Ok(a) => a,
        Err(e) => {
            debug!(
                "XlsxDocument: drawing {} failed to parse ({}); dropping anchors",
                drawing_path, e
            );
            return (Vec::new(), Vec::new());
        },
    };

    // Resolve picture anchors → bytes.
    let mut pictures = Vec::with_capacity(parsed.pictures.len());
    for a in parsed.pictures {
        let rel = match drawing_rels.get_by_id(&a.embed_rid) {
            Some(r) => r,
            None => continue,
        };
        let media_path = resolve_relative_zip_path(&drawing_path, &rel.target);
        let bytes = match opc::read_zip_entry(archive, entries, &media_path) {
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

/// Workbook-level in-cell rich-value images (Excel 365's
/// `=IMAGE(...)`/"Place in Cell"), keyed by the 1-based `vm` value a
/// cell carries. Excel stores these as a `t="e"` cell with a literal
/// `<v>#VALUE!</v>` fallback for old readers, plus `vm` pointing through
/// a chain of small parts to the real image: `xl/metadata.xml`
/// (`vm` -> `valueMetadata` bk -> `futureMetadata` bk -> `rvb.i`) ->
/// `xl/richData/rdrichvalue.xml` (`i` -> one `<rv>`'s struct + values) ->
/// `rdrichvaluestructure.xml` (which value is the image identifier) ->
/// `richValueRel.xml` (+ its own `.rels`) -> `xl/media/*`. Any missing or
/// malformed part along the way yields an empty map — this is a
/// best-effort recovery, not a hard requirement for opening the file.
fn read_rich_value_images<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    entries: &opc::ZipEntryIndex,
) -> std::collections::HashMap<u32, crate::xlsx::worksheet::WorksheetPicture> {
    let mut out = std::collections::HashMap::new();

    let Ok(metadata_xml) = XlsxDocument::read_xml_entry(archive, entries, "xl/metadata.xml") else {
        return out;
    };
    let Some((rich_type_id, future_rvb, value_metadata)) = parse_rich_value_metadata(&metadata_xml)
    else {
        return out;
    };

    let Ok(rv_xml) = XlsxDocument::read_xml_entry(archive, entries, "xl/richData/rdrichvalue.xml")
    else {
        return out;
    };
    let rv_list = parse_rdrichvalue(&rv_xml);

    let Ok(struct_xml) =
        XlsxDocument::read_xml_entry(archive, entries, "xl/richData/rdrichvaluestructure.xml")
    else {
        return out;
    };
    let image_key_positions = parse_rich_value_structures(&struct_xml);

    let rel_path = "xl/richData/richValueRel.xml";
    let Ok(rel_xml) = XlsxDocument::read_xml_entry(archive, entries, rel_path) else {
        return out;
    };
    let rel_rids = parse_rich_value_rel(&rel_xml);

    let rel_rels_path = sheet_rels_path(rel_path);
    let rel_rels = match XlsxDocument::read_xml_entry(archive, entries, &rel_rels_path) {
        Ok(d) => Relationships::parse(&d).unwrap_or_else(|_| Relationships::empty()),
        Err(_) => return out,
    };

    for (position, rvb_index) in value_metadata.iter().enumerate() {
        // `rc.t` must reference the metadataType we identified as
        // XLRICHVALUE — a `<bk>` referencing a different metadata type
        // (e.g. dynamic-array spill markers) isn't an image at all.
        let Some((rc_type, v)) = rvb_index else {
            continue;
        };
        if *rc_type != rich_type_id {
            continue;
        }
        let Some(Some(rv_index)) = future_rvb.get(*v as usize) else {
            continue;
        };
        let Some((struct_idx, values)) = rv_list.get(*rv_index as usize) else {
            continue;
        };
        let Some(Some(key_pos)) = image_key_positions.get(*struct_idx as usize) else {
            continue;
        };
        let Some(local_id_str) = values.get(*key_pos) else {
            continue;
        };
        let Ok(local_id) = local_id_str.parse::<usize>() else {
            continue;
        };
        let Some(rid) = rel_rids.get(local_id) else {
            continue;
        };
        let Some(rel) = rel_rels.get_by_id(rid) else {
            continue;
        };
        let media_path = resolve_relative_zip_path(rel_path, &rel.target);
        let Ok(bytes) = opc::read_zip_entry(archive, entries, &media_path) else {
            continue;
        };
        let ext = Path::new(&rel.target)
            .extension()
            .and_then(|s| s.to_str())
            .map(|s| s.to_ascii_lowercase())
            .unwrap_or_else(|| guess_image_format_from_bytes(&bytes).to_string());

        // No anchor coordinates exist for an in-cell rich value — it has
        // no `<xdr:from>`/`<xdr:to>` at all, unlike a drawing-anchored
        // picture. Zero EMUs is the same "position unknown" sentinel
        // `convert_xlsx.rs` already falls back on for a cell-anchored
        // drawing picture it couldn't place: the image still reaches the
        // IR, just inline rather than pixel-positioned.
        out.insert(
            (position as u32) + 1,
            crate::xlsx::worksheet::WorksheetPicture {
                data: bytes,
                format: ext,
                x_emu: 0,
                y_emu: 0,
                cx_emu: 0,
                cy_emu: 0,
                alt_text: None,
            },
        );
    }

    out
}

/// Parse `xl/metadata.xml`. Returns `(rich_type_id, future_rvb,
/// value_metadata)`:
/// - `rich_type_id`: the 1-based `<metadataType>` index whose `name` is
///   `"XLRICHVALUE"`.
/// - `future_rvb[v]`: the `<xlrd:rvb i="…">` value at position `v`
///   (0-based) of `<futureMetadata name="XLRICHVALUE">`'s `<bk>` array.
/// - `value_metadata[p]`: the `(t, v)` of the first `<rc>` in
///   `<valueMetadata>`'s `<bk>` at position `p` (0-based) — `p + 1` is
///   the cell's own `vm` value.
fn parse_rich_value_metadata(
    xml: &[u8],
) -> Option<(u32, Vec<Option<u32>>, Vec<Option<(u32, u32)>>)> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(xml);
    reader.config_mut().trim_text(true);

    #[derive(PartialEq)]
    enum Section {
        None,
        MetadataTypes,
        FutureMetadataRich,
        ValueMetadata,
    }
    let mut section = Section::None;
    let mut metadata_type_count = 0u32;
    let mut rich_type_id = None;
    let mut future_rvb: Vec<Option<u32>> = Vec::new();
    let mut value_metadata: Vec<Option<(u32, u32)>> = Vec::new();
    let mut in_bk = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    "metadataTypes" => section = Section::MetadataTypes,
                    "futureMetadata"
                        if xml::optional_attr_str(e, "name").ok().flatten().as_deref()
                            == Some("XLRICHVALUE") =>
                    {
                        section = Section::FutureMetadataRich;
                    },
                    "futureMetadata" => section = Section::None,
                    "valueMetadata" => section = Section::ValueMetadata,
                    "metadataType" if section == Section::MetadataTypes => {
                        metadata_type_count += 1;
                        if xml::optional_attr_str(e, "name").ok().flatten().as_deref()
                            == Some("XLRICHVALUE")
                        {
                            rich_type_id = Some(metadata_type_count);
                        }
                    },
                    "bk" if section == Section::FutureMetadataRich => {
                        future_rvb.push(None);
                        in_bk = true;
                    },
                    "bk" if section == Section::ValueMetadata => {
                        value_metadata.push(None);
                        in_bk = true;
                    },
                    "rvb" if section == Section::FutureMetadataRich && in_bk => {
                        if let Some(i) = xml::optional_attr_str(e, "i")
                            .ok()
                            .flatten()
                            .and_then(|v| v.parse::<u32>().ok())
                        {
                            if let Some(slot) = future_rvb.last_mut() {
                                *slot = Some(i);
                            }
                        }
                    },
                    "rc" if section == Section::ValueMetadata && in_bk => {
                        let t = xml::optional_attr_str(e, "t")
                            .ok()
                            .flatten()
                            .and_then(|v| v.parse::<u32>().ok());
                        let v = xml::optional_attr_str(e, "v")
                            .ok()
                            .flatten()
                            .and_then(|v| v.parse::<u32>().ok());
                        if let (Some(t), Some(v)) = (t, v) {
                            if let Some(slot @ None) = value_metadata.last_mut() {
                                *slot = Some((t, v));
                            }
                        }
                    },
                    _ => {},
                }
            },
            Ok(Event::End(ref e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    "bk" => in_bk = false,
                    "metadataTypes" | "futureMetadata" | "valueMetadata" => section = Section::None,
                    _ => {},
                }
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {},
        }
    }

    Some((rich_type_id?, future_rvb, value_metadata))
}

/// Parse `xl/richData/rdrichvalue.xml`: one `(struct_index, values)` per
/// `<rv>`, in document order (index `i` in `xlrd:rvb` refers to this
/// position).
fn parse_rdrichvalue(xml: &[u8]) -> Vec<(u32, Vec<String>)> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(xml);
    reader.config_mut().trim_text(true);

    let mut out: Vec<(u32, Vec<String>)> = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == "rv" => {
                let s = xml::optional_attr_str(e, "s")
                    .ok()
                    .flatten()
                    .and_then(|v| v.parse::<u32>().ok())
                    .unwrap_or(0);
                out.push((s, Vec::new()));
            },
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == "v" => {
                if let Ok(text) = xml::read_text_content_fast(&mut reader) {
                    if let Some((_, values)) = out.last_mut() {
                        values.push(text);
                    }
                }
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {},
        }
    }
    out
}

/// Parse `xl/richData/rdrichvaluestructure.xml`: for each `<s>` (struct
/// type, in document order), the 0-based position of its
/// `<k n="_rvRel:LocalImageIdentifier">` key among that struct's `<k>`
/// children, or `None` if the struct has no such key (it isn't an
/// image-carrying rich-value type).
fn parse_rich_value_structures(xml: &[u8]) -> Vec<Option<usize>> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(xml);
    reader.config_mut().trim_text(true);

    let mut out = Vec::new();
    let mut current: Option<(usize, Option<usize>)> = None; // (key count so far, image key position)
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == "s" => {
                current = Some((0, None));
            },
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.local_name().as_ref() == "k" => {
                if let Some((count, image_pos)) = current.as_mut() {
                    if xml::optional_attr_str(e, "n").ok().flatten().as_deref()
                        == Some("_rvRel:LocalImageIdentifier")
                    {
                        *image_pos = Some(*count);
                    }
                    *count += 1;
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == "s" => {
                if let Some((_, image_pos)) = current.take() {
                    out.push(image_pos);
                }
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {},
        }
    }
    out
}

/// Parse `xl/richData/richValueRel.xml`: the `r:id` of each `<rel>`, in
/// document order (a rich value's `_rvRel:LocalImageIdentifier` is a
/// 0-based index into this array).
fn parse_rich_value_rel(xml: &[u8]) -> Vec<String> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(xml);
    reader.config_mut().trim_text(true);

    let mut out = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.local_name().as_ref() == "rel" =>
            {
                if let Some(rid) = xml::optional_attr_str(e, "r:id").ok().flatten() {
                    out.push(rid.into_owned());
                }
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {},
        }
    }
    out
}

/// Resolve a `..`-relative target inside an OPC package back to an
/// absolute ZIP-entry path. Mirrors `PartName::resolve_relative` but
/// operates on plain ZIP paths (the `from_zip` fast path doesn't use
/// `PartName`).
pub(super) fn resolve_relative_zip_path(source: &str, target: &str) -> String {
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
                let local = e.local_name().as_ref().to_string();
                match local.as_str() {
                    "absoluteAnchor" | "oneCellAnchor" | "twoCellAnchor" => {
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
                    "pic" if in_anchor => {
                        kind = AnchorKind::Picture;
                    },
                    "sp" if in_anchor => {
                        kind = AnchorKind::Text;
                    },
                    "txBody" if in_anchor => {
                        in_txbody = true;
                    },
                    "r" if in_txbody => {
                        in_run = true;
                    },
                    "t" if in_run => {
                        in_a_t = true;
                    },
                    "rPr" if in_run => {
                        for attr in e.attributes().with_checks(false) {
                            let attr = attr.map_err(crate::core::Error::from)?;
                            let key = attr.key.as_ref();
                            let raw = crate::core::xml::unescape_attr_value(&attr)?;
                            match key {
                                "sz" => {
                                    // sz is in hundredths of a pt.
                                    if let Ok(n) = raw.parse::<i32>() {
                                        font_size_pt = Some(n as f32 / 100.0);
                                    }
                                },
                                "b" => bold = raw == "1" || raw == "true",
                                "i" => italic = raw == "1" || raw == "true",
                                _ => {},
                            }
                        }
                    },
                    "solidFill" if in_run => {
                        in_solid_fill = true;
                    },
                    "cNvPr" if in_anchor => {
                        if let Some(d) = crate::core::xml::optional_attr_str(e, "descr")? {
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
                let local = e.local_name().as_ref().to_string();
                match local.as_str() {
                    "pos" => {
                        if let Some(v) = crate::core::xml::optional_attr_str(e, "x")? {
                            x_emu = v.parse().unwrap_or(0);
                        }
                        if let Some(v) = crate::core::xml::optional_attr_str(e, "y")? {
                            y_emu = v.parse().unwrap_or(0);
                        }
                    },
                    "ext" => {
                        if let Some(v) = crate::core::xml::optional_attr_str(e, "cx")? {
                            cx_emu = v.parse().unwrap_or(0);
                        }
                        if let Some(v) = crate::core::xml::optional_attr_str(e, "cy")? {
                            cy_emu = v.parse().unwrap_or(0);
                        }
                    },
                    "off" if cx_emu == 0 && cy_emu == 0 && matches!(kind, AnchorKind::Unknown) => {
                        // Honour `<off>` only at the outermost anchor level,
                        // before we've descended into `<xdr:pic>` or
                        // `<xdr:sp>`. Otherwise the `<a:off>` inside a
                        // shape's `<a:xfrm>` (which expresses a transform
                        // local to the shape, not the anchor origin) would
                        // overwrite the absolute coordinates parsed from
                        // `<xdr:pos>`.
                        if let Some(v) = crate::core::xml::optional_attr_str(e, "x")? {
                            x_emu = v.parse().unwrap_or(x_emu);
                        }
                        if let Some(v) = crate::core::xml::optional_attr_str(e, "y")? {
                            y_emu = v.parse().unwrap_or(y_emu);
                        }
                    },
                    "blip" => {
                        for attr in e.attributes().with_checks(false) {
                            let attr = attr.map_err(crate::core::Error::from)?;
                            let key = attr.key.as_ref();
                            if key == "r:embed" || key.ends_with(":embed") || key == "embed" {
                                embed_rid = Some(crate::core::xml::unescape_attr_value(&attr)?);
                                break;
                            }
                        }
                    },
                    "cNvPr" => {
                        if let Some(d) = crate::core::xml::optional_attr_str(e, "descr")? {
                            alt_text = Some(d.into_owned());
                        }
                    },
                    "latin" if in_run => {
                        if let Some(t) = crate::core::xml::optional_attr_str(e, "typeface")? {
                            font_name = Some(t.into_owned());
                        }
                    },
                    "srgbClr" if in_solid_fill => {
                        if let Some(v) = crate::core::xml::optional_attr_str(e, "val")? {
                            color_hex = Some(v.into_owned().to_uppercase());
                        }
                    },
                    "rPr" if in_run => {
                        for attr in e.attributes().with_checks(false) {
                            let attr = attr.map_err(crate::core::Error::from)?;
                            let key = attr.key.as_ref();
                            let raw = crate::core::xml::unescape_attr_value(&attr)?;
                            match key {
                                "sz" => {
                                    if let Ok(n) = raw.parse::<i32>() {
                                        font_size_pt = Some(n as f32 / 100.0);
                                    }
                                },
                                "b" => bold = raw == "1" || raw == "true",
                                "i" => italic = raw == "1" || raw == "true",
                                _ => {},
                            }
                        }
                    },
                    _ => {},
                }
            },
            Event::Text(ref e) if in_a_t => {
                let s = crate::core::xml::unescape_text(e)?;
                text_buf.push_str(&s);
            },
            Event::GeneralRef(ref e) if in_a_t => {
                text_buf.push_str(&crate::core::xml::resolve_general_ref(e)?);
            },
            Event::End(ref e) => {
                let local = e.local_name().as_ref().to_string();
                match local.as_str() {
                    "t" => in_a_t = false,
                    "r" => in_run = false,
                    "txBody" => in_txbody = false,
                    "solidFill" => in_solid_fill = false,
                    s if matches!(s, "absoluteAnchor" | "oneCellAnchor" | "twoCellAnchor")
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

/// Hand-built XLSX packages for in-crate regression tests.
///
/// Several bugs only show up end-to-end (the part is parsed correctly and
/// dropped later, or the package-level encoding is what's wrong), so the
/// tests need a real zip rather than a bare XML buffer — but not a corpus
/// file.
#[cfg(test)]
pub(crate) mod test_support {
    use std::io::{Cursor, Write};

    /// Build a zip from `(entry_name, bytes)` pairs, in order.
    pub(crate) fn zip_parts(parts: &[(&str, &[u8])]) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts: zip::write::FileOptions<'_, ()> = zip::write::FileOptions::default();
        for (name, data) in parts {
            zip.start_file(*name, opts).unwrap();
            zip.write_all(data).unwrap();
        }
        zip.finish().unwrap().into_inner()
    }

    /// Wrap one worksheet (and optional extra parts) into a single-sheet
    /// XLSX package. `sheet_xml` is the full `xl/worksheets/sheet1.xml` body.
    pub(crate) fn single_sheet_xlsx(sheet_xml: &str, extra: &[(&str, &[u8])]) -> Vec<u8> {
        let rels = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
</Relationships>"#;
        let workbook = br#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#;
        let mut parts: Vec<(&str, &[u8])> = vec![
            ("xl/_rels/workbook.xml.rels", rels),
            ("xl/workbook.xml", workbook),
            ("xl/worksheets/sheet1.xml", sheet_xml.as_bytes()),
        ];
        parts.extend_from_slice(extra);
        zip_parts(&parts)
    }

    /// Open a hand-built package through the normal XLSX reader.
    pub(crate) fn open_bytes(bytes: Vec<u8>) -> super::XlsxDocument {
        super::XlsxDocument::from_reader(Cursor::new(bytes)).expect("package opens")
    }
}

#[cfg(test)]
mod tests {
    use super::test_support::*;
    use super::*;

    /// Every sheet probes for its optional `_rels`, drawing and comments
    /// parts; when the probe missed the exact name it scanned the whole
    /// archive, so a workbook of N trivial sheets cost O(N²): 2,000 sheets
    /// took 6 s in release and 8,000 took 56 s; 4,000 took ~20 s in a debug
    /// build against this budget, which the indexed lookup meets 4× over.
    #[test]
    fn test_many_sheet_workbook_opens_in_linear_time() {
        const N: usize = 4000;
        let mut rels = String::from(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        let mut sheets = String::new();
        let mut parts: Vec<(String, Vec<u8>)> = Vec::new();
        for i in 1..=N {
            rels.push_str(&format!(
                r#"<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>"#
            ));
            sheets.push_str(&format!(r#"<sheet name="S{i}" sheetId="{i}" r:id="rId{i}"/>"#));
            parts.push((
                format!("xl/worksheets/sheet{i}.xml"),
                format!(
                    r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>{i}</v></c></row></sheetData></worksheet>"#
                )
                .into_bytes(),
            ));
        }
        rels.push_str("</Relationships>");
        let workbook = format!(
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{sheets}</sheets></workbook>"#
        );
        let mut all: Vec<(&str, &[u8])> = vec![
            ("xl/_rels/workbook.xml.rels", rels.as_bytes()),
            ("xl/workbook.xml", workbook.as_bytes()),
        ];
        all.extend(parts.iter().map(|(n, d)| (n.as_str(), d.as_slice())));
        let bytes = zip_parts(&all);

        let started = std::time::Instant::now();
        let doc = open_bytes(bytes);
        assert_eq!(doc.workbook.sheets.len(), N);
        assert!(
            started.elapsed() < std::time::Duration::from_secs(15),
            "{N} one-cell sheets took {:?} to open",
            started.elapsed()
        );
    }

    /// TableCell::col_span/row_span were hardcoded to 1
    /// on every spreadsheet cell; merged_cells was parsed and then
    /// never read on the to_ir() path.
    #[test]
    fn test_merged_cell_range_sets_col_span_on_the_anchor_and_excludes_covered_cells() {
        // TableCell::col_span/row_span were hardcoded to 1
        // on every spreadsheet cell; merged_cells was parsed and then
        // never read on the to_ir() path, so a merged header/label
        // flattened to an ordinary unspanned grid.
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Header</t></is></c>
      <c r="B1" t="inlineStr"><is><t></t></is></c>
      <c r="C1" t="inlineStr"><is><t></t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>a</t></is></c>
      <c r="B2" t="inlineStr"><is><t>b</t></is></c>
      <c r="C2" t="inlineStr"><is><t>c</t></is></c>
    </row>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="A1:C1"/>
  </mergeCells>
</worksheet>"#;
        let doc = open_bytes(single_sheet_xlsx(sheet_xml, &[]));
        let ir = crate::convert_xlsx::xlsx_to_ir(&doc);
        let table = ir.sections[0]
            .elements
            .iter()
            .find_map(|e| match e {
                crate::ir::Element::Table(t) => Some(t),
                _ => None,
            })
            .expect("expected a table element");

        let header_row = &table.rows[0];
        assert_eq!(
            header_row.cells.len(),
            1,
            "the 2 covered cells must be excluded, leaving only the anchor: {:?}",
            header_row.cells
        );
        assert_eq!(header_row.cells[0].col_span, 3, "anchor must carry the real span");
        assert_eq!(header_row.cells[0].row_span, 1);

        let data_row = &table.rows[1];
        assert_eq!(data_row.cells.len(), 3, "an unmerged row must keep all 3 cells");
    }

    /// A formula cell with no cached `<v>` (the default output
    /// shape of closedxml and similar writers) rendered as a blank cell
    /// indistinguishable from a genuinely empty one, and the formula text
    /// never reached any consumer at all.
    #[test]
    fn test_uncached_formula_cell_shows_formula_text_and_reaches_the_ir() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="array" ref="A1:B2">1+2</f></c>
    </row>
  </sheetData>
</worksheet>"#;
        let doc = open_bytes(single_sheet_xlsx(sheet_xml, &[]));
        let ir = crate::convert_xlsx::xlsx_to_ir(&doc);
        let table = ir.sections[0]
            .elements
            .iter()
            .find_map(|e| match e {
                crate::ir::Element::Table(t) => Some(t),
                _ => None,
            })
            .expect("expected a table element");
        let cell = &table.rows[0].cells[0];
        assert_eq!(cell.formula.as_deref(), Some("1+2"), "formula text must reach the IR");
        assert!(
            !cell.content.is_empty(),
            "a formula cell with no cached value must not render as a blank cell"
        );
        let text = ir.plain_text();
        assert!(
            text.contains("=1+2"),
            "plain_text() must show the formula as a fallback when there's no cached value, got: {text:?}"
        );
    }

    /// `Document::plain_text()`/`to_markdown()` dispatch to `XlsxDocument`'s
    /// own renderer (`xlsx/text.rs`), a separate path from `to_ir()` —
    /// fixing only the IR side left the CLI's default `text`/`markdown`
    /// output still blank for an uncached formula cell. Both paths must
    /// show the fallback.
    /// Regression: a cell comment reached the IR (as an endnote) and so
    /// the HTML surface, but the direct `plain_text()`/`to_markdown()`
    /// renderers dropped it — the same text present on one surface and
    /// absent on another.
    #[test]
    fn test_comments_reach_plain_text_and_markdown() {
        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>data</t></is></c></row></sheetData>
</worksheet>"#;
        let sheet_rels = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    Target="../comments1.xml"/>
</Relationships>"#;
        let comments = br#"<?xml version="1.0" encoding="UTF-8"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>Reviewer</author></authors>
  <commentList><comment ref="B2" authorId="0"><text><t>a real cell comment</t></text></comment></commentList>
</comments>"#;
        let doc = open_bytes(single_sheet_xlsx(
            sheet,
            &[
                ("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels),
                ("xl/comments1.xml", comments),
            ],
        ));
        assert_eq!(doc.worksheets[0].comments.len(), 1, "fixture must parse its comment");
        let text = doc.plain_text();
        assert!(
            text.contains("B2 (Reviewer): a real cell comment"),
            "plain_text must carry the comment with its cell and author: {text}"
        );
        let md = doc.to_markdown();
        assert!(
            md.contains("> **B2 (Reviewer):** a real cell comment"),
            "to_markdown must carry the comment: {md}"
        );
    }

    #[test]
    fn test_uncached_formula_cell_shows_formula_text_via_the_low_level_xlsx_renderer() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="array" ref="A1:B2">1+2</f></c>
    </row>
  </sheetData>
</worksheet>"#;
        let doc = open_bytes(single_sheet_xlsx(sheet_xml, &[]));
        assert_eq!(
            doc.plain_text().trim(),
            "Sheet1\n=1+2",
            "XlsxDocument::plain_text() must fall back to the formula, not blank"
        );
        assert!(
            doc.to_markdown().contains("=1+2"),
            "XlsxDocument::to_markdown() must fall back to the formula, got: {:?}",
            doc.to_markdown()
        );
    }

    /// A formula cell that *does* have a cached value keeps showing that
    /// value (unaffected default behaviour) while also exposing the
    /// formula text on `TableCell::formula`.
    #[test]
    fn test_cached_formula_cell_keeps_its_value_and_also_exposes_the_formula() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>1+2</f><v>3</v></c>
    </row>
  </sheetData>
</worksheet>"#;
        let doc = open_bytes(single_sheet_xlsx(sheet_xml, &[]));
        let ir = crate::convert_xlsx::xlsx_to_ir(&doc);
        let table = ir.sections[0]
            .elements
            .iter()
            .find_map(|e| match e {
                crate::ir::Element::Table(t) => Some(t),
                _ => None,
            })
            .expect("expected a table element");
        let cell = &table.rows[0].cells[0];
        assert_eq!(cell.formula.as_deref(), Some("1+2"));
        let text = ir.plain_text();
        assert!(text.contains('3'), "cached value must still be shown, got: {text:?}");
        assert!(
            !text.contains("=1+2"),
            "the formula fallback text must not override a real cached value, got: {text:?}"
        );
    }

    /// Same gap as DOCX, confirmed independently for XLSX.
    #[test]
    fn test_encrypted_xlsx_gives_a_friendly_error_via_the_format_specific_reader() {
        let mut cfb = vec![0u8; 512];
        cfb[0..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
        let err = XlsxDocument::from_reader(std::io::Cursor::new(cfb)).unwrap_err();
        let msg = err.to_string();
        assert!(
            msg.contains("password-protected"),
            "expected a friendly password-protected message, got: {msg}"
        );
    }

    /// The macro-presence signal is the workbook part's own `vbaProject`
    /// relationship; it was asserted for DOCX only.
    #[test]
    fn test_vba_project_relationship_sets_has_macros() {
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>"#;
        let rels = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>"#;
        let workbook = br#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#;
        let with = open_bytes(zip_parts(&[
            ("xl/_rels/workbook.xml.rels", rels),
            ("xl/workbook.xml", workbook),
            ("xl/worksheets/sheet1.xml", sheet_xml.as_bytes()),
            ("xl/vbaProject.bin", b"fake vba bytes"),
        ]));
        assert!(with.has_macros);
        assert!(crate::convert_xlsx::xlsx_to_ir(&with).metadata.has_macros);
        let without = open_bytes(single_sheet_xlsx(sheet_xml, &[]));
        assert!(!without.has_macros);
    }

    /// AppProperties::parse existed, fully tested, but
    /// nothing on the read side ever called it (fast zip path).
    #[test]
    fn test_app_properties_are_read_on_open() {
        let app_xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Company>Acme Corp</Company>
  <Words>1250</Words>
</Properties>"#;
        let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>"#;
        let doc = open_bytes(single_sheet_xlsx(sheet_xml, &[("docProps/app.xml", app_xml)]));
        let app = doc
            .app_properties
            .expect("app_properties must be populated");
        assert_eq!(app.company.as_deref(), Some("Acme Corp"));
        assert_eq!(app.words, Some(1250));
    }

    /// A UTF-16BE `xl/workbook.xml` used to parse as a stream of
    /// unrecognised tags and yield zero sheets, silently, with `Ok`.
    #[test]
    fn test_utf16be_workbook_xml_decodes_instead_of_yielding_zero_sheets() {
        let workbook = r#"<?xml version="1.0" encoding="UTF-16BE"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Utf16" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#;
        // UTF-16BE, no BOM — spec-legal per XML 1.0 §4.3.3 when the text
        // declaration names the encoding.
        let utf16: Vec<u8> = workbook
            .encode_utf16()
            .flat_map(|u| u.to_be_bytes())
            .collect();

        let rels = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
</Relationships>"#;
        let sheet = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>"#;

        let doc = open_bytes(zip_parts(&[
            ("xl/_rels/workbook.xml.rels", rels),
            ("xl/workbook.xml", &utf16),
            ("xl/worksheets/sheet1.xml", sheet),
        ]));

        assert_eq!(doc.workbook.sheets.len(), 1, "the one declared sheet must survive");
        assert_eq!(doc.workbook.sheets[0].name, "Utf16");
        assert_eq!(doc.worksheets.len(), 1);
        assert_eq!(doc.plain_text(), "Utf16\n1");
    }

    /// XLSX half of the date-overflow hang — a huge number under a style whose `<numFmt>`
    /// override makes a built-in *date* id (50) mean something else must
    /// render as a number, promptly. Before the fix the cell was classified
    /// as a date and `from_serial` spun for ~2.5e16 iterations.
    #[test]
    fn test_huge_number_under_overridden_date_format_id_does_not_hang() {
        let styles = br#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="50" formatCode="0.00000E+0"/></numFmts>
  <cellXfs count="1"><xf numFmtId="50" applyNumberFormat="1"/></cellXfs>
</styleSheet>"#;
        let sheet = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" s="0"><v>1e300</v></c></row></sheetData>
</worksheet>"#;

        let doc = open_bytes(single_sheet_xlsx(sheet, &[("xl/styles.xml", styles)]));

        let started = std::time::Instant::now();
        let text = doc.plain_text();
        assert!(
            started.elapsed() < std::time::Duration::from_secs(2),
            "rendering one cell took {:?}",
            started.elapsed()
        );
        assert!(
            !text.starts_with("19") && !text.starts_with("20"),
            "1e300 must not render as a calendar date, got {text:?}"
        );
    }

    #[test]
    fn test_sheet_rels_path_top_level() {
        assert_eq!(
            sheet_rels_path("xl/worksheets/sheet1.xml"),
            "xl/worksheets/_rels/sheet1.xml.rels"
        );
        assert_eq!(sheet_rels_path("sheet1.xml"), "_rels/sheet1.xml.rels");
    }

    #[test]
    fn test_resolve_relative_zip_path_absolute() {
        assert_eq!(
            resolve_relative_zip_path("xl/worksheets/sheet1.xml", "/xl/media/img1.png"),
            "xl/media/img1.png"
        );
    }

    #[test]
    fn test_resolve_relative_zip_path_dotdot() {
        assert_eq!(
            resolve_relative_zip_path("xl/worksheets/sheet1.xml", "../drawings/drawing1.xml"),
            "xl/drawings/drawing1.xml"
        );
    }

    #[test]
    fn test_resolve_relative_zip_path_dot_segment() {
        assert_eq!(
            resolve_relative_zip_path("xl/worksheets/sheet1.xml", "./local.xml"),
            "xl/worksheets/local.xml"
        );
    }

    #[test]
    fn test_resolve_relative_zip_path_source_at_root() {
        assert_eq!(resolve_relative_zip_path("file.xml", "sub/x.xml"), "sub/x.xml");
    }

    #[test]
    fn test_guess_image_format_signatures() {
        assert_eq!(guess_image_format_from_bytes(&[0x89, b'P', b'N', b'G', 13, 10, 26, 10]), "png");
        assert_eq!(guess_image_format_from_bytes(&[0xFF, 0xD8, 0xFF, 0xE0]), "jpeg");
        assert_eq!(guess_image_format_from_bytes(b"GIF89a..."), "gif");
        assert_eq!(guess_image_format_from_bytes(b"GIF87a..."), "gif");
        assert_eq!(guess_image_format_from_bytes(b"BM\0\0\0"), "bmp");
        assert_eq!(guess_image_format_from_bytes(b"II*\0\x08\0"), "tiff");
        assert_eq!(guess_image_format_from_bytes(b"MM\0*\0\x08"), "tiff");
        assert_eq!(guess_image_format_from_bytes(&[0xD7, 0xCD, 0xC6, 0x9A]), "wmf");
        assert_eq!(guess_image_format_from_bytes(&[0x01, 0x00, 0x00, 0x00, 0x58]), "emf");
        // Fall back to png for unknown payloads.
        assert_eq!(guess_image_format_from_bytes(&[0, 0, 0]), "png");
    }

    #[test]
    fn test_extract_chart_text_minimal_title() {
        let xml = br#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:p><a:r><a:t>Quarterly Sales</a:t></a:r></a:p>
        </c:rich>
      </c:tx>
    </c:title>
  </c:chart>
</c:chartSpace>"#;
        let out = extract_chart_text(xml);
        assert!(out.contains("Title: Quarterly Sales"), "got: {out}");
    }

    #[test]
    fn test_extract_chart_text_series_and_categories() {
        let xml = br#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea>
    <c:barChart>
      <c:ser>
        <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:pt><c:v>Budget</c:v></c:pt></c:strCache></c:strRef></c:tx>
        <c:cat><c:strRef><c:strCache>
          <c:pt><c:v>Q1</c:v></c:pt>
          <c:pt><c:v>Q2</c:v></c:pt>
        </c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:numCache>
          <c:pt><c:v>1000</c:v></c:pt>
          <c:pt><c:v>2000</c:v></c:pt>
        </c:numCache></c:numRef></c:val>
      </c:ser>
    </c:barChart>
  </c:plotArea></c:chart>
</c:chartSpace>"#;
        let out = extract_chart_text(xml);
        assert!(out.contains("Categories: Q1, Q2"), "got: {out}");
        assert!(out.contains("Budget: 1000, 2000"), "got: {out}");
    }

    /// Excel splits a formatted title across runs, sometimes
    /// mid-word; the inter-run space rides on one run's edge, so trimming
    /// each run before concatenating deleted it.
    #[test]
    fn test_chart_title_keeps_spaces_between_runs() {
        let xml = br#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p>
      <a:r><a:t>Chart </a:t></a:r>
      <a:r><a:t>Title</a:t></a:r>
      <a:r><a:t> - </a:t></a:r>
      <a:r><a:t>with </a:t></a:r>
      <a:r><a:t>a</a:t></a:r>
      <a:r><a:t>dd</a:t></a:r>
      <a:r><a:t>iti</a:t></a:r>
      <a:r><a:t>o</a:t></a:r>
      <a:r><a:t>nal </a:t></a:r>
      <a:r><a:t>format</a:t></a:r>
      <a:r><a:t>ting</a:t></a:r>
    </a:p></c:rich></c:tx></c:title>
  </c:chart>
</c:chartSpace>"#;
        let out = extract_chart_text(xml);
        assert_eq!(out, "Title: Chart Title - with additional formatting", "got: {out}");
    }

    /// scatter/bubble series hold their points in
    /// `<c:xVal>`/`<c:yVal>`/`<c:bubbleSize>`, not `<c:cat>`/`<c:val>`.
    #[test]
    fn test_scatter_and_bubble_series_data_points_are_captured() {
        let xml = br#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea>
    <c:scatterChart>
      <c:ser>
        <c:tx><c:strRef><c:strCache><c:pt><c:v>Y</c:v></c:pt></c:strCache></c:strRef></c:tx>
        <c:xVal><c:numRef><c:numCache>
          <c:pt><c:v>0</c:v></c:pt><c:pt><c:v>1</c:v></c:pt>
        </c:numCache></c:numRef></c:xVal>
        <c:yVal><c:numRef><c:numCache>
          <c:pt><c:v>0.5</c:v></c:pt><c:pt><c:v>1.5</c:v></c:pt>
        </c:numCache></c:numRef></c:yVal>
      </c:ser>
    </c:scatterChart>
  </c:plotArea></c:chart>
</c:chartSpace>"#;
        let out = extract_chart_text(xml);
        assert!(out.contains("Categories: 0, 1"), "x-values missing: {out}");
        assert!(out.contains("Y: 0.5, 1.5"), "y-values missing: {out}");

        let bubble = br#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea>
    <c:bubbleChart>
      <c:ser>
        <c:yVal><c:numRef><c:numCache><c:pt><c:v>7</c:v></c:pt></c:numCache></c:numRef></c:yVal>
        <c:bubbleSize><c:numRef><c:numCache>
          <c:pt><c:v>3</c:v></c:pt>
        </c:numCache></c:numRef></c:bubbleSize>
      </c:ser>
    </c:bubbleChart>
  </c:plotArea></c:chart>
</c:chartSpace>"#;
        let out = extract_chart_text(bubble);
        assert!(out.contains("7"), "bubble y-value missing: {out}");
        assert!(out.contains("3"), "bubble size missing: {out}");
    }

    #[test]
    fn test_parse_drawing_anchors_picture_one_cell() {
        let xml = br#"<?xml version="1.0"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>914400</xdr:colOff>
              <xdr:row>0</xdr:row><xdr:rowOff>457200</xdr:rowOff></xdr:from>
    <xdr:ext cx="2000000" cy="1500000"/>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="2" name="Image1" descr="my-alt"/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId4"/>
      </xdr:blipFill>
    </xdr:pic>
  </xdr:oneCellAnchor>
</xdr:wsDr>"#;
        let parsed = parse_drawing_anchors(xml).expect("parse ok");
        assert_eq!(parsed.pictures.len(), 1);
        assert_eq!(parsed.pictures[0].embed_rid, "rId4");
        assert_eq!(parsed.pictures[0].cx_emu, 2_000_000);
        assert_eq!(parsed.pictures[0].cy_emu, 1_500_000);
        assert_eq!(parsed.pictures[0].alt_text.as_deref(), Some("my-alt"));
    }

    #[test]
    fn test_parse_drawing_anchors_text_shape() {
        let xml = br#"<?xml version="1.0"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:absoluteAnchor>
    <xdr:pos x="100000" y="200000"/>
    <xdr:ext cx="3000000" cy="500000"/>
    <xdr:sp>
      <xdr:txBody>
        <a:p><a:r><a:t>Hello shape</a:t></a:r></a:p>
      </xdr:txBody>
    </xdr:sp>
  </xdr:absoluteAnchor>
</xdr:wsDr>"#;
        let parsed = parse_drawing_anchors(xml).expect("parse ok");
        assert_eq!(parsed.text_shapes.len(), 1);
        assert_eq!(parsed.text_shapes[0].text, "Hello shape");
        assert_eq!(parsed.text_shapes[0].cx_emu, 3_000_000);
    }

    #[test]
    fn test_parse_drawing_anchors_empty_doc_is_ok() {
        let xml = br#"<?xml version="1.0"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>"#;
        let parsed = parse_drawing_anchors(xml).expect("parse ok");
        assert!(parsed.pictures.is_empty());
        assert!(parsed.text_shapes.is_empty());
    }

    /// Build a minimal SpreadsheetML package whose main part carries
    /// `content_type`.
    /// A `vm`-tagged `t="e"` cell whose fallback `<v>` is
    /// the literal `"#VALUE!"` is Excel 365's in-cell rich-value image
    /// (`=IMAGE(...)`/"Place in Cell"), not a real formula error. The
    /// real image is reachable by resolving `vm` through `xl/
    /// metadata.xml` -> `xl/richData/{rdrichvalue,
    /// rdrichvaluestructure,richValueRel}.xml` -> `xl/media/*`. This
    /// fixture mirrors the exact shape of the real-corpus reproducer
    /// (`phpspreadsheet_drawing_in_cell.xlsx`) byte for byte.
    #[test]
    fn test_a_rich_value_image_cell_resolves_to_a_real_image_not_a_value_error() {
        const PNG: &[u8] = &[
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00,
            0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08,
            0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00, 0x18, 0xDD, 0x8D,
            0xB0, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
        ];

        let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2"><c r="B2" t="e" vm="1"><v>#VALUE!</v></c></row>
  </sheetData>
</worksheet>"#;

        let metadata_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"><metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/></metadataTypes><futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>"#;

        let rdrichvalue_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1"><rv s="0"><v>0</v><v>5</v></rv></rvData>"#;

        let rdrichvaluestructure_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1"><s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/></s></rvStructures>"#;

        let richvaluerel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><rel r:id="rId1"/></richValueRels>"#;

        let richvaluerel_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#;

        let bytes = single_sheet_xlsx(
            std::str::from_utf8(sheet_xml).unwrap(),
            &[
                ("xl/metadata.xml", metadata_xml.as_slice()),
                ("xl/richData/rdrichvalue.xml", rdrichvalue_xml.as_slice()),
                ("xl/richData/rdrichvaluestructure.xml", rdrichvaluestructure_xml.as_slice()),
                ("xl/richData/richValueRel.xml", richvaluerel_xml.as_slice()),
                ("xl/richData/_rels/richValueRel.xml.rels", richvaluerel_rels.as_slice()),
                ("xl/media/image1.png", PNG),
            ],
        );
        let doc = open_bytes(bytes);
        let ws = &doc.worksheets[0];

        assert_eq!(ws.images.len(), 1, "the rich-value image must reach ws.images");
        assert_eq!(ws.images[0].data, PNG);
        assert_eq!(ws.images[0].format, "png");

        let cell = &ws.rows[0].cells[0];
        assert!(
            !matches!(cell.value, CellValue::Error(_)),
            "the fabricated #VALUE! error must be cleared: {:?}",
            cell.value
        );
    }
}
