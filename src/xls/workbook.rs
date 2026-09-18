//! Workbook-level parsing: sheets, SST, cell grid construction.

use std::io::{Read, Seek};

use crate::cfb::CfbReader;

use super::cell::{Cell, CellValue, parse_cell_record};
use super::error::{Result, XlsError};
use super::images::{XlsImage, extract_images};
use super::records::*;
use super::sst::{parse_sst, read_short_unicode_string, read_unicode_string};

/// A parsed legacy XLS document.
#[derive(Debug)]
pub struct XlsDocument {
    /// Worksheets in workbook order.
    pub sheets: Vec<Sheet>,
    /// Named ranges from `NAME` (`0x0018`) records ([MS-XLS] §2.4.174),
    /// resolved via `EXTERNSHEET`/`SUPBOOK`. A name whose formula isn't a
    /// single (possibly 3-D) cell or area reference resolves with an empty
    /// `value` rather than a guessed one (issue #251, XLS half).
    pub defined_names: Vec<DefinedName>,
    images: Vec<XlsImage>,
    has_macros: bool,
    /// `true` when the record-parsing safety cap ran out before the
    /// Workbook stream did — trailing sheets, or the whole workbook, are
    /// missing from `sheets` with no other signal a caller could use to
    /// tell that apart from a file that genuinely ended there (issue
    /// #236).
    truncated: bool,
}

/// A named range recovered from a `NAME` record.
#[derive(Debug, Clone, PartialEq)]
pub struct DefinedName {
    pub name: String,
    /// The referenced range as `Sheet1!$A$1:$B$10`, or empty when the
    /// formula wasn't a plain cell/area reference this parser resolves.
    pub value: String,
    /// 0-based sheet index when the name is local to one sheet.
    pub local_sheet_id: Option<u32>,
    pub hidden: bool,
}

#[cfg(test)]
impl XlsDocument {
    /// Build a document from sheets alone, so conversion can be exercised on
    /// a grid shape without a BIFF fixture to encode it in.
    pub(crate) fn from_sheets(sheets: Vec<Sheet>) -> Self {
        Self {
            sheets,
            defined_names: Vec::new(),
            images: Vec::new(),
            has_macros: false,
            truncated: false,
        }
    }
}

/// A worksheet from an XLS workbook.
#[derive(Debug, Default)]
pub struct Sheet {
    /// Sheet display name.
    pub name: String,
    /// Rendered display text per cell, mirroring `rows`.
    ///
    /// XLS stores dates as plain numbers whose *format* makes them a date;
    /// `FORMAT` and `XF` were never parsed, so `38971` came out as `38971`
    /// rather than `2006-09-05`. `rows` still carries the raw values.
    pub display: Vec<Vec<String>>,
    /// Cell values, indexed as `rows[row][col]`.
    pub rows: Vec<Vec<CellValue>>,
    /// Whether `BOUNDSHEET` marked this sheet hidden or very hidden.
    ///
    /// "Hidden" means "not shown in the UI by default", not "deleted": the
    /// sheet is kept and flagged rather than dropped, matching XLSX's
    /// `Section::hidden`.
    pub hidden: bool,
    /// Merged cell ranges from `MERGEDCELLS` (`0x00E5`), as 0-based
    /// `(row_first, row_last, col_first, col_last)` tuples. The record
    /// wasn't parsed at all before — merge information for a legacy
    /// `.xls` was discarded before it was even in memory, not just
    /// dropped at IR conversion (issue #235, XLS half).
    pub merged_cells: Vec<(u16, u16, u16, u16)>,
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
    pub fn from_reader<R: Read + Seek>(mut reader: R) -> Result<Self> {
        // A pre-OLE2 Excel file (BIFF2/3/4 — Excel 2.x through 4.x) has no
        // compound-file container at all: it is a raw record stream that
        // starts with its own `BOF`. `CfbReader` rejects it with "bad magic
        // signature", which is indistinguishable from a corrupt or
        // unrelated file. Name the format instead, mirroring the Word
        // 6.0/95 case in `doc::fib`.
        if let Some(biff) = detect_raw_biff(&mut reader)? {
            return Err(XlsError::UnsupportedVersion(format!(
                "BIFF{biff} (pre-OLE2 Excel {}); only BIFF8 (Excel 97 and later) is supported",
                match biff {
                    2 => "2.x",
                    3 => "3.0",
                    _ => "4.0",
                }
            )));
        }

        let mut cfb = CfbReader::new(reader)?;

        // Try "Workbook" (BIFF8) first, then "Book" (BIFF5).
        let stream_data = if cfb.has_stream("Workbook") {
            cfb.open_stream("Workbook")?
        } else if cfb.has_stream("Book") {
            cfb.open_stream("Book")?
        } else {
            return Err(XlsError::MissingStream("neither Workbook nor Book stream found".into()));
        };
        let has_macros = cfb.has_root_entry("_VBA_PROJECT");
        // Drop CFB early to free file handle and memory.
        drop(cfb);

        let mut doc = Self::parse_workbook_stream(&stream_data)?;
        doc.images = extract_images(&stream_data);
        doc.has_macros = has_macros;
        Ok(doc)
    }

    /// Open an XLS file from a path.
    pub fn open<P: AsRef<std::path::Path>>(path: P) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    fn parse_workbook_stream(data: &[u8]) -> Result<Self> {
        Self::parse_workbook_stream_with_budget(data, 20_000_000)
    }

    /// As `parse_workbook_stream`, with the record-parsing safety cap
    /// (issue #236) as an explicit parameter so tests can exercise the
    /// exhausted-budget path without a multi-million-record fixture.
    fn parse_workbook_stream_with_budget(data: &[u8], record_budget: u32) -> Result<Self> {
        let mut sheet_infos: Vec<SheetInfo> = Vec::new();
        let mut sst: Vec<String> = Vec::new();
        let mut sheets = Vec::new();
        // Number-format tables from the globals substream.
        let mut formats: std::collections::HashMap<u16, String> = std::collections::HashMap::new();
        let mut xf_numfmt: Vec<u16> = Vec::new();
        // [MS-XLS] §2.4.77 DATEMODE: a nonzero `f1904DateSystem` means the
        // workbook's date serials are 1904-based, not 1900-based — a 1462-
        // day (Excel epoch delta) offset if unaccounted for. Defaults to
        // false (1900 system) when absent, matching Excel's own default
        // and every date-serial-carrying record parsed before DATEMODE
        // appears (issue #233).
        let mut date1904 = false;

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
        let mut merged_cells: Vec<(u16, u16, u16, u16)> = Vec::new();
        let mut raw_names: Vec<RawName> = Vec::new();
        let mut supbook_internal: Vec<bool> = Vec::new();
        let mut externsheet: Vec<(u16, i16, i16)> = Vec::new();
        let mut sheet_idx = 0usize;
        let mut pending_formula_string: Option<(u16, u16)> = None;
        // An embedded chart (or other embedded object) is its own nested
        // BOF..EOF substream *inside* the parent worksheet's own substream
        // ([MS-XLS] §2.1.7.20.1, dt=0x0020 for a chart sheet). Without
        // tracking nesting depth, the chart's own closing EOF was mistaken
        // for the worksheet's, truncating the sheet's data and permanently
        // desyncing `sheet_idx` from `sheet_infos` for every sheet after it
        // (issue #237).
        let mut nested_bof_depth = 0u32;
        // Safety cap against a pathologically record-dense file (millions of
        // minimal, near-empty records), not against ordinary large ones.
        // 500,000 was low enough to hit on real, legitimate workbooks —
        // aspose-cells_Sample.xls's 101 sheets truncated to 69 mid-parse
        // with no signal at all, and 3 govdocs1 corpus files (7-29 MB) lost
        // their *entire* content this way (issue #236). The default caller
        // (`parse_workbook_stream`) passes 20,000,000, which still bounds a
        // crafted file's worst-case work (tens of millions of cheap
        // record-type dispatches is well under a second), while
        // comfortably clearing every real file measured so far.
        let mut record_budget = record_budget;
        // Whether the budget ran out before the record stream did — the
        // rest of the workbook (trailing sheets, or all of it) is missing,
        // with no other signal a caller could use to tell that apart from
        // a file that genuinely ended there.
        let mut record_budget_exhausted = false;

        for rec in RecordIter::new(data) {
            if record_budget == 0 {
                record_budget_exhausted = true;
                break;
            }
            record_budget -= 1;
            let rec = rec?;
            match phase {
                Phase::Globals => match rec.record_type {
                    RT_FILEPASS => {
                        // A FILEPASS record means every following record is
                        // encrypted. Returning an empty workbook with `Ok`
                        // told the caller the file simply had no data, which
                        // is indistinguishable from a genuinely empty
                        // spreadsheet.
                        return Err(XlsError::Encrypted);
                    },
                    RT_BOUNDSHEET => {
                        if let Ok(info) = parse_boundsheet(&rec.data) {
                            sheet_infos.push(info);
                        }
                    },
                    RT_SST => {
                        sst = parse_sst(&rec.data)?;
                    },
                    // `FORMAT` maps a format id to its code string; `XF`
                    // maps a cell's `ixfe` to a format id. Both are needed
                    // to tell a date from any other number.
                    RT_FORMAT => {
                        if let Some((id, code)) = parse_format_record(&rec.data) {
                            formats.insert(id, code);
                        }
                    },
                    RT_XF => {
                        // [MS-XLS] §2.4.353: `ifmt` is a u16 at offset 2.
                        xf_numfmt.push(if rec.data.len() >= 4 {
                            u16::from_le_bytes([rec.data[2], rec.data[3]])
                        } else {
                            0
                        });
                    },
                    RT_DATEMODE => {
                        // [MS-XLS] §2.4.77: a single BOOLEAN (u16 LE) at
                        // offset 0, nonzero means the 1904 date system.
                        date1904 = rec.data.len() >= 2
                            && u16::from_le_bytes([rec.data[0], rec.data[1]]) != 0;
                    },
                    RT_NAME => {
                        if let Some(rn) = parse_name_record(&rec.data) {
                            raw_names.push(rn);
                        }
                    },
                    RT_SUPBOOK => {
                        // [MS-XLS] §2.4.271: `cch` at offset 2 (u16) is
                        // 0x0401 exactly for the "self-referencing" SupBook
                        // that names this same workbook, vs. an external
                        // file/DDE/add-in link.
                        supbook_internal.push(
                            rec.data.len() >= 4
                                && u16::from_le_bytes([rec.data[2], rec.data[3]]) == 0x0401,
                        );
                    },
                    RT_EXTERNSHEET => {
                        // [MS-XLS] §2.4.316: cXTI (u16) then cXTI 6-byte XTI
                        // structures (iSupBook u16, itabFirst i16, itabLast
                        // i16).
                        if rec.data.len() >= 2 {
                            let count = u16::from_le_bytes([rec.data[0], rec.data[1]]) as usize;
                            let mut off = 2usize;
                            for _ in 0..count {
                                if off + 6 > rec.data.len() {
                                    break;
                                }
                                let i_sup_book = u16::from_le_bytes([rec.data[off], rec.data[off + 1]]);
                                let itab_first =
                                    i16::from_le_bytes([rec.data[off + 2], rec.data[off + 3]]);
                                let itab_last =
                                    i16::from_le_bytes([rec.data[off + 4], rec.data[off + 5]]);
                                externsheet.push((i_sup_book, itab_first, itab_last));
                                off += 6;
                            }
                        }
                    },
                    RT_EOF => {
                        phase = Phase::BetweenSheets;
                    },
                    _ => {},
                },
                Phase::BetweenSheets => {
                    if rec.record_type == RT_BOF {
                        phase = Phase::InSheet;
                        cells.clear();
                        merged_cells.clear();
                        pending_formula_string = None;
                        nested_bof_depth = 0;
                    }
                },
                Phase::InSheet => match rec.record_type {
                    RT_BOF => {
                        // An embedded chart (or other object) nested inside
                        // this worksheet's own substream — its EOF must not
                        // be mistaken for the worksheet's own.
                        nested_bof_depth += 1;
                    },
                    RT_EOF if nested_bof_depth > 0 => {
                        nested_bof_depth -= 1;
                    },
                    RT_EOF => {
                        let (name, hidden) = match sheet_infos.get(sheet_idx) {
                            Some(info) => (info.name.clone(), info.hidden),
                            None => (format!("Sheet{}", sheet_idx + 1), false),
                        };
                        // A hidden sheet's records were parsed and then
                        // thrown away, so a workbook whose data sat on a
                        // sheet its author merely *hid* came back short with
                        // no error and no notice. Excel round-trips such a
                        // sheet perfectly; keep it and flag it, the way the
                        // XLSX reader already does.
                        let display = build_display(&cells, &formats, &xf_numfmt, date1904);
                        let rows = build_grid(&mut cells);
                        sheets.push(Sheet {
                            name,
                            display,
                            rows,
                            hidden,
                            merged_cells: std::mem::take(&mut merged_cells),
                            ..Default::default()
                        });
                        sheet_idx += 1;
                        phase = Phase::BetweenSheets;
                    },
                    RT_STRING => {
                        if let Some((row, col)) = pending_formula_string.take() {
                            if rec.data.len() >= 3 {
                                if let Ok((s, _)) = read_unicode_string(&rec.data, 0) {
                                    cells.push(Cell {
                                        row,
                                        col,
                                        xf_index: 0,
                                        value: CellValue::String(s),
                                    });
                                }
                            }
                        }
                    },
                    RT_MERGEDCELLS => {
                        // [MS-XLS] §2.4.180: cmcs (u16 count), then cmcs
                        // Ref8 structures (rwFirst, rwLast, colFirst,
                        // colLast — all 0-based u16 LE). A large file can
                        // split this across multiple MERGEDCELLS records;
                        // extending rather than overwriting handles that.
                        if rec.data.len() >= 2 {
                            let count = u16::from_le_bytes([rec.data[0], rec.data[1]]) as usize;
                            let mut off = 2usize;
                            for _ in 0..count {
                                if off + 8 > rec.data.len() {
                                    break;
                                }
                                let row_first = u16::from_le_bytes([rec.data[off], rec.data[off + 1]]);
                                let row_last =
                                    u16::from_le_bytes([rec.data[off + 2], rec.data[off + 3]]);
                                let col_first =
                                    u16::from_le_bytes([rec.data[off + 4], rec.data[off + 5]]);
                                let col_last =
                                    u16::from_le_bytes([rec.data[off + 6], rec.data[off + 7]]);
                                merged_cells.push((row_first, row_last, col_first, col_last));
                                off += 8;
                            }
                        }
                    },
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
                    },
                    _ => {
                        pending_formula_string = None;
                        // Skip LABEL/RSTRING parsing for non-BIFF8 (avoids slow unicode fallback).
                        if !biff8 && matches!(rec.record_type, RT_LABEL | RT_RSTRING) {
                            // BIFF5 LABEL: extract text directly.
                            if rec.data.len() >= 8 {
                                let row = u16::from_le_bytes([rec.data[0], rec.data[1]]);
                                let col = u16::from_le_bytes([rec.data[2], rec.data[3]]);
                                let str_len =
                                    u16::from_le_bytes([rec.data[6], rec.data[7]]) as usize;
                                let start = 8;
                                let end = (start + str_len).min(rec.data.len());
                                let s: String =
                                    rec.data[start..end].iter().map(|&b| b as char).collect();
                                cells.push(Cell {
                                    row,
                                    col,
                                    xf_index: u16::from_le_bytes([rec.data[4], rec.data[5]]),
                                    value: CellValue::String(s),
                                });
                            }
                        } else if let Ok(parsed) = parse_cell_record(&rec, &sst) {
                            cells.extend(parsed);
                        }
                    },
                },
            }
        }

        let defined_names =
            resolve_defined_names(&raw_names, &externsheet, &supbook_internal, &sheet_infos);

        Ok(Self {
            sheets,
            defined_names,
            images: Vec::new(),
            // Set by the caller (from_reader), which has the CfbReader
            // this function doesn't.
            has_macros: false,
            truncated: record_budget_exhausted,
        })
    }

    /// Get all extracted images.
    pub fn images(&self) -> &[XlsImage] {
        &self.images
    }

    /// `true` when the file carries a `_VBA_PROJECT` storage — a cheap
    /// macro-presence signal, no VBA interpretation (issue #283).
    pub fn has_macros(&self) -> bool {
        self.has_macros
    }

    /// `true` when the record-parsing safety cap cut the Workbook stream
    /// short — trailing sheets, or the whole workbook, are missing from
    /// `sheets` (issue #236).
    pub fn truncated(&self) -> bool {
        self.truncated
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
            for (r, row) in sheet.rows.iter().enumerate() {
                let line: Vec<String> =
                    (0..row.len()).map(|c| cell_display_text(sheet, r, c)).collect();
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
            if !sheet.rows.is_empty() {
                for c in 0..col_count {
                    let text = cell_display_text(sheet, 0, c);
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
            for (r, _) in sheet.rows.iter().enumerate().skip(1) {
                out.push('|');
                for c in 0..col_count {
                    let text = cell_display_text(sheet, r, c);
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

/// Detect a raw (non-CFB) BIFF2/3/4 stream, returning the BIFF generation.
///
/// The reader is left at the position it was handed over at, so a negative
/// result costs the caller nothing. A BIFF2-4 file opens with a `BOF` whose
/// `SID` names the generation — `0x0009` (BIFF2), `0x0209` (BIFF3),
/// `0x0409` (BIFF4) — followed by that record's own small length, which is
/// what distinguishes it from arbitrary bytes that happen to collide.
fn detect_raw_biff<R: Read + Seek>(reader: &mut R) -> Result<Option<u8>> {
    let start = reader.stream_position()?;
    let mut head = [0u8; 4];
    let read = fill(reader, &mut head)?;
    reader.seek(std::io::SeekFrom::Start(start))?;
    if read < 4 {
        return Ok(None);
    }
    let sid = u16::from_le_bytes([head[0], head[1]]);
    let len = u16::from_le_bytes([head[2], head[3]]);
    // A BIFF2-4 BOF body is 4-8 bytes (version + stream type, plus build
    // fields from BIFF4); anything longer is not one.
    if len > 16 {
        return Ok(None);
    }
    Ok(match sid {
        0x0009 => Some(2),
        0x0209 => Some(3),
        0x0409 => Some(4),
        _ => None,
    })
}

/// Read as much of `buf` as the reader has, returning how many bytes landed.
fn fill<R: Read>(reader: &mut R, buf: &mut [u8]) -> std::io::Result<usize> {
    let mut filled = 0;
    while filled < buf.len() {
        match reader.read(&mut buf[filled..])? {
            0 => break,
            n => filled += n,
        }
    }
    Ok(filled)
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

/// A `NAME` record ([MS-XLS] §2.4.174) before its `rgce` formula has been
/// resolved against `EXTERNSHEET`/`SUPBOOK` (those records can appear
/// before or after `NAME` in the Globals substream, so resolution happens
/// once the whole substream has been read).
struct RawName {
    name: String,
    /// 0 = workbook-global; otherwise a 1-based index into `sheet_infos`.
    itab: u16,
    hidden: bool,
    rgce: Vec<u8>,
}

/// Parse a `NAME` record (`0x0018`).
///
/// Layout: grbit(2) chKey(1) cch(1) cce(2) reserved(2) itab(2) reserved(4)
/// then Name as `XLUnicodeStringNoCch` (1 flag byte + `cch` chars, 1 or 2
/// bytes each), then `rgce` (`cce` bytes) — confirmed against the worked
/// example in [MS-XLS] §2.5.155.
fn parse_name_record(data: &[u8]) -> Option<RawName> {
    if data.len() < 14 {
        return None;
    }
    let grbit = u16::from_le_bytes([data[0], data[1]]);
    let is_builtin = (grbit & 0x0020) != 0; // fBuiltin, [MS-XLS] §2.5.155
    let cch = data[3] as usize;
    let cce = u16::from_le_bytes([data[4], data[5]]) as usize;
    let itab = u16::from_le_bytes([data[8], data[9]]);

    let mut pos = 14usize;
    if pos >= data.len() {
        return None;
    }
    let is_wide = (data[pos] & 0x01) != 0;
    let first_content_byte = data.get(pos + 1).copied();
    pos += 1;
    let name_bytes = if is_wide { cch * 2 } else { cch };
    if pos + name_bytes > data.len() {
        return None;
    }
    // A built-in name (Print_Area, _FilterDatabase, ...) stores its `Name`
    // field as a single raw byte ID, not text — decoded as a character it
    // came out as an unreadable control code (`Print_Area` -> U+0006).
    // XLSX's reader already surfaces these with their OOXML
    // `_xlnm.`-prefixed reserved names, so map the ID the same way rather
    // than leaving the two formats' output inconsistent.
    let name = if is_builtin && cch == 1 {
        first_content_byte.and_then(builtin_name).unwrap_or_default().to_string()
    } else if is_wide {
        let chars: Vec<u16> = data[pos..pos + name_bytes]
            .chunks_exact(2)
            .map(|c| u16::from_le_bytes([c[0], c[1]]))
            .collect();
        String::from_utf16_lossy(&chars)
    } else {
        data[pos..pos + name_bytes].iter().map(|&b| b as char).collect()
    };
    pos += name_bytes;

    if pos + cce > data.len() {
        return None;
    }
    let rgce = data[pos..pos + cce].to_vec();

    Some(RawName {
        name,
        itab,
        hidden: (grbit & 0x0001) != 0,
        rgce,
    })
}

/// Map a `NAME` record's built-in ID byte ([MS-XLS] §2.5.28 `BuiltInName`)
/// to the same `_xlnm.`-prefixed reserved name OOXML/XLSX uses, so a
/// defined name means the same thing regardless of which format it came
/// from.
fn builtin_name(id: u8) -> Option<&'static str> {
    Some(match id {
        0x00 => "_xlnm.Consolidate_Area",
        0x01 => "_xlnm.Auto_Open",
        0x02 => "_xlnm.Auto_Close",
        0x03 => "_xlnm.Extract",
        0x04 => "_xlnm.Database",
        0x05 => "_xlnm.Criteria",
        0x06 => "_xlnm.Print_Area",
        0x07 => "_xlnm.Print_Titles",
        0x08 => "_xlnm.Recorder",
        0x09 => "_xlnm.Data_Form",
        0x0A => "_xlnm.Auto_Activate",
        0x0B => "_xlnm.Auto_Deactivate",
        0x0C => "_xlnm.Sheet_Title",
        0x0D => "_xlnm._FilterDatabase",
        _ => return None,
    })
}

/// Render a 0-based column index as spreadsheet letters (0 -> "A", 25 ->
/// "Z", 26 -> "AA").
fn col_letters(mut col: u32) -> String {
    let mut s = Vec::new();
    loop {
        let rem = (col % 26) as u8;
        s.push(b'A' + rem);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    s.reverse();
    String::from_utf8(s).unwrap()
}

/// Resolve `NAME` records' `rgce` formulas into `Sheet!$A$1:$B$10`-style
/// text.
///
/// Only a formula that is a single `PtgRef3d` (`0x1A`) or `PtgArea3d`
/// (`0x1B`) token — the shape every plain named range compiles to — is
/// decompiled; anything else (unions, functions, multi-token formulas)
/// resolves with an empty `value` rather than a guessed, possibly wrong,
/// one. Both token kinds carry an `ixti` that indexes `externsheet`, whose
/// `(supbook_idx, first_sheet, last_sheet)` resolves to a sheet name only
/// when `supbook_idx` names *this* workbook (`supbook_internal`) — an
/// external-workbook reference is left unresolved rather than misreported
/// as local.
fn resolve_defined_names(
    raw_names: &[RawName],
    externsheet: &[(u16, i16, i16)],
    supbook_internal: &[bool],
    sheet_infos: &[SheetInfo],
) -> Vec<DefinedName> {
    let sheet_name_for = |itab_first: i16, itab_last: i16, sup: u16| -> Option<String> {
        if !*supbook_internal.get(sup as usize)? {
            return None;
        }
        if itab_first < 0 || itab_last < 0 {
            return None;
        }
        let first = sheet_infos.get(itab_first as usize)?;
        if itab_first == itab_last {
            Some(first.name.clone())
        } else {
            let last = sheet_infos.get(itab_last as usize)?;
            Some(format!("{}:{}", first.name, last.name))
        }
    };

    let cell_ref = |row: u16, col_raw: u16| -> String {
        format!("${}${}", col_letters((col_raw & 0x3FFF) as u32), row + 1)
    };

    raw_names
        .iter()
        .map(|rn| {
            let value = decompile_single_ref(&rn.rgce, externsheet, &sheet_name_for, &cell_ref)
                .unwrap_or_default();
            DefinedName {
                name: rn.name.clone(),
                value,
                local_sheet_id: if rn.itab == 0 {
                    None
                } else {
                    Some((rn.itab - 1) as u32)
                },
                hidden: rn.hidden,
            }
        })
        .collect()
}

/// Decompile `rgce` when it is exactly one `PtgRef3d`/`PtgArea3d` token.
fn decompile_single_ref(
    rgce: &[u8],
    externsheet: &[(u16, i16, i16)],
    sheet_name_for: &dyn Fn(i16, i16, u16) -> Option<String>,
    cell_ref: &dyn Fn(u16, u16) -> String,
) -> Option<String> {
    if rgce.is_empty() {
        return None;
    }
    let base_ptg = rgce[0] & 0x1F;
    let ixti_and = |off: usize| -> Option<u16> {
        rgce.get(off).zip(rgce.get(off + 1)).map(|(a, b)| u16::from_le_bytes([*a, *b]))
    };

    match base_ptg {
        // PtgRef3d: ptg(1) ixti(2) row(2) col(2) = 7 bytes total.
        0x1A if rgce.len() == 7 => {
            let ixti = ixti_and(1)?;
            let (sup, first, last) = *externsheet.get(ixti as usize)?;
            let sheet = sheet_name_for(first, last, sup)?;
            let row = u16::from_le_bytes([rgce[3], rgce[4]]);
            let col = u16::from_le_bytes([rgce[5], rgce[6]]);
            Some(format!("{sheet}!{}", cell_ref(row, col)))
        },
        // PtgArea3d: ptg(1) ixti(2) rwFirst(2) rwLast(2) colFirst(2)
        // colLast(2) = 11 bytes total.
        0x1B if rgce.len() == 11 => {
            let ixti = ixti_and(1)?;
            let (sup, first, last) = *externsheet.get(ixti as usize)?;
            let sheet = sheet_name_for(first, last, sup)?;
            let rw_first = u16::from_le_bytes([rgce[3], rgce[4]]);
            let rw_last = u16::from_le_bytes([rgce[5], rgce[6]]);
            let col_first = u16::from_le_bytes([rgce[7], rgce[8]]);
            let col_last = u16::from_le_bytes([rgce[9], rgce[10]]);
            Some(format!(
                "{sheet}!{}:{}",
                cell_ref(rw_first, col_first),
                cell_ref(rw_last, col_last)
            ))
        },
        _ => None,
    }
}

/// Build a 2D grid from sparse cells.
///
/// Takes ownership of cell values via `std::mem::take` to avoid cloning.
/// Parse a `FORMAT` record: `ifmt` (u16) then the format code as a
/// BIFF8 unicode string.
fn parse_format_record(data: &[u8]) -> Option<(u16, String)> {
    if data.len() < 4 {
        return None;
    }
    let id = u16::from_le_bytes([data[0], data[1]]);
    let (code, _) = read_unicode_string(data, 2).ok()?;
    Some((id, code))
}

/// The format-aware text for `sheet.rows[r][c]`: `sheet.display[r][c]` when
/// present, falling back to the cell's raw `CellValue::as_text()`.
///
/// `Document::plain_text()`/`to_markdown()` dispatch to `XlsDocument`'s own
/// `plain_text()`/`to_markdown()` below, a separate path from `to_ir()`
/// (`convert_xls.rs`, which already read `sheet.display` correctly). Those
/// two methods read `sheet.rows` directly via `CellValue::as_text()`,
/// bypassing `build_display`'s number-format/date rendering entirely — a
/// date cell came out as a raw serial (`38971`) from the CLI's default
/// `text`/`markdown` output even though `to_ir()` got it right, the same
/// dual-renderer gap fixed for XLSX formula text in #279. The fallback to
/// raw text covers the (normally unreachable) case where `display` and
/// `rows` disagree in shape.
fn cell_display_text(sheet: &Sheet, r: usize, c: usize) -> String {
    sheet.display.get(r).and_then(|row| row.get(c)).cloned().unwrap_or_else(|| {
        sheet.rows.get(r).and_then(|row| row.get(c)).map(CellValue::as_text).unwrap_or_default()
    })
}

/// Render each cell's display text, applying the workbook's number formats.
///
/// Mirrors the XLSX side: a number whose format is a date format renders as
/// an ISO date, and any other non-General format is applied to the value.
/// Everything else falls back to `CellValue::as_text`.
fn build_display(
    cells: &[Cell],
    formats: &std::collections::HashMap<u16, String>,
    xf_numfmt: &[u16],
    date1904: bool,
) -> Vec<Vec<String>> {
    use crate::xlsx::{date, numfmt};

    let format_for = |xf: u16| -> Option<(u16, Option<&str>)> {
        let fmt_id = *xf_numfmt.get(xf as usize)?;
        Some((fmt_id, formats.get(&fmt_id).map(|s| s.as_str())))
    };

    let mut grid: Vec<Vec<String>> = Vec::new();
    for cell in cells {
        let text = match &cell.value {
            CellValue::Number(n) => match format_for(cell.xf_index) {
                Some((fmt_id, code)) => {
                    // A workbook may redefine a built-in id — [ECMA-376]
                    // §18.8.30 permits ids 0-163 to be overridden — so an
                    // explicit `FORMAT` code wins over the built-in meaning
                    // of its id. Testing the id first classified a
                    // scientific-notation cell whose id-50 format the file
                    // overrode to `0.00000E+0` as a date, and the date
                    // conversion then ran for minutes on its magnitude.
                    // Mirrors `date::is_date_cell`, fixed the same way in
                    // #207.
                    let is_date = match code {
                        Some(c) => date::is_date_format_string(c),
                        None => date::is_date_format_id(fmt_id as u32),
                    };
                    if is_date {
                        match date::DateTimeValue::from_serial(*n, date1904) {
                            Some(dt) => dt.to_iso_string(),
                            None => cell.value.as_text(),
                        }
                    } else if fmt_id != 0 {
                        numfmt::apply_format(*n, fmt_id as u32, code)
                    } else {
                        cell.value.as_text()
                    }
                },
                None => cell.value.as_text(),
            },
            other => other.as_text(),
        };
        let (r, c) = (cell.row as usize, cell.col as usize);
        if r > 65535 || c > 255 {
            continue;
        }
        if grid.len() <= r {
            grid.resize(r + 1, Vec::new());
        }
        if grid[r].len() <= c {
            grid[r].resize(c + 1, String::new());
        }
        grid[r][c] = text;
    }
    grid
}

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
    while grid
        .last()
        .is_some_and(|row| row.iter().all(|c| matches!(c, CellValue::Empty)))
    {
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
    while grid
        .last()
        .is_some_and(|row| row.iter().all(|c| matches!(c, CellValue::Empty)))
    {
        grid.pop();
    }

    grid
}

impl crate::core::OfficeDocument for XlsDocument {
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

    // ── Hand-built BIFF8 fixtures ─────────────────────────────────────────
    //
    // The record-walking state machine is the part several of these tests
    // exercise, and a minimal stream built here is far easier to reason
    // about (and to keep in the repo) than a real workbook.

    /// Wrap `data` in a BIFF record header.
    pub(super) fn biff_rec(rt: u16, data: &[u8]) -> Vec<u8> {
        let mut v = rt.to_le_bytes().to_vec();
        v.extend_from_slice(&(data.len() as u16).to_le_bytes());
        v.extend_from_slice(data);
        v
    }

    /// A BIFF8 `BOF` opening a substream of the given doctype.
    pub(super) fn bof(dt: u16) -> Vec<u8> {
        let mut d = Vec::new();
        d.extend_from_slice(&0x0600u16.to_le_bytes()); // BIFF8
        d.extend_from_slice(&dt.to_le_bytes());
        d.extend_from_slice(&[0u8; 12]);
        biff_rec(RT_BOF, &d)
    }

    pub(super) fn eof() -> Vec<u8> {
        biff_rec(RT_EOF, &[])
    }

    /// `BOUNDSHEET`: 0 = visible, 1 = hidden, 2 = very hidden.
    pub(super) fn boundsheet(name: &str, visibility: u8) -> Vec<u8> {
        let mut d = Vec::new();
        d.extend_from_slice(&0u32.to_le_bytes()); // stream offset, unused here
        d.push(visibility);
        d.push(0); // worksheet
        d.push(name.len() as u8);
        d.push(0); // 8-bit (compressed) characters
        d.extend_from_slice(name.as_bytes());
        biff_rec(RT_BOUNDSHEET, &d)
    }

    pub(super) fn number(row: u16, col: u16, xf: u16, value: f64) -> Vec<u8> {
        let mut d = Vec::new();
        d.extend_from_slice(&row.to_le_bytes());
        d.extend_from_slice(&col.to_le_bytes());
        d.extend_from_slice(&xf.to_le_bytes());
        d.extend_from_slice(&value.to_le_bytes());
        biff_rec(RT_NUMBER, &d)
    }

    /// A BIFF8 `LABEL` (inline 8-bit string) cell.
    pub(super) fn label(row: u16, col: u16, text: &str) -> Vec<u8> {
        let mut d = Vec::new();
        d.extend_from_slice(&row.to_le_bytes());
        d.extend_from_slice(&col.to_le_bytes());
        d.extend_from_slice(&0u16.to_le_bytes()); // xf
        d.extend_from_slice(&(text.chars().count() as u16).to_le_bytes());
        d.push(0); // 8-bit characters
        d.extend_from_slice(text.as_bytes());
        biff_rec(RT_LABEL, &d)
    }

    /// Assemble a workbook stream: a globals substream carrying one
    /// `BOUNDSHEET` per sheet, then each sheet's own `BOF..EOF` substream.
    pub(super) fn workbook_stream(sheets: &[(&str, u8, Vec<u8>)]) -> Vec<u8> {
        workbook_stream_with_globals(&[], sheets)
    }

    /// As `workbook_stream`, with extra records appended to the globals
    /// substream (`DATEMODE`, `CODEPAGE`, `NAME`, ...).
    pub(super) fn workbook_stream_with_globals(
        globals: &[Vec<u8>],
        sheets: &[(&str, u8, Vec<u8>)],
    ) -> Vec<u8> {
        let mut s = bof(0x0005);
        for (name, visibility, _) in sheets {
            s.extend(boundsheet(name, *visibility));
        }
        for g in globals {
            s.extend_from_slice(g);
        }
        s.extend(eof());
        for (_, _, body) in sheets {
            s.extend(bof(0x0010));
            s.extend_from_slice(body);
            s.extend(eof());
        }
        s
    }

    /// A hidden sheet's records were parsed and then discarded, so the
    /// sheet vanished from `to_ir()` entirely — silent deletion of data its
    /// author only hid. XLSX keeps and flags such a sheet; XLS now does
    /// too (#231).
    #[test]
    fn test_xls_hidden_sheets_kept_and_flagged() {
        let stream = workbook_stream(&[
            ("Sheet1", 1, label(0, 0, "Sheet1A1")),   // hidden
            ("Sheet2", 0, label(0, 0, "Sheet2A1")),   // visible
            ("Sheet3", 2, label(0, 0, "Sheet3A1")),   // very hidden
        ]);
        let doc = XlsDocument::parse_workbook_stream(&stream).expect("parses");

        assert_eq!(doc.sheets.len(), 3, "no sheet may be dropped for being hidden");
        assert_eq!(
            doc.sheets.iter().map(|s| s.name.as_str()).collect::<Vec<_>>(),
            ["Sheet1", "Sheet2", "Sheet3"]
        );
        assert_eq!(
            doc.sheets.iter().map(|s| s.hidden).collect::<Vec<_>>(),
            [true, false, true],
            "visibility is reported, not acted on"
        );
        // The hidden sheets' data survives.
        let text = doc.plain_text();
        for expected in ["Sheet1A1", "Sheet2A1", "Sheet3A1"] {
            assert!(text.contains(expected), "lost {expected} from: {text}");
        }

        // ...and the flag reaches the IR.
        let ir = crate::convert_xls::xls_to_ir(&doc);
        assert_eq!(ir.sections.len(), 3);
        assert_eq!(
            ir.sections.iter().map(|s| s.hidden).collect::<Vec<_>>(),
            [true, false, true]
        );
    }

    /// An embedded chart is its own nested `BOF..EOF` substream inside the
    /// parent worksheet's substream. Before #237, its `EOF` was mistaken
    /// for the worksheet's own: the sheet's remaining cells were dropped,
    /// `sheet_idx` desynced from `sheet_infos`, and every sheet after it
    /// got the wrong name/visibility.
    #[test]
    fn test_embedded_chart_eof_does_not_close_the_parent_sheet() {
        let mut sheet1_body = label(0, 0, "before_chart");
        sheet1_body.extend(bof(0x0020)); // chart substream, nested inside Sheet1
        sheet1_body.extend(label(0, 1, "inside_chart"));
        sheet1_body.extend(eof()); // chart's own EOF — must not close Sheet1
        sheet1_body.extend(label(1, 0, "after_chart"));

        let stream = workbook_stream(&[
            ("Sheet1", 0, sheet1_body),
            ("Sheet2", 1, label(0, 0, "Sheet2A1")), // hidden
        ]);
        let doc = XlsDocument::parse_workbook_stream(&stream).expect("parses");

        assert_eq!(
            doc.sheets.len(),
            2,
            "the chart's nested EOF must not fabricate an extra sheet, got {:?}",
            doc.sheets.iter().map(|s| &s.name).collect::<Vec<_>>()
        );
        assert_eq!(doc.sheets[0].name, "Sheet1");
        assert!(!doc.sheets[0].hidden);
        assert_eq!(doc.sheets[1].name, "Sheet2");
        assert!(doc.sheets[1].hidden, "Sheet2's real hidden flag must survive, not desync to false");

        let text = doc.plain_text();
        assert!(text.contains("before_chart"), "text: {text}");
        assert!(
            text.contains("after_chart"),
            "cells after the embedded chart must not be dropped, text: {text}"
        );
        assert!(text.contains("Sheet2A1"), "text: {text}");
    }

    // ── Record-parsing safety cap (issue #236) ──────────────────────────────

    /// A budget of 0 must not panic or hang — just truncate before any
    /// record is processed at all.
    #[test]
    fn test_zero_budget_truncates_before_any_record_and_is_flagged() {
        let stream = workbook_stream(&[("Sheet1", 0, label(0, 0, "x"))]);
        let doc = XlsDocument::parse_workbook_stream_with_budget(&stream, 0)
            .expect("a zero budget must not error, just truncate");
        assert!(doc.sheets.is_empty());
        assert!(doc.truncated(), "a zero budget must be flagged as truncated");
    }

    /// A budget that runs out exactly between two sheets (globals: BOF +
    /// 2×BOUNDSHEET + EOF = 4 records, then each sheet: BOF + LABEL + EOF =
    /// 3 records) must keep the first sheet intact, drop the second
    /// entirely, and flag the workbook as truncated — mirroring the real
    /// corpus files the issue cites (aspose-cells_Sample.xls: 101 sheets
    /// truncated to 69; several govdocs1 files: entire workbook lost).
    #[test]
    fn test_budget_exhausted_between_sheets_keeps_earlier_sheets_and_flags_truncation() {
        let stream = workbook_stream(&[
            ("Sheet1", 0, label(0, 0, "Sheet1A1")),
            ("Sheet2", 0, label(0, 0, "Sheet2A1")),
        ]);
        let doc = XlsDocument::parse_workbook_stream_with_budget(&stream, 7)
            .expect("parses up to the cap");
        assert_eq!(doc.sheets.len(), 1, "Sheet1 must survive, Sheet2 must be dropped entirely");
        assert_eq!(doc.sheets[0].name, "Sheet1");
        assert!(doc.truncated(), "a mid-workbook cutoff must be flagged as truncated");

        // A budget generous enough to cover both sheets must not flag
        // truncation at all.
        let doc_full =
            XlsDocument::parse_workbook_stream_with_budget(&stream, 20_000_000).expect("parses");
        assert_eq!(doc_full.sheets.len(), 2);
        assert!(!doc_full.truncated());
    }

    /// The truncation flag reaches `to_ir()`'s `Metadata::text_truncated`,
    /// the same signal DOC's piece-table gap (#230) already surfaces.
    #[test]
    fn test_truncation_reaches_metadata() {
        let stream = workbook_stream(&[
            ("Sheet1", 0, label(0, 0, "Sheet1A1")),
            ("Sheet2", 0, label(0, 0, "Sheet2A1")),
        ]);
        let doc = XlsDocument::parse_workbook_stream_with_budget(&stream, 7).expect("parses");
        let ir = crate::convert_xls::xls_to_ir(&doc);
        assert!(ir.metadata.text_truncated);
    }

    #[test]
    fn build_grid_from_cells() {
        let mut cells = vec![
            Cell {
                xf_index: 0,
                row: 0,
                col: 0,
                value: CellValue::String("A1".into()),
            },
            Cell {
                xf_index: 0,
                row: 0,
                col: 1,
                value: CellValue::Number(42.0),
            },
            Cell {
                xf_index: 0,
                row: 1,
                col: 0,
                value: CellValue::String("A2".into()),
            },
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
            has_macros: false,
            defined_names: Vec::new(),
            truncated: false,
            sheets: vec![Sheet {
                display: Vec::new(),
                name: "Sheet1".into(),
                rows: vec![
                    vec![
                        CellValue::String("Name".into()),
                        CellValue::String("Age".into()),
                    ],
                    vec![CellValue::String("Alice".into()), CellValue::Number(30.0)],
                ],
                ..Default::default()
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
            has_macros: false,
            defined_names: Vec::new(),
            truncated: false,
            sheets: vec![Sheet {
                display: Vec::new(),
                name: "Data".into(),
                rows: vec![
                    vec![CellValue::String("X".into()), CellValue::String("Y".into())],
                    vec![CellValue::Number(1.0), CellValue::Number(2.0)],
                ],
                ..Default::default()
            }],
        };
        let md = doc.to_markdown();
        assert!(md.contains("## Data"));
        assert!(md.contains("| X | Y |"));
        assert!(md.contains("| --- | --- |"));
        assert!(md.contains("| 1 | 2 |"));
    }

    /// `Document::plain_text()`/`to_markdown()` dispatch to these two
    /// methods, a separate path from `to_ir()` — which already read
    /// `sheet.display` correctly. Before this fix these methods read
    /// `sheet.rows` directly (raw `CellValue::as_text()`), so a date cell
    /// showed its raw serial (`38971`) here even when `to_ir()`/`markdown`
    /// via the shared IR renderer got it right — the same dual-renderer
    /// gap already hit for XLSX formula text (#279).
    #[test]
    fn plain_text_and_markdown_prefer_the_format_aware_display_text() {
        let sheet = Sheet {
            name: "Sheet1".into(),
            rows: vec![vec![CellValue::Number(38971.0)]],
            display: vec![vec!["2006-09-11".to_string()]],
            ..Default::default()
        };
        let doc = make_doc(vec![sheet]);
        assert!(
            doc.plain_text().contains("2006-09-11"),
            "plain_text() must use the formatted date, got: {:?}",
            doc.plain_text()
        );
        assert!(
            !doc.plain_text().contains("38971"),
            "plain_text() must not fall back to the raw serial when display is populated"
        );
        assert!(
            doc.to_markdown().contains("2006-09-11"),
            "to_markdown() must use the formatted date, got: {:?}",
            doc.to_markdown()
        );
    }

    fn make_doc(sheets: Vec<Sheet>) -> XlsDocument {
        XlsDocument {
            images: Vec::new(),
            has_macros: false,
            defined_names: Vec::new(),
            truncated: false,
            sheets,
        }
    }

    #[test]
    fn ir_empty_doc_produces_no_sections() {
        let ir = crate::convert_xls::xls_to_ir(&make_doc(vec![]));
        assert!(ir.sections.is_empty());
        assert!(ir.metadata.title.is_none());
    }

    #[test]
    fn ir_empty_sheet_has_no_table() {
        let ir = crate::convert_xls::xls_to_ir(&make_doc(vec![Sheet {
            display: Vec::new(),
            name: "Empty".into(),
            rows: vec![],
            ..Default::default()
        }]));
        assert_eq!(ir.sections[0].title.as_deref(), Some("Empty"));
        assert!(ir.sections[0].elements.is_empty());
    }

    #[test]
    fn ir_sheet_with_data_produces_table_with_header_row() {
        use crate::ir::Element;
        let rows = vec![
            vec![
                CellValue::String("Name".into()),
                CellValue::String("Score".into()),
            ],
            vec![CellValue::String("Alice".into()), CellValue::Number(95.0)],
        ];
        let ir = crate::convert_xls::xls_to_ir(&make_doc(vec![Sheet {
            display: Vec::new(),
            name: "Results".into(),
            rows,
            ..Default::default()
        }]));
        assert_eq!(ir.metadata.title.as_deref(), Some("Results"));
        assert!(matches!(ir.sections[0].elements[0], Element::Table(_)));
        if let Element::Table(ref t) = ir.sections[0].elements[0] {
            assert!(t.rows[0].is_header);
            assert!(!t.rows[1].is_header);
        }
    }

    /// An empty cell *between* populated ones keeps its column position and
    /// renders as an empty paragraph. Only trailing empties are trimmed.
    #[test]
    fn ir_empty_cell_value_produces_empty_paragraph_content() {
        use crate::ir::Element;
        let ir = crate::convert_xls::xls_to_ir(&make_doc(vec![Sheet {
            display: Vec::new(),
            name: "S".into(),
            rows: vec![vec![
                CellValue::String("a".into()),
                CellValue::Empty,
                CellValue::String("b".into()),
            ]],
            ..Default::default()
        }]));
        let Element::Table(ref t) = ir.sections[0].elements[0] else {
            panic!("expected a table");
        };
        assert_eq!(t.rows[0].cells.len(), 3, "the interior empty keeps its column");
        let Element::Paragraph(ref p) = t.rows[0].cells[1].content[0] else {
            panic!("expected a paragraph");
        };
        assert!(p.content.is_empty());
    }

    /// A BIFF sheet reports its whole declared grid, which for one real file
    /// is 65,536 x 256 and essentially all empty. Materialising that padding
    /// built 16.7M IR cells and ran the process out of memory, so trailing
    /// empty cells and rows are dropped — the same trim `convert_xlsx` has
    /// always done.
    #[test]
    fn ir_trailing_empty_cells_and_rows_are_trimmed() {
        use crate::ir::Element;
        let ir = crate::convert_xls::xls_to_ir(&make_doc(vec![Sheet {
            display: Vec::new(),
            name: "S".into(),
            rows: vec![
                vec![
                    CellValue::String("x".into()),
                    CellValue::Empty,
                    CellValue::Empty,
                ],
                vec![CellValue::Empty, CellValue::Empty],
            ],
            ..Default::default()
        }]));
        let Element::Table(ref t) = ir.sections[0].elements[0] else {
            panic!("expected a table");
        };
        assert_eq!(t.rows.len(), 1, "the all-empty trailing row is dropped");
        assert_eq!(t.rows[0].cells.len(), 1, "trailing empty cells are dropped");
    }

    #[test]
    fn ir_a_wholly_empty_sheet_produces_no_table() {
        let ir = crate::convert_xls::xls_to_ir(&make_doc(vec![Sheet {
            display: Vec::new(),
            name: "S".into(),
            rows: vec![vec![CellValue::Empty; 8]; 4],
            ..Default::default()
        }]));
        assert!(ir.sections[0].elements.is_empty(), "an empty grid must not materialise a table");
    }

    #[test]
    fn ir_multiple_sheets_produce_multiple_sections() {
        let doc = make_doc(vec![
            Sheet {
                display: Vec::new(),
                name: "A".into(),
                rows: vec![vec![CellValue::Number(1.0)]],
                ..Default::default()
            },
            Sheet {
                display: Vec::new(),
                name: "B".into(),
                rows: vec![vec![CellValue::String("x".into())]],
                ..Default::default()
            },
        ]);
        let ir = crate::convert_xls::xls_to_ir(&doc);
        assert_eq!(ir.sections.len(), 2);
        assert_eq!(ir.sections[1].title.as_deref(), Some("B"));
    }

    /// A raw BIFF2/3/4 stream has no CFB container, so the CFB layer used
    /// to reject it with "bad magic signature" — indistinguishable from a
    /// corrupt or unrelated file. Name the format instead (#227).
    #[test]
    fn test_raw_biff2_4_reports_the_unsupported_legacy_format() {
        for (sid, label) in [(0x0009u16, "BIFF2"), (0x0209, "BIFF3"), (0x0409, "BIFF4")] {
            let mut stream = Vec::new();
            stream.extend_from_slice(&sid.to_le_bytes());
            stream.extend_from_slice(&6u16.to_le_bytes()); // BOF body length
            stream.extend_from_slice(&[0x00, 0x04, 0x10, 0x00, 0x00, 0x00]);
            stream.extend_from_slice(&RT_EOF.to_le_bytes());
            stream.extend_from_slice(&0u16.to_le_bytes());

            let err = XlsDocument::from_reader(std::io::Cursor::new(stream))
                .expect_err("a pre-OLE2 file cannot be read");
            let msg = err.to_string();
            assert!(
                matches!(err, XlsError::UnsupportedVersion(_)),
                "expected UnsupportedVersion for {label}, got: {msg}"
            );
            assert!(msg.contains(label), "error must name {label}: {msg}");
            assert!(!msg.contains("magic"), "the CFB-layer message must not leak: {msg}");
        }
    }

    /// The BIFF2-4 sniff must not claim an ordinary BIFF8 `.xls`, whose
    /// container starts with the CFB signature.
    #[test]
    fn test_biff2_4_detection_ignores_a_cfb_container() {
        let mut stream = crate::cfb::CFB_SIGNATURE.to_vec();
        stream.resize(512, 0);
        let mut cursor = std::io::Cursor::new(stream);
        assert_eq!(detect_raw_biff(&mut cursor).unwrap(), None);
        // ...and the reader is left where it was handed over.
        assert_eq!(cursor.position(), 0);
    }

    /// A workbook may override a built-in `numFmtId`; ECMA-376 §18.8.30
    /// permits it for ids 0-163, and a real Gnumeric file redefines id 50
    /// (nominally a locale date) as `0.00000E+0`. Testing the id before the
    /// declared code classified such a cell as a date and handed its
    /// magnitude to the calendar walk, which ran for minutes (#225).
    #[test]
    fn test_xls_overridden_builtin_date_id_is_not_treated_as_a_date() {
        let mut formats = std::collections::HashMap::new();
        formats.insert(50u16, "0.00000E+0".to_string());
        // xf 0 -> fmt id 50.
        let xf_numfmt = vec![50u16];

        let cells = vec![Cell {
            xf_index: 0,
            row: 0,
            col: 0,
            value: CellValue::Number(4.052_85e199),
        }];

        let started = std::time::Instant::now();
        let display = build_display(&cells, &formats, &xf_numfmt, false);
        assert!(
            started.elapsed() < std::time::Duration::from_secs(2),
            "build_display took {:?}",
            started.elapsed()
        );
        assert!(
            !display[0][0].starts_with("1900-") && !display[0][0].starts_with("9999-"),
            "the overridden format is not a date: {}",
            display[0][0]
        );
        assert!(
            display[0][0].contains('E'),
            "expected scientific notation, got {}",
            display[0][0]
        );

        // A built-in date id with *no* declared override is still a date.
        let cells = vec![Cell {
            xf_index: 0,
            row: 0,
            col: 0,
            value: CellValue::Number(38971.0),
        }];
        let display = build_display(&cells, &std::collections::HashMap::new(), &[14u16], false);
        assert_eq!(display[0][0], "2006-09-11");
    }

    /// The same date serial renders a different calendar date depending on
    /// `date1904` — the 1900/1904 epoch delta is exactly 1462 days. Before
    /// #233, `build_display` had no `date1904` parameter at all and always
    /// rendered as if the workbook were 1900-mode.
    #[test]
    fn test_build_display_honours_the_date1904_flag() {
        let cells = vec![Cell {
            xf_index: 0,
            row: 0,
            col: 0,
            value: CellValue::Number(38971.0),
        }];
        let display_1900 =
            build_display(&cells, &std::collections::HashMap::new(), &[14u16], false);
        let display_1904 =
            build_display(&cells, &std::collections::HashMap::new(), &[14u16], true);
        assert_eq!(display_1900[0][0], "2006-09-11");
        assert_ne!(
            display_1900[0][0], display_1904[0][0],
            "the same serial must render differently under the 1904 date system"
        );
    }

    /// End-to-end: a real `DATEMODE` record (`0x0022`) in the globals
    /// substream must reach the cell that renders the date, via the full
    /// `parse_workbook_stream` record walk — not just `build_display`
    /// called directly (issue #233).
    #[test]
    fn test_datemode_record_reaches_the_rendered_cell() {
        // ifmt=14 (built-in "m/d/yyyy") at offset 2 of a minimal XF record.
        let xf_date = biff_rec(RT_XF, &[0, 0, 14, 0]);
        let cell = number(0, 0, 0, 38971.0);

        let globals_1900 = vec![xf_date.clone()];
        let stream_1900 =
            workbook_stream_with_globals(&globals_1900, &[("Sheet1", 0, cell.clone())]);
        let doc_1900 = XlsDocument::parse_workbook_stream(&stream_1900).expect("parses");

        let globals_1904 = vec![xf_date, biff_rec(RT_DATEMODE, &1u16.to_le_bytes())];
        let stream_1904 = workbook_stream_with_globals(&globals_1904, &[("Sheet1", 0, cell)]);
        let doc_1904 = XlsDocument::parse_workbook_stream(&stream_1904).expect("parses");

        assert_eq!(doc_1900.sheets[0].display[0][0], "2006-09-11");
        assert_ne!(
            doc_1900.sheets[0].display[0][0], doc_1904.sheets[0].display[0][0],
            "a DATEMODE=1904 record must change the rendered date"
        );
    }

    #[test]
    fn ir_format_is_xls() {
        let ir = crate::convert_xls::xls_to_ir(&make_doc(vec![]));
        assert_eq!(ir.metadata.format, crate::format::DocumentFormat::Xls);
    }

    // ── NAME record (defined names, #251 XLS half) ─────────────────────────

    fn supbook_internal_record(ctab: u16) -> Vec<u8> {
        let mut d = Vec::new();
        d.extend_from_slice(&ctab.to_le_bytes());
        d.extend_from_slice(&0x0401u16.to_le_bytes()); // self-referencing marker
        biff_rec(RT_SUPBOOK, &d)
    }

    fn supbook_external_record() -> Vec<u8> {
        let mut d = Vec::new();
        d.extend_from_slice(&0u16.to_le_bytes()); // ctab
        d.extend_from_slice(&1u16.to_le_bytes()); // cch = 1 -> virtPath follows
        d.push(0); // 8-bit chars
        d.push(b' '); // single-char virtPath (irrelevant to internal detection)
        biff_rec(RT_SUPBOOK, &d)
    }

    fn externsheet_record(entries: &[(u16, i16, i16)]) -> Vec<u8> {
        let mut d = Vec::new();
        d.extend_from_slice(&(entries.len() as u16).to_le_bytes());
        for (sup, first, last) in entries {
            d.extend_from_slice(&sup.to_le_bytes());
            d.extend_from_slice(&first.to_le_bytes());
            d.extend_from_slice(&last.to_le_bytes());
        }
        biff_rec(RT_EXTERNSHEET, &d)
    }

    fn name_record(name: &str, itab: u16, hidden: bool, rgce: &[u8]) -> Vec<u8> {
        let mut d = Vec::new();
        let grbit: u16 = if hidden { 0x0001 } else { 0x0000 };
        d.extend_from_slice(&grbit.to_le_bytes());
        d.push(0); // chKey
        d.push(name.len() as u8); // cch
        d.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        d.extend_from_slice(&0u16.to_le_bytes()); // reserved3
        d.extend_from_slice(&itab.to_le_bytes()); // itab
        d.extend_from_slice(&[0u8; 4]); // reserved4..7
        d.push(0); // fHighByte = 0 (compressed 8-bit chars)
        d.extend_from_slice(name.as_bytes());
        d.extend_from_slice(rgce);
        biff_rec(RT_NAME, &d)
    }

    /// A built-in `NAME` record: `Name` is a 1-byte `BuiltInName` ID, not
    /// text.
    fn builtin_name_record(id: u8, itab: u16, rgce: &[u8]) -> Vec<u8> {
        let mut d = Vec::new();
        let grbit: u16 = 0x0020; // fBuiltin
        d.extend_from_slice(&grbit.to_le_bytes());
        d.push(0); // chKey
        d.push(1); // cch
        d.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        d.extend_from_slice(&0u16.to_le_bytes()); // reserved3
        d.extend_from_slice(&itab.to_le_bytes()); // itab
        d.extend_from_slice(&[0u8; 4]); // reserved4..7
        d.push(0); // fHighByte = 0
        d.push(id);
        d.extend_from_slice(rgce);
        biff_rec(RT_NAME, &d)
    }

    fn ptg_area3d(ixti: u16, rw_first: u16, rw_last: u16, col_first: u16, col_last: u16) -> Vec<u8> {
        let mut d = vec![0x1Bu8];
        d.extend_from_slice(&ixti.to_le_bytes());
        d.extend_from_slice(&rw_first.to_le_bytes());
        d.extend_from_slice(&rw_last.to_le_bytes());
        d.extend_from_slice(&col_first.to_le_bytes());
        d.extend_from_slice(&col_last.to_le_bytes());
        d
    }

    fn ptg_ref3d(ixti: u16, row: u16, col: u16) -> Vec<u8> {
        let mut d = vec![0x1Au8];
        d.extend_from_slice(&ixti.to_le_bytes());
        d.extend_from_slice(&row.to_le_bytes());
        d.extend_from_slice(&col.to_le_bytes());
        d
    }

    /// A global name over an area, and a sheet-local name over a single
    /// cell, both resolve through `SUPBOOK`/`EXTERNSHEET` to real
    /// `Sheet!$A$1`-style text and the right `local_sheet_id` (#251).
    #[test]
    fn test_name_records_resolve_area_and_cell_refs_via_externsheet() {
        let globals = vec![
            supbook_internal_record(2),
            externsheet_record(&[(0, 0, 0)]),
            name_record("MyRange", 0, false, &ptg_area3d(0, 0, 9, 0, 1)),
            name_record("MyCell", 1, true, &ptg_ref3d(0, 3, 4)),
        ];
        let stream = workbook_stream_with_globals(
            &globals,
            &[("Data", 0, label(0, 0, "x")), ("Summary", 0, label(0, 0, "y"))],
        );
        let doc = XlsDocument::parse_workbook_stream(&stream).expect("parses");

        assert_eq!(doc.defined_names.len(), 2);

        let my_range = &doc.defined_names[0];
        assert_eq!(my_range.name, "MyRange");
        assert_eq!(my_range.value, "Data!$A$1:$B$10");
        assert_eq!(my_range.local_sheet_id, None, "itab 0 is workbook-global");
        assert!(!my_range.hidden);

        let my_cell = &doc.defined_names[1];
        assert_eq!(my_cell.name, "MyCell");
        assert_eq!(my_cell.value, "Data!$E$4");
        assert_eq!(my_cell.local_sheet_id, Some(0), "itab 1 -> sheet_infos[0]");
        assert!(my_cell.hidden);
    }

    /// A name pointing through an external-workbook `SUPBOOK` (not the
    /// self-referencing one) must not be misreported as pointing at a local
    /// sheet — better an empty `value` than a wrong one.
    #[test]
    fn test_name_record_via_external_supbook_leaves_value_unresolved() {
        let globals = vec![
            supbook_external_record(),
            externsheet_record(&[(0, 0, 0)]),
            name_record("External", 0, false, &ptg_ref3d(0, 0, 0)),
        ];
        let stream = workbook_stream_with_globals(&globals, &[("Sheet1", 0, label(0, 0, "x"))]);
        let doc = XlsDocument::parse_workbook_stream(&stream).expect("parses");

        assert_eq!(doc.defined_names.len(), 1);
        assert_eq!(doc.defined_names[0].name, "External");
        assert_eq!(doc.defined_names[0].value, "");
    }

    /// A formula more complex than a single cell/area reference (here: no
    /// tokens at all) must not be guessed at.
    #[test]
    fn test_name_record_with_unsupported_formula_shape_leaves_value_empty() {
        let globals = vec![
            supbook_internal_record(1),
            externsheet_record(&[(0, 0, 0)]),
            name_record("Weird", 0, false, &[]),
        ];
        let stream = workbook_stream_with_globals(&globals, &[("Sheet1", 0, label(0, 0, "x"))]);
        let doc = XlsDocument::parse_workbook_stream(&stream).expect("parses");

        assert_eq!(doc.defined_names.len(), 1);
        assert_eq!(doc.defined_names[0].value, "");
    }

    /// The resolved names reach `to_ir()` as `DocumentIR::defined_names`.
    #[test]
    fn test_defined_names_reach_the_ir() {
        let globals = vec![
            supbook_internal_record(1),
            externsheet_record(&[(0, 0, 0)]),
            name_record("MyRange", 0, false, &ptg_area3d(0, 0, 0, 0, 0)),
        ];
        let stream = workbook_stream_with_globals(&globals, &[("Sheet1", 0, label(0, 0, "x"))]);
        let doc = XlsDocument::parse_workbook_stream(&stream).expect("parses");

        let ir = crate::convert_xls::xls_to_ir(&doc);
        assert_eq!(ir.defined_names.len(), 1);
        assert_eq!(ir.defined_names[0].name, "MyRange");
        assert_eq!(ir.defined_names[0].value, "Sheet1!$A$1:$A$1");
    }

    /// A built-in name's `Name` field is a `BuiltInName` ID byte, not text
    /// (real `.xls` corpus files decoded it as raw control characters like
    /// U+0006 before this): it now maps to the same `_xlnm.`-prefixed
    /// reserved name XLSX already surfaces for `<definedName
    /// name="_xlnm.Print_Area">`.
    #[test]
    fn test_builtin_name_decodes_to_the_xlnm_reserved_name() {
        let globals = vec![
            supbook_internal_record(1),
            externsheet_record(&[(0, 0, 0)]),
            builtin_name_record(0x06, 0, &ptg_area3d(0, 0, 9, 0, 4)), // Print_Area
        ];
        let stream = workbook_stream_with_globals(&globals, &[("Sheet1", 0, label(0, 0, "x"))]);
        let doc = XlsDocument::parse_workbook_stream(&stream).expect("parses");

        assert_eq!(doc.defined_names.len(), 1);
        assert_eq!(doc.defined_names[0].name, "_xlnm.Print_Area");
        assert_eq!(doc.defined_names[0].value, "Sheet1!$A$1:$E$10");
    }
}
