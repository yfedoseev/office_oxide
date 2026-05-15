use quick_xml::events::Event;

use crate::core::xml;

use super::cell::{Cell, CellRef, CellValue};

/// A parsed worksheet from `xl/worksheets/sheetN.xml`.
#[derive(Debug, Clone)]
pub struct Worksheet {
    /// Sheet display name.
    pub name: String,
    /// Dimension string like "A1:G50", if present.
    pub dimension: Option<String>,
    /// Data rows.
    pub rows: Vec<Row>,
    /// Merged cell ranges like "A1:C1".
    pub merged_cells: Vec<String>,
    /// Hyperlinks defined on this sheet.
    pub hyperlinks: Vec<HyperlinkInfo>,
    /// Per-sheet page geometry parsed from `<pageMargins>` + `<pageSetup>`.
    pub page_setup: Option<PageSetup>,
    /// Pictures anchored on this worksheet via `xl/drawings/drawingN.xml`.
    /// Resolved at parse time: anchor + image bytes are materialised
    /// into this `Vec` so consumers don't need to re-walk the OPC
    /// reader. Empty when the worksheet has no drawing rel.
    pub images: Vec<WorksheetPicture>,
    /// Layout-preserving text shapes anchored on this worksheet via a
    /// DrawingML drawing part. Each entry is one `<xdr:sp>` carrying a
    /// single styled run — populated by the round-trip from
    /// `to_xlsx_bytes_layout`. Empty when the worksheet has no
    /// `<xdr:sp>` shapes (the common XLSX case).
    pub text_shapes: Vec<WorksheetTextShape>,
}

/// A text shape anchored on a worksheet via a DrawingML drawing part.
/// Mirrors `xlsx::write::SheetTextShape`.
#[derive(Debug, Clone)]
pub struct WorksheetTextShape {
    /// Text content of the shape.
    pub text: String,
    /// Font face name.
    pub font_name: Option<String>,
    /// Font size in points (full-pt scale).
    pub font_size_pt: Option<f32>,
    /// Bold weight.
    pub bold: bool,
    /// Italic style.
    pub italic: bool,
    /// 6-char hex colour, when present.
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

/// A picture anchored on a worksheet via a DrawingML drawing part.
///
/// Coordinates are in EMU (914400 per inch) and absolute relative to
/// the sheet origin (top-left). When the source used a one-cell or
/// two-cell anchor we approximate the equivalent absolute origin by
/// summing the from-cell coordinates. The bytes are the raw image
/// part contents; `format` is the lowercase file extension.
#[derive(Debug, Clone)]
pub struct WorksheetPicture {
    /// Image bytes.
    pub data: Vec<u8>,
    /// Lowercase file extension (`"png"`, `"jpeg"`, ...).
    pub format: String,
    /// X anchor in EMU.
    pub x_emu: i64,
    /// Y anchor in EMU.
    pub y_emu: i64,
    /// Rendered width in EMU.
    pub cx_emu: i64,
    /// Rendered height in EMU.
    pub cy_emu: i64,
    /// Optional `<xdr:cNvPr descr=…>` accessibility text.
    pub alt_text: Option<String>,
}

/// Per-sheet page geometry (inches for margins, twips for dimensions).
///
/// Parsed from `<pageMargins>` (margins in inches per ECMA-376) and
/// `<pageSetup>` (size as `paperWidth`/`paperHeight` with a unit suffix
/// — `mm`, `cm`, `in` — or as a `paperSize` enum).  Stored in twips for
/// IR parity (1 inch = 1440 twips, 1 mm = 1440/25.4 ≈ 56.6929 twips).
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct PageSetup {
    /// Page width in twips. Zero if no page setup was seen.
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

/// A row from `<sheetData>`.
#[derive(Debug, Clone)]
pub struct Row {
    /// 1-based row number from the `r` attribute.
    pub index: u32,
    /// Cells in this row.
    pub cells: Vec<Cell>,
}

/// Hyperlink information from `<hyperlinks>`.
#[derive(Debug, Clone)]
pub struct HyperlinkInfo {
    /// Cell reference like "A1".
    pub cell_ref: String,
    /// Hyperlink destination.
    pub target: HyperlinkTarget,
    /// Optional tooltip text.
    pub tooltip: Option<String>,
}

/// Hyperlink target type.
#[derive(Debug, Clone)]
pub enum HyperlinkTarget {
    /// External URL.
    External(String),
    /// Internal sheet/cell location.
    Internal(String),
}

impl Worksheet {
    /// Parse a worksheet XML part.
    pub fn parse(
        xml_data: &[u8],
        name: String,
        rels: &crate::core::relationships::Relationships,
    ) -> crate::core::Result<Self> {
        // Use plain Reader (not NsReader) for performance — worksheet XML is always
        // in the SML namespace, so namespace resolution is unnecessary overhead.
        // This is the hot path: worksheets can have thousands of cells.
        let mut reader = xml::make_fast_reader(xml_data);
        let mut dimension = None;
        let mut rows = Vec::new();
        let mut merged_cells = Vec::new();
        let mut hyperlinks = Vec::new();
        // Page setup is collected lazily because <pageMargins> and
        // <pageSetup> arrive as separate sibling elements and either may
        // appear without the other. We materialize the IR value at the
        // end iff at least one was seen.
        let mut margins_in: Option<PageMarginsIn> = None;
        let mut page_setup_raw: Option<PageSetupRaw> = None;

        loop {
            match reader.read_event()? {
                Event::Start(ref e) => match e.local_name().as_ref() {
                    b"dimension" => {
                        dimension = xml::optional_attr_str(e, b"ref")?.map(|v| v.into_owned());
                        reader.read_to_end(e.to_end().name())?;
                    },
                    b"row" => {
                        rows.push(parse_row_fast(&mut reader, e)?);
                    },
                    b"mergeCell" => {
                        if let Some(range) = xml::optional_attr_str(e, b"ref")? {
                            merged_cells.push(range.into_owned());
                        }
                        reader.read_to_end(e.to_end().name())?;
                    },
                    b"hyperlink" => {
                        if let Some(hl) = parse_hyperlink(e, rels)? {
                            hyperlinks.push(hl);
                        }
                        reader.read_to_end(e.to_end().name())?;
                    },
                    b"pageMargins" => {
                        margins_in = parse_page_margins(e)?;
                        reader.read_to_end(e.to_end().name())?;
                    },
                    b"pageSetup" => {
                        page_setup_raw = parse_page_setup_attrs(e)?;
                        reader.read_to_end(e.to_end().name())?;
                    },
                    _ => {},
                },
                Event::Empty(ref e) => match e.local_name().as_ref() {
                    b"dimension" => {
                        dimension = xml::optional_attr_str(e, b"ref")?.map(|v| v.into_owned());
                    },
                    b"mergeCell" => {
                        if let Some(range) = xml::optional_attr_str(e, b"ref")? {
                            merged_cells.push(range.into_owned());
                        }
                    },
                    b"hyperlink" => {
                        if let Some(hl) = parse_hyperlink(e, rels)? {
                            hyperlinks.push(hl);
                        }
                    },
                    b"pageMargins" => {
                        margins_in = parse_page_margins(e)?;
                    },
                    b"pageSetup" => {
                        page_setup_raw = parse_page_setup_attrs(e)?;
                    },
                    _ => {},
                },
                Event::Eof => break,
                _ => {},
            }
        }

        let page_setup = build_page_setup(margins_in, page_setup_raw);

        Ok(Worksheet {
            name,
            dimension,
            rows,
            merged_cells,
            hyperlinks,
            page_setup,
            images: Vec::new(),
            text_shapes: Vec::new(),
        })
    }
}

/// Raw `<pageMargins>` values in inches (per ECMA-376 §18.3.1.62).
#[derive(Debug, Clone, Copy)]
struct PageMarginsIn {
    left: f64,
    right: f64,
    top: f64,
    bottom: f64,
    header: f64,
    footer: f64,
}

/// Raw `<pageSetup>` shape — physical dimensions in twips plus orientation.
#[derive(Debug, Clone, Copy, Default)]
struct PageSetupRaw {
    width_twips: u32,
    height_twips: u32,
    landscape: bool,
}

fn parse_page_margins(
    e: &quick_xml::events::BytesStart,
) -> crate::core::Result<Option<PageMarginsIn>> {
    let parse = |k: &[u8]| -> crate::core::Result<Option<f64>> {
        Ok(xml::optional_attr_str(e, k)?
            .and_then(|v| fast_float2::parse::<f64, _>(v.as_ref()).ok()))
    };
    let left = parse(b"left")?;
    let right = parse(b"right")?;
    let top = parse(b"top")?;
    let bottom = parse(b"bottom")?;
    let header = parse(b"header")?;
    let footer = parse(b"footer")?;
    if left.is_none() && right.is_none() && top.is_none() && bottom.is_none() {
        return Ok(None);
    }
    Ok(Some(PageMarginsIn {
        left: left.unwrap_or(0.7),
        right: right.unwrap_or(0.7),
        top: top.unwrap_or(0.75),
        bottom: bottom.unwrap_or(0.75),
        header: header.unwrap_or(0.3),
        footer: footer.unwrap_or(0.3),
    }))
}

/// Translate an inch / mm / cm dimension token (e.g. "210mm", "8.5in",
/// "21cm", or a bare "210" assumed mm) into twips.  Returns `None` for
/// blanks or values that fail to parse.
fn dim_to_twips(s: &str) -> Option<u32> {
    let s = s.trim();
    if s.is_empty() {
        return None;
    }
    let (num_part, factor): (&str, f64) = if let Some(rest) = s.strip_suffix("mm") {
        (rest, 1440.0 / 25.4)
    } else if let Some(rest) = s.strip_suffix("cm") {
        (rest, 1440.0 / 2.54)
    } else if let Some(rest) = s.strip_suffix("in") {
        (rest, 1440.0)
    } else {
        // Bare numeric — ECMA-376 says the default unit varies by locale;
        // mm is the safest bet for arbitrary writers (and matches what we
        // emit in `build_worksheet_xml`).
        (s, 1440.0 / 25.4)
    };
    let v: f64 = fast_float2::parse(num_part.trim()).ok()?;
    if v <= 0.0 {
        return None;
    }
    Some((v * factor).round() as u32)
}

/// Translate the OOXML `paperSize` enum into (width_twips, height_twips).
/// Covers the dimensions we're most likely to encounter in a PDF→XLSX
/// round-trip — Letter, Legal, A3, A4, A5, B4, B5, Executive, Tabloid.
/// Unknown values fall back to A4 portrait.
fn paper_size_enum_to_twips(id: u32) -> (u32, u32) {
    match id {
        1 => (12240, 15840),  // Letter 8.5 × 11"
        5 => (12240, 20160),  // Legal 8.5 × 14"
        7 => (10440, 15120),  // Executive 7.25 × 10.5"
        8 => (16838, 23811),  // A3 297 × 420 mm
        9 => (11906, 16838),  // A4 210 × 297 mm
        11 => (8392, 11906),  // A5 148 × 210 mm
        12 => (14171, 20012), // B4 250 × 353 mm
        13 => (9979, 14171),  // B5 176 × 250 mm
        3 => (15840, 24480),  // Tabloid 11 × 17"
        _ => (11906, 16838),  // Default A4
    }
}

fn parse_page_setup_attrs(
    e: &quick_xml::events::BytesStart,
) -> crate::core::Result<Option<PageSetupRaw>> {
    let pw = xml::optional_attr_str(e, b"paperWidth")?.and_then(|v| dim_to_twips(v.as_ref()));
    let ph = xml::optional_attr_str(e, b"paperHeight")?.and_then(|v| dim_to_twips(v.as_ref()));
    let paper_size = xml::optional_attr_str(e, b"paperSize")?
        .and_then(|v| atoi_simd::parse_pos::<u32, false>(v.as_bytes()).ok());
    let orientation = xml::optional_attr_str(e, b"orientation")?;
    let landscape = matches!(orientation.as_deref(), Some("landscape"));

    let (width_twips, height_twips) = match (pw, ph) {
        (Some(w), Some(h)) => (w, h),
        _ => match paper_size {
            Some(id) => paper_size_enum_to_twips(id),
            None => return Ok(None),
        },
    };

    Ok(Some(PageSetupRaw {
        width_twips,
        height_twips,
        landscape,
    }))
}

fn build_page_setup(
    margins: Option<PageMarginsIn>,
    raw: Option<PageSetupRaw>,
) -> Option<PageSetup> {
    if margins.is_none() && raw.is_none() {
        return None;
    }
    let in_to_twips = |v: f64| (v * 1440.0).round().max(0.0) as u32;
    let m = margins.unwrap_or(PageMarginsIn {
        left: 0.7,
        right: 0.7,
        top: 0.75,
        bottom: 0.75,
        header: 0.3,
        footer: 0.3,
    });
    let r = raw.unwrap_or_default();
    let ps = PageSetup {
        width_twips: r.width_twips,
        height_twips: r.height_twips,
        margin_top_twips: in_to_twips(m.top),
        margin_bottom_twips: in_to_twips(m.bottom),
        margin_left_twips: in_to_twips(m.left),
        margin_right_twips: in_to_twips(m.right),
        header_distance_twips: in_to_twips(m.header),
        footer_distance_twips: in_to_twips(m.footer),
        landscape: r.landscape,
    };
    Some(ps)
}

fn parse_hyperlink(
    e: &quick_xml::events::BytesStart,
    rels: &crate::core::relationships::Relationships,
) -> crate::core::Result<Option<HyperlinkInfo>> {
    let cell_ref = match xml::optional_attr_str(e, b"ref")? {
        Some(v) => v.into_owned(),
        None => return Ok(None),
    };
    let tooltip = xml::optional_attr_str(e, b"tooltip")?.map(|v| v.into_owned());

    // r:id → external hyperlink via relationships
    let r_id = xml::optional_attr_str(e, b"r:id")?;
    let location = xml::optional_attr_str(e, b"location")?;

    let target = if let Some(rid) = r_id {
        if let Some(rel) = rels.get_by_id(&rid) {
            HyperlinkTarget::External(rel.target.clone())
        } else {
            return Ok(None);
        }
    } else if let Some(loc) = location {
        HyperlinkTarget::Internal(loc.into_owned())
    } else {
        return Ok(None);
    };

    Ok(Some(HyperlinkInfo {
        cell_ref,
        target,
        tooltip,
    }))
}

/// Fast row parser using plain Reader (no namespace resolution).
fn parse_row_fast(
    reader: &mut quick_xml::Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> crate::core::Result<Row> {
    let index: u32 = xml::optional_attr_str(start, b"r")?
        .and_then(|v| atoi_simd::parse_pos::<u32, false>(v.as_bytes()).ok())
        .unwrap_or(1);
    let mut cells = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.local_name().as_ref() == b"c" {
                    cells.push(parse_cell_fast(reader, e)?);
                } else {
                    reader.read_to_end(e.to_end().name())?;
                }
            },
            Event::Empty(ref e) if e.local_name().as_ref() == b"c" => {
                cells.push(parse_empty_cell(e)?);
            },
            Event::End(ref e) if e.local_name().as_ref() == b"row" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(Row { index, cells })
}

fn parse_empty_cell(e: &quick_xml::events::BytesStart) -> crate::core::Result<Cell> {
    let ref_str = xml::optional_attr_str(e, b"r")?
        .map(|v| v.into_owned())
        .unwrap_or_default();
    let reference = CellRef::parse(&ref_str).unwrap_or(CellRef { col: 0, row: 0 });
    let style_index = xml::optional_attr_str(e, b"s")?
        .and_then(|v| atoi_simd::parse_pos::<u32, false>(v.as_bytes()).ok());

    Ok(Cell {
        reference,
        value: CellValue::Empty,
        style_index,
        formula: None,
    })
}

/// Fast cell parser using plain Reader (no namespace resolution).
fn parse_cell_fast(
    reader: &mut quick_xml::Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> crate::core::Result<Cell> {
    let ref_str = xml::optional_attr_str(start, b"r")?
        .map(|v| v.into_owned())
        .unwrap_or_default();
    let reference = CellRef::parse(&ref_str).unwrap_or(CellRef { col: 0, row: 0 });

    let cell_type = xml::optional_attr_str(start, b"t")?.map(|v| v.into_owned());
    let style_index = xml::optional_attr_str(start, b"s")?
        .and_then(|v| atoi_simd::parse_pos::<u32, false>(v.as_bytes()).ok());

    let mut raw_value: Option<String> = None;
    let mut formula: Option<String> = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"v" => {
                    raw_value = Some(read_text_fast(reader)?);
                },
                b"f" => {
                    formula = Some(read_text_fast(reader)?);
                },
                b"is" => {
                    raw_value = Some(parse_inline_string_fast(reader)?);
                },
                _ => {
                    reader.read_to_end(e.to_end().name())?;
                },
            },
            Event::Empty(ref e) if e.local_name().as_ref() == b"f" => {
                formula = None;
            },
            Event::End(ref e) if e.local_name().as_ref() == b"c" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    let value = match cell_type.as_deref() {
        Some("s") => {
            match raw_value
                .as_deref()
                .and_then(|v| atoi_simd::parse_pos::<u32, false>(v.as_bytes()).ok())
            {
                Some(idx) => CellValue::SharedString(idx),
                None => CellValue::Empty,
            }
        },
        Some("str") | Some("inlineStr") => match raw_value {
            Some(s) => CellValue::String(s),
            None => CellValue::Empty,
        },
        Some("b") => match raw_value.as_deref() {
            Some("1") | Some("true") => CellValue::Boolean(true),
            Some("0") | Some("false") => CellValue::Boolean(false),
            _ => CellValue::Empty,
        },
        Some("e") => match raw_value {
            Some(s) => CellValue::Error(s),
            None => CellValue::Error(String::new()),
        },
        _ => match raw_value {
            Some(s) => match fast_float2::parse::<f64, _>(&s) {
                Ok(n) => CellValue::Number(n),
                Err(_) => CellValue::String(s),
            },
            None => CellValue::Empty,
        },
    };

    Ok(Cell {
        reference,
        value,
        style_index,
        formula,
    })
}

/// Read text content of the current element using fast Reader.
fn read_text_fast(reader: &mut quick_xml::Reader<&[u8]>) -> crate::core::Result<String> {
    xml::read_text_content_fast(reader)
}

/// Fast inline string parser: `<is><t>text</t></is>` or `<is><r>...<t>text</t>...</r></is>`.
fn parse_inline_string_fast(reader: &mut quick_xml::Reader<&[u8]>) -> crate::core::Result<String> {
    let mut text = String::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.local_name().as_ref() == b"t" {
                    text.push_str(&read_text_fast(reader)?);
                } else {
                    reader.read_to_end(e.to_end().name())?;
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"is" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(text)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::core::relationships::Relationships;

    fn empty_rels() -> Relationships {
        Relationships::empty()
    }

    #[test]
    fn parse_simple_worksheet() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1"><v>42</v></c>
    </row>
    <row r="2">
      <c r="A2" t="b"><v>1</v></c>
      <c r="B2" t="e"><v>#DIV/0!</v></c>
    </row>
  </sheetData>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "Sheet1".to_string(), &empty_rels()).unwrap();
        assert_eq!(ws.name, "Sheet1");
        assert_eq!(ws.dimension.as_deref(), Some("A1:B2"));
        assert_eq!(ws.rows.len(), 2);

        // Row 1
        assert_eq!(ws.rows[0].cells.len(), 2);
        assert!(matches!(ws.rows[0].cells[0].value, CellValue::SharedString(0)));
        assert!(matches!(ws.rows[0].cells[1].value, CellValue::Number(n) if n == 42.0));

        // Row 2
        assert!(matches!(ws.rows[1].cells[0].value, CellValue::Boolean(true)));
        assert!(matches!(&ws.rows[1].cells[1].value, CellValue::Error(e) if e == "#DIV/0!"));
    }

    #[test]
    fn parse_worksheet_with_formula() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>10</v></c>
      <c r="B1"><f>A1*2</f><v>20</v></c>
    </row>
  </sheetData>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "Sheet1".to_string(), &empty_rels()).unwrap();
        let cell = &ws.rows[0].cells[1];
        assert_eq!(cell.formula.as_deref(), Some("A1*2"));
        assert!(matches!(cell.value, CellValue::Number(n) if n == 20.0));
    }

    #[test]
    fn parse_worksheet_page_setup() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>
  <pageSetup paperWidth="215.90mm" paperHeight="279.40mm" orientation="portrait"/>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        let ps = ws.page_setup.expect("page_setup parsed");
        // 215.9mm ≈ 8.5", 279.4mm ≈ 11", both in twips
        assert!((ps.width_twips as i32 - 12240).abs() <= 1, "width {:?}", ps.width_twips);
        assert!((ps.height_twips as i32 - 15840).abs() <= 1, "height {:?}", ps.height_twips);
        // 0.5" margin = 720 twips
        assert_eq!(ps.margin_top_twips, 720);
        assert_eq!(ps.margin_left_twips, 720);
        assert!(!ps.landscape);
    }

    #[test]
    fn parse_worksheet_page_setup_paper_enum() {
        // paperSize=9 = A4 → 11906x16838 twips.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <pageSetup paperSize="9" orientation="landscape"/>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        let ps = ws.page_setup.expect("page_setup parsed");
        assert_eq!(ps.width_twips, 11906);
        assert_eq!(ps.height_twips, 16838);
        assert!(ps.landscape);
    }

    #[test]
    fn parse_worksheet_merged_cells() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="A1:C1"/>
  </mergeCells>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "Sheet1".to_string(), &empty_rels()).unwrap();
        assert_eq!(ws.merged_cells, vec!["A1:C1"]);
    }

    // ── dim_to_twips ─────────────────────────────────────────────────────

    #[test]
    fn dim_to_twips_inches() {
        // 1 inch = 1440 twips.
        assert_eq!(dim_to_twips("1in"), Some(1440));
        assert_eq!(dim_to_twips("8.5in"), Some(12240));
    }

    #[test]
    fn dim_to_twips_millimeters() {
        // 210mm = 11906 twips (A4 width); allow ±1 for rounding.
        let twips = dim_to_twips("210mm").unwrap();
        assert!((twips as i32 - 11906).abs() <= 1, "got {twips}");
    }

    #[test]
    fn dim_to_twips_centimeters() {
        // 1cm = 1440/2.54 ≈ 567 twips.
        let twips = dim_to_twips("1cm").unwrap();
        assert!((twips as i32 - 567).abs() <= 1, "got {twips}");
    }

    #[test]
    fn dim_to_twips_bare_number_assumed_mm() {
        // Bare numeric defaults to mm.
        let a = dim_to_twips("210mm").unwrap();
        let b = dim_to_twips("210").unwrap();
        assert_eq!(a, b);
    }

    #[test]
    fn dim_to_twips_empty_and_zero() {
        assert_eq!(dim_to_twips(""), None);
        assert_eq!(dim_to_twips("   "), None);
        // Zero / negative dimensions are nonsensical: rejected.
        assert_eq!(dim_to_twips("0mm"), None);
        assert_eq!(dim_to_twips("-5in"), None);
    }

    #[test]
    fn dim_to_twips_invalid_string() {
        assert_eq!(dim_to_twips("garbage"), None);
        assert_eq!(dim_to_twips("abcmm"), None);
    }

    // ── paper_size_enum_to_twips ────────────────────────────────────────

    #[test]
    fn paper_size_letter() {
        assert_eq!(paper_size_enum_to_twips(1), (12240, 15840));
    }

    #[test]
    fn paper_size_legal() {
        assert_eq!(paper_size_enum_to_twips(5), (12240, 20160));
    }

    #[test]
    fn paper_size_a4() {
        assert_eq!(paper_size_enum_to_twips(9), (11906, 16838));
    }

    #[test]
    fn paper_size_unknown_falls_back_to_a4() {
        assert_eq!(paper_size_enum_to_twips(9999), (11906, 16838));
    }

    // ── build_page_setup ────────────────────────────────────────────────

    #[test]
    fn build_page_setup_returns_none_when_both_missing() {
        assert!(build_page_setup(None, None).is_none());
    }

    #[test]
    fn build_page_setup_margins_only_zeroes_dimensions() {
        // <pageMargins> without <pageSetup> → dimensions left at 0 so
        // a downstream consumer falls back to its default page size.
        let margins = Some(PageMarginsIn {
            left: 1.0,
            right: 1.0,
            top: 1.0,
            bottom: 1.0,
            header: 0.5,
            footer: 0.5,
        });
        let ps = build_page_setup(margins, None).unwrap();
        assert_eq!(ps.width_twips, 0);
        assert_eq!(ps.height_twips, 0);
        // 1 inch margins = 1440 twips.
        assert_eq!(ps.margin_top_twips, 1440);
        assert_eq!(ps.margin_left_twips, 1440);
        assert_eq!(ps.header_distance_twips, 720); // 0.5 in
    }

    #[test]
    fn build_page_setup_dimensions_only_uses_default_margins() {
        // <pageSetup> alone uses ECMA-376 default 0.7/0.7/0.75/0.75 inch margins.
        let raw = Some(PageSetupRaw {
            width_twips: 12240,
            height_twips: 15840,
            landscape: false,
        });
        let ps = build_page_setup(None, raw).unwrap();
        assert_eq!(ps.width_twips, 12240);
        assert_eq!(ps.height_twips, 15840);
        // 0.7in = 1008 twips.
        assert_eq!(ps.margin_left_twips, 1008);
        // 0.75in = 1080 twips.
        assert_eq!(ps.margin_top_twips, 1080);
    }

    #[test]
    fn build_page_setup_combines_both() {
        let margins = Some(PageMarginsIn {
            left: 0.5,
            right: 0.5,
            top: 0.5,
            bottom: 0.5,
            header: 0.3,
            footer: 0.3,
        });
        let raw = Some(PageSetupRaw {
            width_twips: 11906,
            height_twips: 16838,
            landscape: true,
        });
        let ps = build_page_setup(margins, raw).unwrap();
        assert_eq!(ps.width_twips, 11906);
        assert!(ps.landscape);
        assert_eq!(ps.margin_left_twips, 720); // 0.5in
    }

    #[test]
    fn parse_worksheet_landscape_with_paper_enum() {
        // Verifies that landscape attribute survives the parse_page_setup_attrs path.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <pageSetup paperSize="1" orientation="landscape"/>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        let ps = ws.page_setup.expect("page_setup");
        assert_eq!(ps.width_twips, 12240); // Letter
        assert!(ps.landscape);
    }

    #[test]
    fn parse_worksheet_default_when_no_setup() {
        // No <pageMargins> or <pageSetup> → no page_setup at all.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        assert!(ws.page_setup.is_none());
    }
}
