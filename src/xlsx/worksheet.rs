use quick_xml::events::Event;

use crate::core::xml;

use super::cell::{Cell, CellRef, CellValue};
use super::shared_formula::SharedFormulas;

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
    /// Cell comments from the sheet's `xl/comments*.xml` part, in document
    /// order. Empty when the sheet has no comments.
    pub comments: Vec<SheetComment>,
    /// Layout-preserving text shapes anchored on this worksheet via a
    /// DrawingML drawing part. Each entry is one `<xdr:sp>` carrying a
    /// single styled run — populated by the round-trip from
    /// `to_xlsx_bytes_layout`. Empty when the worksheet has no
    /// `<xdr:sp>` shapes (the common XLSX case).
    pub text_shapes: Vec<WorksheetTextShape>,
    /// Conditional formatting rules from `<conditionalFormatting>`/
    /// `<cfRule>`. Empty when the worksheet defines none (issue #252).
    pub conditional_formats: Vec<crate::ir::ConditionalFormat>,
    /// Data validation rules from `<dataValidations>`/`<dataValidation>`.
    /// Empty when the worksheet defines none (issue #275).
    pub data_validations: Vec<crate::ir::DataValidation>,
}

/// One cell comment from `xl/comments*.xml`.
#[derive(Debug, Clone)]
pub struct SheetComment {
    /// Cell reference the comment is attached to, e.g. `"B2"`.
    pub cell_ref: String,
    /// Comment author, resolved through the part's `<authors>` list.
    pub author: Option<String>,
    /// Comment body text.
    pub text: String,
}

/// Parse an `xl/comments*.xml` part.
///
/// Comments are real document content — review notes, data provenance,
/// caveats on a figure — and were never read at all, so they reached no
/// consumer.
pub fn parse_comments(xml_data: &[u8]) -> crate::core::Result<Vec<SheetComment>> {
    use quick_xml::events::Event;
    let mut reader = xml::make_fast_reader(xml_data);
    let mut authors: Vec<String> = Vec::new();
    let mut out: Vec<SheetComment> = Vec::new();
    let mut in_authors = false;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"authors" => in_authors = true,
                b"author" if in_authors => {
                    authors.push(xml::read_text_content_fast(&mut reader)?);
                },
                b"comment" => {
                    let cell_ref = xml::optional_attr_str(e, b"ref")?
                        .map(|v| v.into_owned())
                        .unwrap_or_default();
                    let author_id = xml::optional_attr_str(e, b"authorId")?
                        .and_then(|v| v.parse::<usize>().ok());
                    let text = xml::read_text_content_fast(&mut reader)?;
                    let text = text.trim().to_string();
                    if !text.is_empty() {
                        out.push(SheetComment {
                            cell_ref,
                            author: author_id.and_then(|i| authors.get(i).cloned()),
                            text,
                        });
                    }
                },
                _ => {},
            },
            Event::End(ref e) if e.local_name().as_ref() == b"authors" => in_authors = false,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(out)
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
        let mut shared = SharedFormulas::default();
        let mut conditional_formats = Vec::new();
        let mut data_validations = Vec::new();

        loop {
            match reader.read_event()? {
                Event::Start(ref e) => match e.local_name().as_ref() {
                    b"dimension" => {
                        dimension = xml::optional_attr_str(e, b"ref")?.map(|v| v.into_owned());
                        reader.read_to_end(e.to_end().name())?;
                    },
                    b"row" => {
                        // A <row> without r= is implicitly the one after the
                        // previous row (ECMA-376 18.3.1.73); some writers omit
                        // it throughout the sheet.
                        let implied = rows.last().map_or(1, |r: &Row| r.index + 1);
                        rows.push(parse_row_fast(&mut reader, e, implied, &mut shared)?);
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
                    b"conditionalFormatting" => {
                        conditional_formats.extend(parse_conditional_formatting(&mut reader, e)?);
                    },
                    b"dataValidations" => {
                        data_validations.extend(parse_data_validations(&mut reader)?);
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

        // A shared group's master usually precedes its followers, but not
        // always — calamine ships a fixture with the records reversed — so
        // followers that arrived first are filled in now.
        if shared.has_pending() {
            for (at, formula) in shared.resolve_pending() {
                if let Some(cell) = rows
                    .iter_mut()
                    .find(|r| r.index == at.row + 1)
                    .and_then(|r| r.cells.iter_mut().find(|c| c.reference == at))
                {
                    cell.formula = Some(formula);
                }
            }
        }

        Ok(Worksheet {
            comments: Vec::new(),
            name,
            dimension,
            rows,
            merged_cells,
            hyperlinks,
            page_setup,
            images: Vec::new(),
            text_shapes: Vec::new(),
            conditional_formats,
            data_validations,
        })
    }
}

/// Parse one `<conditionalFormatting sqref="...">` block: its `sqref`
/// (which range(s) it applies to) and each `<cfRule>` child inside it.
///
/// `<cfRule type="cellIs" operator="greaterThan"><formula>100</formula></cfRule>`
/// is the common shape; colour-scale/data-bar/icon-set rules instead carry
/// a `<colorScale>`/`<dataBar>`/`<iconSet>` child with no `<formula>` at
/// all, which is fine — `formulas` is just empty for those (issue #252).
fn parse_conditional_formatting(
    reader: &mut quick_xml::Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> crate::core::Result<Vec<crate::ir::ConditionalFormat>> {
    use quick_xml::events::Event;

    let sqref = xml::optional_attr_str(start, b"sqref")?.map(|v| v.into_owned()).unwrap_or_default();
    let mut out = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == b"cfRule" => {
                let rule_type =
                    xml::optional_attr_str(e, b"type")?.map(|v| v.into_owned()).unwrap_or_default();
                let operator = xml::optional_attr_str(e, b"operator")?.map(|v| v.into_owned());
                let formulas = read_cf_rule_formulas(reader)?;
                out.push(crate::ir::ConditionalFormat {
                    range: sqref.clone(),
                    rule_type,
                    operator,
                    formulas,
                });
            },
            Event::Empty(ref e) if e.local_name().as_ref() == b"cfRule" => {
                // A rule with no children at all (no formula, no colour
                // scale/data bar/icon set) — rare, but structurally valid.
                let rule_type =
                    xml::optional_attr_str(e, b"type")?.map(|v| v.into_owned()).unwrap_or_default();
                let operator = xml::optional_attr_str(e, b"operator")?.map(|v| v.into_owned());
                out.push(crate::ir::ConditionalFormat {
                    range: sqref.clone(),
                    rule_type,
                    operator,
                    formulas: Vec::new(),
                });
            },
            Event::End(ref e) if e.local_name().as_ref() == b"conditionalFormatting" => break,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(out)
}

/// Read every `<formula>...</formula>` child of the `<cfRule>` whose Start
/// event was just consumed, up through its matching `</cfRule>`. A
/// `between`/`notBetween` operator carries two `<formula>` children (the
/// low and high bounds); most others carry one; colour-scale/data-bar/
/// icon-set rules carry none.
fn read_cf_rule_formulas(reader: &mut quick_xml::Reader<&[u8]>) -> crate::core::Result<Vec<String>> {
    use quick_xml::events::Event;
    let mut formulas = Vec::new();
    let mut depth = 1u32;
    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                depth += 1;
                if e.local_name().as_ref() == b"formula" {
                    formulas.push(xml::read_text_content_fast(reader)?);
                    depth -= 1; // read_text_content_fast already consumed </formula>
                }
            },
            Event::Empty(_) => {},
            Event::End(_) => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(formulas)
}

/// Read the contents of `<dataValidations>` (already positioned just past
/// its Start event): every `<dataValidation sqref="..." type="..."
/// operator="...">` child, up through the matching `</dataValidations>`.
fn parse_data_validations(
    reader: &mut quick_xml::Reader<&[u8]>,
) -> crate::core::Result<Vec<crate::ir::DataValidation>> {
    use quick_xml::events::Event;
    let mut out = Vec::new();
    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == b"dataValidation" => {
                out.push(parse_data_validation_with_body(reader, e)?);
            },
            Event::Empty(ref e) if e.local_name().as_ref() == b"dataValidation" => {
                out.push(data_validation_from_attrs(e)?);
            },
            Event::End(ref e) if e.local_name().as_ref() == b"dataValidations" => break,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(out)
}

/// The attributes shared by both the Start and Empty forms of
/// `<dataValidation>`: `sqref`, `type` (default `"none"`), `operator`
/// (default `"between"`, only meaningful for types that use one — same
/// `None`-for-list/custom/none convention as the XLS reader), and
/// `allowBlank`.
fn data_validation_from_attrs(
    e: &quick_xml::events::BytesStart,
) -> crate::core::Result<crate::ir::DataValidation> {
    let range = xml::optional_attr_str(e, b"sqref")?.map(|v| v.into_owned()).unwrap_or_default();
    let validation_type =
        xml::optional_attr_str(e, b"type")?.map(|v| v.into_owned()).unwrap_or_else(|| "none".to_string());
    let operator = if matches!(validation_type.as_str(), "list" | "custom" | "none") {
        None
    } else {
        Some(
            xml::optional_attr_str(e, b"operator")?
                .map(|v| v.into_owned())
                .unwrap_or_else(|| "between".to_string()),
        )
    };
    let allow_blank = xml::optional_attr_str(e, b"allowBlank")?.as_deref() == Some("1");
    Ok(crate::ir::DataValidation { range, validation_type, operator, formula1: None, formula2: None, allow_blank })
}

/// The Start form of `<dataValidation>` additionally carries `<formula1>`/
/// `<formula2>` children (a comparison value, an explicit list source
/// like `"Yes,No,Maybe"`, or a cell-range/formula reference) up through
/// the matching `</dataValidation>`.
fn parse_data_validation_with_body(
    reader: &mut quick_xml::Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> crate::core::Result<crate::ir::DataValidation> {
    use quick_xml::events::Event;
    let mut dv = data_validation_from_attrs(start)?;
    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == b"formula1" => {
                dv.formula1 = Some(xml::read_text_content_fast(reader)?);
            },
            Event::Start(ref e) if e.local_name().as_ref() == b"formula2" => {
                dv.formula2 = Some(xml::read_text_content_fast(reader)?);
            },
            Event::End(ref e) if e.local_name().as_ref() == b"dataValidation" => break,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(dv)
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

impl PageMarginsIn {
    /// ECMA-376 default margins (inches). Single source of truth used by
    /// both `parse_page_margins` (when an attribute is absent) and
    /// `build_page_setup` (when no `<pageMargins>` element was present).
    const DEFAULTS: PageMarginsIn = PageMarginsIn {
        left: 0.7,
        right: 0.7,
        top: 0.75,
        bottom: 0.75,
        header: 0.3,
        footer: 0.3,
    };
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
    let d = PageMarginsIn::DEFAULTS;
    Ok(Some(PageMarginsIn {
        left: left.unwrap_or(d.left),
        right: right.unwrap_or(d.right),
        top: top.unwrap_or(d.top),
        bottom: bottom.unwrap_or(d.bottom),
        header: header.unwrap_or(d.header),
        footer: footer.unwrap_or(d.footer),
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
            Some(id) => {
                // `paperSize` names a *portrait* stock; `orientation` then
                // rotates it. Reporting the portrait width for a landscape
                // sheet handed the consumer a page narrower than the one
                // Excel prints, which is a confidently wrong measurement
                // rather than a missing one.
                let (w, h) = paper_size_enum_to_twips(id);
                if landscape { (h, w) } else { (w, h) }
            },
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
    let m = margins.unwrap_or(PageMarginsIn::DEFAULTS);
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
    implied_index: u32,
    shared: &mut SharedFormulas,
) -> crate::core::Result<Row> {
    let index: u32 = xml::optional_attr_str(start, b"r")?
        .and_then(|v| atoi_simd::parse_pos::<u32, false>(v.as_bytes()).ok())
        .unwrap_or(implied_index);
    let mut cells = Vec::new();
    // Likewise a <c> without r= sits in the column after its predecessor.
    // Defaulting these to column 0 collapses the whole row onto one cell.
    let mut next_col: u32 = 0;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.local_name().as_ref() == b"c" {
                    let cell = parse_cell_fast(reader, e, index, next_col, shared)?;
                    next_col = cell.reference.col.saturating_add(1);
                    cells.push(cell);
                } else {
                    reader.read_to_end(e.to_end().name())?;
                }
            },
            Event::Empty(ref e) if e.local_name().as_ref() == b"c" => {
                let cell = parse_empty_cell(e, index, next_col)?;
                next_col = cell.reference.col.saturating_add(1);
                cells.push(cell);
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

fn parse_empty_cell(
    e: &quick_xml::events::BytesStart,
    row: u32,
    implied_col: u32,
) -> crate::core::Result<Cell> {
    let ref_str = xml::optional_attr_str(e, b"r")?
        .map(|v| v.into_owned())
        .unwrap_or_default();
    let reference = CellRef::parse(&ref_str).unwrap_or(CellRef {
        col: implied_col,
        row,
    });
    let style_index = xml::optional_attr_str(e, b"s")?
        .and_then(|v| atoi_simd::parse_pos::<u32, false>(v.as_bytes()).ok());

    Ok(Cell {
        reference,
        value: CellValue::Empty,
        style_index,
        formula: None,
    })
}

/// Read the `si` group index of an `<f>` element, but only when it really
/// is a shared formula (`t="shared"`). Array and dataTable formulas also
/// carry a `ref`, and a plain formula carries neither.
fn shared_si(e: &quick_xml::events::BytesStart) -> crate::core::Result<Option<u32>> {
    if xml::optional_attr_str(e, b"t")?.as_deref() != Some("shared") {
        return Ok(None);
    }
    Ok(xml::optional_attr_str(e, b"si")?
        .and_then(|v| atoi_simd::parse_pos::<u32, false>(v.as_bytes()).ok()))
}

/// Fast cell parser using plain Reader (no namespace resolution).
fn parse_cell_fast(
    reader: &mut quick_xml::Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
    row: u32,
    implied_col: u32,
    shared: &mut SharedFormulas,
) -> crate::core::Result<Cell> {
    let ref_str = xml::optional_attr_str(start, b"r")?
        .map(|v| v.into_owned())
        .unwrap_or_default();
    let reference = CellRef::parse(&ref_str).unwrap_or(CellRef {
        col: implied_col,
        row,
    });

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
                    let si = shared_si(e)?;
                    let text = read_text_fast(reader)?;
                    // A shared group's master carries the text once, here.
                    if let Some(si) = si {
                        shared.add_master(si, reference.clone(), text.clone());
                    }
                    formula = Some(text);
                },
                b"is" => {
                    raw_value = Some(parse_inline_string_fast(reader)?);
                },
                _ => {
                    reader.read_to_end(e.to_end().name())?;
                },
            },
            Event::Empty(ref e) if e.local_name().as_ref() == b"f" => {
                // A bare `<f t="shared" si="N"/>` is a follower: no text of
                // its own, but its formula is the group master's translated
                // by the row/column offset. Discarding it made a formula
                // cell indistinguishable from one with no formula (#278).
                formula = match shared_si(e)? {
                    Some(si) => shared.follower(si, &reference),
                    None => None,
                };
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
///
/// Runs with the sheet reader's text trimming turned off. The sheet reader
/// trims — right for `<v>`, wrong here: quick-xml reports every entity
/// reference as its own event, splitting the literal text around it into
/// separate `Event::Text` fragments, and trimming each fragment
/// independently eats the space that sat at the entity boundary, so
/// `AT&amp;T &lt;tag&gt;` read back as `AT&T<tag>` (#269). This matches the
/// non-trimming reader `shared_strings.rs` already uses for the same
/// content model.
fn parse_inline_string_fast(reader: &mut quick_xml::Reader<&[u8]>) -> crate::core::Result<String> {
    let trim_start = reader.config().trim_text_start;
    let trim_end = reader.config().trim_text_end;
    reader.config_mut().trim_text(false);

    let result = read_inline_string_body(reader);

    reader.config_mut().trim_text_start = trim_start;
    reader.config_mut().trim_text_end = trim_end;
    result
}

fn read_inline_string_body(reader: &mut quick_xml::Reader<&[u8]>) -> crate::core::Result<String> {
    let mut text = String::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"t" => text.push_str(&read_text_fast(reader)?),
                // A rich inline string wraps each run in `<r>`. Descend
                // into it — skipping the element wholesale discarded the
                // `<t>` inside and with it the cell's entire text.
                b"r" => {},
                _ => {
                    reader.read_to_end(e.to_end().name())?;
                },
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

    /// #278 — a shared-formula group carries its text once, on the master;
    /// every follower is a bare `<f t="shared" si="N"/>` that used to set
    /// `formula = None`, indistinguishable from a cell with no formula.
    #[test]
    fn test_shared_formula_followers_reconstruct_their_formula() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2"><c r="B2"><f t="shared" ref="B2:B4" si="0">A2</f><v>1</v></c></row>
    <row r="3"><c r="B3"><f t="shared" si="0"/><v>2</v></c></row>
    <row r="4"><c r="B4"><f t="shared" si="0"/><v>3</v></c></row>
  </sheetData>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        assert_eq!(ws.rows[0].cells[0].formula.as_deref(), Some("A2"));
        assert_eq!(ws.rows[1].cells[0].formula.as_deref(), Some("A3"));
        assert_eq!(ws.rows[2].cells[0].formula.as_deref(), Some("A4"));
    }

    /// Same, with the master written after its followers — the record order
    /// calamine's `..._reversed.xlsx` fixture uses.
    #[test]
    fn test_shared_formula_master_after_followers_still_resolves() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="3"><c r="B3"><f t="shared" si="0"/><v>2</v></c></row>
    <row r="2"><c r="B2"><f t="shared" ref="B2:B3" si="0">A2*2</f><v>1</v></c></row>
  </sheetData>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        assert_eq!(ws.rows[1].cells[0].formula.as_deref(), Some("A2*2"));
        assert_eq!(ws.rows[0].cells[0].formula.as_deref(), Some("A3*2"));
    }

    /// A plain `<f/>` with no `t="shared"` still means "no formula text",
    /// and a follower whose group never had a master stays `None` rather
    /// than inventing one.
    #[test]
    fn test_bare_formula_element_without_shared_group_stays_none() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f/><v>1</v></c>
      <c r="B1"><f t="shared" si="9"/><v>2</v></c>
    </row>
  </sheetData>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        assert_eq!(ws.rows[0].cells[0].formula, None);
        assert_eq!(ws.rows[0].cells[1].formula, None);
    }

    /// #269 — the sheet reader trims text, and quick-xml emits every entity
    /// reference as its own event, so the literal text around `&amp;`/`&lt;`
    /// arrived as separate fragments that were each trimmed independently.
    /// Every space touching an escaped character was deleted.
    #[test]
    fn test_inline_string_keeps_whitespace_around_escaped_characters() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>AT&amp;T &lt;tag&gt; &quot;quoted&quot; &apos;apostrophe&apos; 5 &lt; 10 and 10 &gt; 5</t></is></c>
    </row>
  </sheetData>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        assert!(
            matches!(
                &ws.rows[0].cells[0].value,
                CellValue::String(s)
                    if s == "AT&T <tag> \"quoted\" 'apostrophe' 5 < 10 and 10 > 5"
            ),
            "got {:?}",
            ws.rows[0].cells[0].value
        );
    }

    /// The same function skipped `<r>` wholesale, so a *rich* inline string
    /// lost its text entirely rather than just its spacing.
    #[test]
    fn test_rich_inline_string_runs_are_not_skipped() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is>
        <r><rPr><b/></rPr><t>Hello </t></r>
        <r><t>world</t></r>
      </is></c>
    </row>
  </sheetData>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        assert!(
            matches!(&ws.rows[0].cells[0].value, CellValue::String(s) if s == "Hello world"),
            "got {:?}",
            ws.rows[0].cells[0].value
        );
    }

    /// issue #252 — a `cellIs`/`greaterThan` rule's sqref, type, operator,
    /// and single comparison formula must all reach the IR.
    #[test]
    fn test_conditional_formatting_cell_is_rule_reaches_the_ir() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
  <conditionalFormatting sqref="A1:A10">
    <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
      <formula>100</formula>
    </cfRule>
  </conditionalFormatting>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        assert_eq!(ws.conditional_formats.len(), 1);
        let cf = &ws.conditional_formats[0];
        assert_eq!(cf.range, "A1:A10");
        assert_eq!(cf.rule_type, "cellIs");
        assert_eq!(cf.operator.as_deref(), Some("greaterThan"));
        assert_eq!(cf.formulas, vec!["100".to_string()]);
    }

    /// A `between` rule carries two `<formula>` children — both must be
    /// captured as separate entries, not concatenated together.
    #[test]
    fn test_conditional_formatting_between_rule_keeps_both_formulas_separate() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
  <conditionalFormatting sqref="B1:B5">
    <cfRule type="cellIs" operator="between" priority="1">
      <formula>10</formula>
      <formula>20</formula>
    </cfRule>
  </conditionalFormatting>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        assert_eq!(ws.conditional_formats[0].formulas, vec!["10".to_string(), "20".to_string()]);
    }

    /// A colour-scale rule has no `<formula>` children at all — must not
    /// error or swallow the sheet, and correctly reports an empty
    /// `formulas` list while still recording the rule's type and range.
    #[test]
    fn test_conditional_formatting_color_scale_has_no_formulas() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
  <conditionalFormatting sqref="C1:C20">
    <cfRule type="colorScale" priority="1">
      <colorScale>
        <cfvo type="min"/>
        <cfvo type="max"/>
        <color rgb="FFFF0000"/>
        <color rgb="FF00FF00"/>
      </colorScale>
    </cfRule>
  </conditionalFormatting>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        assert_eq!(ws.conditional_formats.len(), 1);
        assert_eq!(ws.conditional_formats[0].rule_type, "colorScale");
        assert!(ws.conditional_formats[0].formulas.is_empty());
        assert!(ws.conditional_formats[0].operator.is_none());
    }

    /// Multiple `<cfRule>`s under the same `<conditionalFormatting>` share
    /// its `sqref` — each must still produce its own IR entry.
    #[test]
    fn test_conditional_formatting_multiple_rules_share_the_same_range() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
  <conditionalFormatting sqref="D1:D10">
    <cfRule type="cellIs" operator="lessThan" priority="1"><formula>0</formula></cfRule>
    <cfRule type="cellIs" operator="greaterThan" priority="2"><formula>100</formula></cfRule>
  </conditionalFormatting>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        assert_eq!(ws.conditional_formats.len(), 2);
        assert!(ws.conditional_formats.iter().all(|cf| cf.range == "D1:D10"));
        assert_eq!(ws.conditional_formats[0].operator.as_deref(), Some("lessThan"));
        assert_eq!(ws.conditional_formats[1].operator.as_deref(), Some("greaterThan"));
    }

    /// A sheet with no `<conditionalFormatting>` at all must produce an
    /// empty list, not an error.
    #[test]
    fn test_no_conditional_formatting_is_an_empty_list() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        assert!(ws.conditional_formats.is_empty());
    }

    /// issue #275 — a `whole`/`between` rule's sqref, type, operator, both
    /// comparison formulas, and `allowBlank` must all reach the IR.
    #[test]
    fn test_data_validation_whole_between_rule_reaches_the_ir() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
  <dataValidations count="1">
    <dataValidation type="whole" operator="between" allowBlank="1" sqref="A1:A10">
      <formula1>1</formula1>
      <formula2>10</formula2>
    </dataValidation>
  </dataValidations>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        assert_eq!(ws.data_validations.len(), 1);
        let dv = &ws.data_validations[0];
        assert_eq!(dv.range, "A1:A10");
        assert_eq!(dv.validation_type, "whole");
        assert_eq!(dv.operator.as_deref(), Some("between"));
        assert_eq!(dv.formula1.as_deref(), Some("1"));
        assert_eq!(dv.formula2.as_deref(), Some("10"));
        assert!(dv.allow_blank);
    }

    /// A `list` rule's explicit inline source (`"Yes,No,Maybe"`) must
    /// reach `formula1` verbatim, and `list` carries no operator.
    #[test]
    fn test_data_validation_list_rule_has_no_operator() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="B1"><v>1</v></c></row></sheetData>
  <dataValidations count="1">
    <dataValidation type="list" allowBlank="0" sqref="B1:B5">
      <formula1>"Yes,No,Maybe"</formula1>
    </dataValidation>
  </dataValidations>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        assert_eq!(ws.data_validations.len(), 1);
        let dv = &ws.data_validations[0];
        assert_eq!(dv.validation_type, "list");
        assert!(dv.operator.is_none());
        assert_eq!(dv.formula1.as_deref(), Some("\"Yes,No,Maybe\""));
        assert!(!dv.allow_blank);
    }

    /// The childless Empty-tag form (`<dataValidation .../>`, no
    /// `<formula1>`) must still parse its attributes correctly.
    #[test]
    fn test_data_validation_empty_tag_form_still_parses_attrs() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="C1"><v>1</v></c></row></sheetData>
  <dataValidations count="1">
    <dataValidation type="textLength" operator="lessThanOrEqual" sqref="C1:C5"/>
  </dataValidations>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        assert_eq!(ws.data_validations.len(), 1);
        let dv = &ws.data_validations[0];
        assert_eq!(dv.validation_type, "textLength");
        assert_eq!(dv.operator.as_deref(), Some("lessThanOrEqual"));
        assert!(dv.formula1.is_none());
    }

    /// A sheet with no `<dataValidations>` at all must produce an empty
    /// list, not an error.
    #[test]
    fn test_no_data_validations_is_an_empty_list() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        assert!(ws.data_validations.is_empty());
    }

    /// Turning trimming off for the inline-string body must not leak into
    /// the rest of the sheet: `<v>` is still parsed with surrounding
    /// whitespace ignored.
    #[test]
    fn test_inline_string_does_not_leave_the_sheet_reader_untrimmed() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t> padded </t></is></c>
      <c r="B1"><v>
        42
      </v></c>
    </row>
  </sheetData>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        assert!(
            matches!(&ws.rows[0].cells[0].value, CellValue::String(s) if s == " padded "),
            "got {:?}",
            ws.rows[0].cells[0].value
        );
        assert!(
            matches!(ws.rows[0].cells[1].value, CellValue::Number(n) if n == 42.0),
            "got {:?}",
            ws.rows[0].cells[1].value
        );
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
        // paperSize=9 = A4 (11906x16838 twips portrait), rotated because
        // orientation="landscape" — `paperSize` names a portrait stock.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <pageSetup paperSize="9" orientation="landscape"/>
</worksheet>"#;
        let ws = Worksheet::parse(xml, "S".to_string(), &empty_rels()).unwrap();
        let ps = ws.page_setup.expect("page_setup parsed");
        assert_eq!(ps.width_twips, 16838);
        assert_eq!(ps.height_twips, 11906);
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
        // Letter is 12240x15840 portrait; landscape swaps the two.
        assert_eq!(ps.width_twips, 15840);
        assert_eq!(ps.height_twips, 12240);
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
