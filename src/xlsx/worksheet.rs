use quick_xml::events::Event;

use crate::core::xml;

use super::cell::{Cell, CellRef, CellValue};

/// A parsed worksheet from `xl/worksheets/sheetN.xml`.
#[derive(Debug, Clone)]
pub struct Worksheet {
    pub name: String,
    /// Dimension string like "A1:G50", if present.
    pub dimension: Option<String>,
    pub rows: Vec<Row>,
    /// Merged cell ranges like "A1:C1".
    pub merged_cells: Vec<String>,
    pub hyperlinks: Vec<HyperlinkInfo>,
}

/// A row from `<sheetData>`.
#[derive(Debug, Clone)]
pub struct Row {
    /// 1-based row number from the `r` attribute.
    pub index: u32,
    pub cells: Vec<Cell>,
}

/// Hyperlink information from `<hyperlinks>`.
#[derive(Debug, Clone)]
pub struct HyperlinkInfo {
    /// Cell reference like "A1".
    pub cell_ref: String,
    pub target: HyperlinkTarget,
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
                    _ => {},
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(Worksheet {
            name,
            dimension,
            rows,
            merged_cells,
            hyperlinks,
        })
    }
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
            Event::Empty(ref e) => {
                if e.local_name().as_ref() == b"c" {
                    cells.push(parse_empty_cell(e)?);
                }
            },
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"row" {
                    break;
                }
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
            Event::Empty(ref e) => {
                if e.local_name().as_ref() == b"f" {
                    formula = None;
                }
            },
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"c" {
                    break;
                }
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
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"is" {
                    break;
                }
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
}
