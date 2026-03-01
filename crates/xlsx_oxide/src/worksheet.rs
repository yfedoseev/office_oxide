use quick_xml::events::Event;

use office_core::xml;

use crate::cell::{Cell, CellRef, CellValue};

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
        rels: &office_core::relationships::Relationships,
    ) -> office_core::Result<Self> {
        let mut reader = xml::make_reader(xml_data);
        let sml = xml::ns::SML;
        let mut dimension = None;
        let mut rows = Vec::new();
        let mut merged_cells = Vec::new();
        let mut hyperlinks = Vec::new();

        loop {
            match reader.read_resolved_event()? {
                (ref resolve, Event::Start(ref e)) => {
                    if xml::matches_ns(resolve, sml) {
                        match e.local_name().as_ref() {
                            b"dimension" => {
                                dimension = xml::optional_attr_str(e, b"ref")?
                                    .map(|v| v.into_owned());
                                xml::skip_element(&mut reader)?;
                            }
                            b"row" => {
                                rows.push(parse_row(&mut reader, e)?);
                            }
                            b"mergeCell" => {
                                if let Some(range) = xml::optional_attr_str(e, b"ref")? {
                                    merged_cells.push(range.into_owned());
                                }
                                xml::skip_element(&mut reader)?;
                            }
                            b"hyperlink" => {
                                if let Some(hl) = parse_hyperlink(e, rels)? {
                                    hyperlinks.push(hl);
                                }
                                xml::skip_element(&mut reader)?;
                            }
                            _ => {}
                        }
                    }
                }
                (ref resolve, Event::Empty(ref e)) => {
                    if xml::matches_ns(resolve, sml) {
                        match e.local_name().as_ref() {
                            b"dimension" => {
                                dimension = xml::optional_attr_str(e, b"ref")?
                                    .map(|v| v.into_owned());
                            }
                            b"mergeCell" => {
                                if let Some(range) = xml::optional_attr_str(e, b"ref")? {
                                    merged_cells.push(range.into_owned());
                                }
                            }
                            b"hyperlink" => {
                                if let Some(hl) = parse_hyperlink(e, rels)? {
                                    hyperlinks.push(hl);
                                }
                            }
                            _ => {}
                        }
                    }
                }
                (_, Event::Eof) => break,
                _ => {}
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
    rels: &office_core::relationships::Relationships,
) -> office_core::Result<Option<HyperlinkInfo>> {
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

fn parse_row(
    reader: &mut quick_xml::NsReader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> office_core::Result<Row> {
    let sml = xml::ns::SML;
    let index: u32 = xml::optional_attr_str(start, b"r")?
        .and_then(|v| v.parse().ok())
        .unwrap_or(1);
    let mut cells = Vec::new();

    loop {
        match reader.read_resolved_event()? {
            (ref resolve, Event::Start(ref e)) => {
                if xml::matches_ns(resolve, sml) && e.local_name().as_ref() == b"c" {
                    cells.push(parse_cell(reader, e)?);
                } else {
                    xml::skip_element(reader)?;
                }
            }
            (ref resolve, Event::Empty(ref e)) => {
                if xml::matches_ns(resolve, sml) && e.local_name().as_ref() == b"c" {
                    // Empty cell element (no value/formula children)
                    cells.push(parse_empty_cell(e)?);
                }
            }
            (ref resolve, Event::End(ref e)) => {
                if xml::matches_ns(resolve, sml) && e.local_name().as_ref() == b"row" {
                    break;
                }
            }
            (_, Event::Eof) => break,
            _ => {}
        }
    }

    Ok(Row { index, cells })
}

fn parse_empty_cell(e: &quick_xml::events::BytesStart) -> office_core::Result<Cell> {
    let ref_str = xml::optional_attr_str(e, b"r")?
        .map(|v| v.into_owned())
        .unwrap_or_default();
    let reference = CellRef::parse(&ref_str).unwrap_or(CellRef { col: 0, row: 0 });
    let style_index = xml::optional_attr_str(e, b"s")?
        .and_then(|v| v.parse().ok());

    Ok(Cell {
        reference,
        value: CellValue::Empty,
        style_index,
        formula: None,
    })
}

fn parse_cell(
    reader: &mut quick_xml::NsReader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> office_core::Result<Cell> {
    let sml = xml::ns::SML;

    let ref_str = xml::optional_attr_str(start, b"r")?
        .map(|v| v.into_owned())
        .unwrap_or_default();
    let reference = CellRef::parse(&ref_str).unwrap_or(CellRef { col: 0, row: 0 });

    let cell_type = xml::optional_attr_str(start, b"t")?
        .map(|v| v.into_owned());
    let style_index = xml::optional_attr_str(start, b"s")?
        .and_then(|v| v.parse().ok());

    let mut raw_value: Option<String> = None;
    let mut formula: Option<String> = None;

    loop {
        match reader.read_resolved_event()? {
            (ref resolve, Event::Start(ref e)) => {
                if xml::matches_ns(resolve, sml) {
                    match e.local_name().as_ref() {
                        b"v" => {
                            raw_value = Some(xml::read_text_content(reader)?);
                        }
                        b"f" => {
                            formula = Some(xml::read_text_content(reader)?);
                        }
                        b"is" => {
                            // Inline string: <is><t>text</t></is>
                            raw_value = Some(parse_inline_string(reader)?);
                        }
                        _ => {
                            xml::skip_element(reader)?;
                        }
                    }
                } else {
                    xml::skip_element(reader)?;
                }
            }
            (ref resolve, Event::Empty(ref e)) => {
                if xml::matches_ns(resolve, sml) && e.local_name().as_ref() == b"f" {
                    // Empty formula element (e.g., shared formula reference)
                    formula = None;
                }
            }
            (ref resolve, Event::End(ref e)) => {
                if xml::matches_ns(resolve, sml) && e.local_name().as_ref() == b"c" {
                    break;
                }
            }
            (_, Event::Eof) => break,
            _ => {}
        }
    }

    let value = match cell_type.as_deref() {
        Some("s") => {
            // Shared string index
            match raw_value.as_deref().and_then(|v| v.parse::<u32>().ok()) {
                Some(idx) => CellValue::SharedString(idx),
                None => CellValue::Empty,
            }
        }
        Some("str") | Some("inlineStr") => {
            // Inline string or formula string result
            match raw_value {
                Some(s) => CellValue::String(s),
                None => CellValue::Empty,
            }
        }
        Some("b") => {
            // Boolean
            match raw_value.as_deref() {
                Some("1") | Some("true") => CellValue::Boolean(true),
                Some("0") | Some("false") => CellValue::Boolean(false),
                _ => CellValue::Empty,
            }
        }
        Some("e") => {
            // Error
            match raw_value {
                Some(s) => CellValue::Error(s),
                None => CellValue::Error(String::new()),
            }
        }
        _ => {
            // Number (default type) or empty
            match raw_value {
                Some(s) => match s.parse::<f64>() {
                    Ok(n) => CellValue::Number(n),
                    Err(_) => CellValue::String(s),
                },
                None => CellValue::Empty,
            }
        }
    };

    Ok(Cell {
        reference,
        value,
        style_index,
        formula,
    })
}

/// Parse inline string content: `<is><t>text</t></is>` or `<is><r>...</r></is>`.
fn parse_inline_string(reader: &mut quick_xml::NsReader<&[u8]>) -> office_core::Result<String> {
    let sml = xml::ns::SML;
    let mut text = String::new();

    loop {
        match reader.read_resolved_event()? {
            (ref resolve, Event::Start(ref e)) => {
                if xml::matches_ns(resolve, sml) && e.local_name().as_ref() == b"t" {
                    text.push_str(&xml::read_text_content(reader)?);
                } else {
                    xml::skip_element(reader)?;
                }
            }
            (ref resolve, Event::End(ref e)) => {
                if xml::matches_ns(resolve, sml) && e.local_name().as_ref() == b"is" {
                    break;
                }
            }
            (_, Event::Eof) => break,
            _ => {}
        }
    }

    Ok(text)
}

#[cfg(test)]
mod tests {
    use super::*;
    use office_core::relationships::Relationships;

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
