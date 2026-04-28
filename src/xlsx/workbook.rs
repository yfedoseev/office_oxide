use quick_xml::events::Event;

use crate::core::xml;

/// Parsed workbook information from `xl/workbook.xml`.
#[derive(Debug, Clone)]
pub struct WorkbookInfo {
    /// Ordered list of sheet metadata entries.
    pub sheets: Vec<SheetInfo>,
    /// Named ranges and print-area definitions.
    pub defined_names: Vec<DefinedName>,
    /// Whether this workbook uses the 1904 date system (common in Mac-created files).
    pub date1904: bool,
}

/// Metadata for a single sheet in the workbook.
#[derive(Debug, Clone)]
pub struct SheetInfo {
    /// Sheet display name.
    pub name: String,
    /// Numeric sheet ID from the workbook XML.
    pub sheet_id: u32,
    /// Relationship ID used to resolve the worksheet part path.
    pub rel_id: String,
    /// Visibility state of this sheet.
    pub state: SheetState,
}

/// Sheet visibility state.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetState {
    /// Sheet is visible.
    Visible,
    /// Sheet is hidden but can be un-hidden by the user.
    Hidden,
    /// Sheet is very hidden (only un-hideable via VBA).
    VeryHidden,
}

/// A defined name in the workbook (named ranges, print areas, etc.).
#[derive(Debug, Clone)]
pub struct DefinedName {
    /// Name string.
    pub name: String,
    /// Formula or reference value.
    pub value: String,
    /// If set, this name is scoped to a specific sheet (0-based index).
    pub local_sheet_id: Option<u32>,
    /// Whether this name is hidden.
    pub hidden: bool,
}

impl WorkbookInfo {
    /// Parse `xl/workbook.xml` from raw XML bytes.
    pub fn parse(xml_data: &[u8]) -> crate::core::Result<Self> {
        let mut reader = xml::make_fast_reader(xml_data);
        let mut sheets = Vec::new();
        let mut defined_names = Vec::new();
        let mut date1904 = false;

        loop {
            match reader.read_event()? {
                Event::Start(ref e) => {
                    match e.local_name().as_ref() {
                        b"sheet" => {
                            let name = xml::required_attr_str(e, b"name")?.into_owned();
                            let sheet_id: u32 = xml::required_attr_str(e, b"sheetId")?.parse()?;
                            // Try r:id first, then fall back to any prefixed `id` attribute
                            // (some files use d3p1:id or other namespace prefixes)
                            let rel_id = match xml::optional_attr_str(e, b"r:id")? {
                                Some(v) => v.into_owned(),
                                None => xml::optional_prefixed_attr_str(e, b"id")?
                                    .map(|v| v.into_owned())
                                    .unwrap_or_default(),
                            };
                            let state = match xml::optional_attr_str(e, b"state")? {
                                Some(ref v) => match v.as_ref() {
                                    "hidden" => SheetState::Hidden,
                                    "veryHidden" => SheetState::VeryHidden,
                                    _ => SheetState::Visible,
                                },
                                None => SheetState::Visible,
                            };
                            sheets.push(SheetInfo {
                                name,
                                sheet_id,
                                rel_id,
                                state,
                            });
                            xml::skip_element_fast(&mut reader)?;
                        },
                        b"workbookPr" => {
                            if let Some(val) = xml::optional_attr_str(e, b"date1904")? {
                                date1904 = matches!(val.as_ref(), "1" | "true");
                            }
                            xml::skip_element_fast(&mut reader)?;
                        },
                        b"definedName" => {
                            let name = xml::required_attr_str(e, b"name")?.into_owned();
                            let local_sheet_id = xml::optional_attr_str(e, b"localSheetId")?
                                .and_then(|v| v.parse().ok());
                            let hidden = xml::optional_attr_str(e, b"hidden")?
                                .is_some_and(|v| matches!(v.as_ref(), "1" | "true"));
                            let value = xml::read_text_content_fast(&mut reader)?;
                            defined_names.push(DefinedName {
                                name,
                                value,
                                local_sheet_id,
                                hidden,
                            });
                        },
                        _ => {},
                    }
                },
                Event::Empty(ref e) => match e.local_name().as_ref() {
                    b"sheet" => {
                        let name = xml::required_attr_str(e, b"name")?.into_owned();
                        let sheet_id: u32 = xml::required_attr_str(e, b"sheetId")?.parse()?;
                        let rel_id = match xml::optional_attr_str(e, b"r:id")? {
                            Some(v) => v.into_owned(),
                            None => xml::optional_prefixed_attr_str(e, b"id")?
                                .map(|v| v.into_owned())
                                .unwrap_or_default(),
                        };
                        let state = match xml::optional_attr_str(e, b"state")? {
                            Some(ref v) => match v.as_ref() {
                                "hidden" => SheetState::Hidden,
                                "veryHidden" => SheetState::VeryHidden,
                                _ => SheetState::Visible,
                            },
                            None => SheetState::Visible,
                        };
                        sheets.push(SheetInfo {
                            name,
                            sheet_id,
                            rel_id,
                            state,
                        });
                    },
                    b"workbookPr" => {
                        if let Some(val) = xml::optional_attr_str(e, b"date1904")? {
                            date1904 = matches!(val.as_ref(), "1" | "true");
                        }
                    },
                    _ => {},
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(WorkbookInfo {
            sheets,
            defined_names,
            date1904,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_workbook_with_sheets() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2" state="hidden"/>
  </sheets>
</workbook>"#;
        let wb = WorkbookInfo::parse(xml).unwrap();
        assert_eq!(wb.sheets.len(), 2);
        assert_eq!(wb.sheets[0].name, "Sheet1");
        assert_eq!(wb.sheets[0].sheet_id, 1);
        assert_eq!(wb.sheets[0].rel_id, "rId1");
        assert_eq!(wb.sheets[0].state, SheetState::Visible);
        assert_eq!(wb.sheets[1].name, "Sheet2");
        assert_eq!(wb.sheets[1].state, SheetState::Hidden);
    }

    #[test]
    fn parse_workbook_date1904() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr date1904="1"/>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;
        let wb = WorkbookInfo::parse(xml).unwrap();
        assert!(wb.date1904);
    }

    #[test]
    fn parse_workbook_defined_names() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$D$10</definedName>
    <definedName name="MyRange" hidden="1">Sheet1!$A:$A</definedName>
  </definedNames>
</workbook>"#;
        let wb = WorkbookInfo::parse(xml).unwrap();
        assert_eq!(wb.defined_names.len(), 2);
        assert_eq!(wb.defined_names[0].name, "_xlnm.Print_Area");
        assert_eq!(wb.defined_names[0].value, "Sheet1!$A$1:$D$10");
        assert_eq!(wb.defined_names[0].local_sheet_id, Some(0));
        assert!(!wb.defined_names[0].hidden);
        assert_eq!(wb.defined_names[1].name, "MyRange");
        assert!(wb.defined_names[1].hidden);
    }
}
