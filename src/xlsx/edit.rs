//! XLSX editing via raw XML manipulation.
//!
//! Uses the `EditablePackage` from core to preserve all parts,
//! modifying individual cells in worksheet XML.

use crate::core::editable::EditablePackage;
use crate::core::opc::PartName;

use super::Result;

/// The value to set in a cell.
#[derive(Debug, Clone)]
pub enum CellValue {
    /// An empty cell.
    Empty,
    /// A string value.
    String(String),
    /// A numeric value.
    Number(f64),
    /// A boolean value.
    Boolean(bool),
}

/// An editable XLSX document that supports cell modification and saving.
pub struct EditableXlsx {
    package: EditablePackage,
}

impl EditableXlsx {
    /// Open an XLSX file for editing.
    pub fn open(path: impl AsRef<std::path::Path>) -> Result<Self> {
        let package = EditablePackage::open(&path)?;
        Ok(Self { package })
    }

    /// Open from any `Read + Seek` source.
    pub fn from_reader<R: std::io::Read + std::io::Seek>(reader: R) -> Result<Self> {
        let package = EditablePackage::from_reader(reader)?;
        Ok(Self { package })
    }

    /// Set a cell value in a worksheet.
    ///
    /// `sheet_index` is 0-based. `cell_ref` is like "A1", "B2", etc.
    pub fn set_cell(&mut self, sheet_index: usize, cell_ref: &str, value: CellValue) -> Result<()> {
        let part_name = PartName::new(&format!("/xl/worksheets/sheet{}.xml", sheet_index + 1))?;
        let Some(data) = self.package.get_part(&part_name) else {
            return Err(super::XlsxError::Core(crate::core::Error::MissingPart(
                part_name.as_str().to_string(),
            )));
        };
        let xml_str = String::from_utf8_lossy(data).into_owned();
        let new_xml = set_cell_in_xml(&xml_str, cell_ref, &value);
        self.package.set_part(part_name, new_xml.into_bytes());
        Ok(())
    }

    /// Save the edited document to a file.
    pub fn save(&self, path: impl AsRef<std::path::Path>) -> Result<()> {
        self.package.save(path)?;
        Ok(())
    }

    /// Write the edited document to any `Write + Seek` destination.
    pub fn write_to<W: std::io::Write + std::io::Seek>(&self, writer: W) -> Result<()> {
        self.package.write_to(writer)?;
        Ok(())
    }
}

/// Parse a cell reference like "A1" into (row_1based, col_letters).
fn parse_cell_ref(cell_ref: &str) -> Option<(u32, &str)> {
    let col_end = cell_ref.bytes().position(|b| b.is_ascii_digit())?;
    if col_end == 0 {
        return None;
    }
    let col_str = &cell_ref[..col_end];
    let row: u32 = cell_ref[col_end..].parse().ok()?;
    Some((row, col_str))
}

/// Set a cell value in worksheet XML by string manipulation.
/// If the cell already exists, its value is replaced. Otherwise, a new cell is inserted.
fn set_cell_in_xml(xml: &str, cell_ref: &str, value: &CellValue) -> String {
    let cell_xml = match value {
        CellValue::Empty => format!(r#"<c r="{cell_ref}"/>"#),
        CellValue::String(s) => {
            let escaped = escape_xml(s);
            format!(r#"<c r="{cell_ref}" t="inlineStr"><is><t>{escaped}</t></is></c>"#)
        },
        CellValue::Number(n) => format!(r#"<c r="{cell_ref}"><v>{n}</v></c>"#),
        CellValue::Boolean(b) => {
            let v = if *b { "1" } else { "0" };
            format!(r#"<c r="{cell_ref}" t="b"><v>{v}</v></c>"#)
        },
    };

    // Try to find existing cell with this reference
    let cell_pattern = format!(r#"<c r="{cell_ref}""#);
    if let Some(start) = xml.find(&cell_pattern) {
        // Find end of this cell element
        let rest = &xml[start..];
        let end = if let Some(close) = rest.find("</c>") {
            start + close + 4
        } else if let Some(close) = rest.find("/>") {
            start + close + 2
        } else {
            return xml.to_string();
        };

        let mut result = String::with_capacity(xml.len());
        result.push_str(&xml[..start]);
        result.push_str(&cell_xml);
        result.push_str(&xml[end..]);
        return result;
    }

    // Cell doesn't exist — find the right row or create one
    let Some((row, _col)) = parse_cell_ref(cell_ref) else {
        return xml.to_string();
    };

    let row_pattern = format!(r#"<row r="{row}""#);
    if let Some(row_start) = xml.find(&row_pattern) {
        // Row exists — insert cell before </row>
        let rest = &xml[row_start..];
        if let Some(row_end) = rest.find("</row>") {
            let insert_pos = row_start + row_end;
            let mut result = String::with_capacity(xml.len() + cell_xml.len());
            result.push_str(&xml[..insert_pos]);
            result.push_str(&cell_xml);
            result.push_str(&xml[insert_pos..]);
            return result;
        }
    }

    // Row doesn't exist — insert before </sheetData>
    if let Some(sd_end) = xml.find("</sheetData>") {
        let row_xml = format!(r#"<row r="{row}">{cell_xml}</row>"#);
        let mut result = String::with_capacity(xml.len() + row_xml.len());
        result.push_str(&xml[..sd_end]);
        result.push_str(&row_xml);
        result.push_str(&xml[sd_end..]);
        return result;
    }

    xml.to_string()
}

fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn set_existing_cell() {
        let xml = r#"<sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData>"#;
        let result = set_cell_in_xml(xml, "A1", &CellValue::Number(99.0));
        assert!(result.contains(r#"<c r="A1"><v>99</v></c>"#));
        assert!(!result.contains("42"));
    }

    #[test]
    fn set_new_cell_existing_row() {
        let xml = r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#;
        let result = set_cell_in_xml(xml, "B1", &CellValue::String("hello".into()));
        assert!(result.contains(r#"<c r="B1" t="inlineStr"><is><t>hello</t></is></c>"#));
        assert!(result.contains(r#"<c r="A1"><v>1</v></c>"#));
    }

    #[test]
    fn set_cell_new_row() {
        let xml = r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#;
        let result = set_cell_in_xml(xml, "A2", &CellValue::Number(2.0));
        assert!(result.contains(r#"<row r="2"><c r="A2"><v>2</v></c></row>"#));
    }

    #[test]
    fn set_boolean_cell() {
        let xml = r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#;
        let result = set_cell_in_xml(xml, "A1", &CellValue::Boolean(true));
        assert!(result.contains(r#"<c r="A1" t="b"><v>1</v></c>"#));
    }

    #[test]
    fn parse_cell_ref_valid() {
        assert_eq!(parse_cell_ref("A1"), Some((1, "A")));
        assert_eq!(parse_cell_ref("ZZ100"), Some((100, "ZZ")));
    }
}
