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

/// Collect the attributes of a `<c ...>` opening tag, minus the ones we re-emit.
///
/// `r` (the reference) and `t` (the type) are always rewritten from the new value.
/// Everything else — crucially `s`, the style index, but also `cm`, `vm` and `ph` —
/// is carried through verbatim, so a written cell keeps its number format, fill and
/// font instead of coming back unstyled.
fn preserved_attrs(open_tag: &str) -> String {
    let body = open_tag
        .trim_start_matches("<c")
        .trim_end_matches('>')
        .trim_end_matches('/');

    let mut out = String::new();
    let bytes = body.as_bytes();
    let mut i = 0;

    while i < bytes.len() {
        while i < bytes.len() && bytes[i].is_ascii_whitespace() {
            i += 1;
        }
        let name_start = i;
        while i < bytes.len() && bytes[i] != b'=' && !bytes[i].is_ascii_whitespace() {
            i += 1;
        }
        let name = &body[name_start..i];

        while i < bytes.len() && (bytes[i].is_ascii_whitespace() || bytes[i] == b'=') {
            i += 1;
        }
        if i >= bytes.len() || bytes[i] != b'"' {
            break;
        }
        i += 1;

        let value_start = i;
        while i < bytes.len() && bytes[i] != b'"' {
            i += 1;
        }
        let value = &body[value_start..i.min(body.len())];
        i += 1;

        if !name.is_empty() && name != "r" && name != "t" {
            out.push_str(&format!(r#" {name}="{value}""#));
        }
    }

    out
}

/// Render a `<c>` element, carrying the original cell's preserved attributes.
fn render_cell(cell_ref: &str, attrs: &str, value: &CellValue) -> String {
    match value {
        CellValue::Empty => format!(r#"<c r="{cell_ref}"{attrs}/>"#),
        CellValue::String(s) => {
            let escaped = escape_xml(s);
            format!(r#"<c r="{cell_ref}"{attrs} t="inlineStr"><is><t>{escaped}</t></is></c>"#)
        },
        CellValue::Number(n) => format!(r#"<c r="{cell_ref}"{attrs}><v>{n}</v></c>"#),
        CellValue::Boolean(b) => {
            let v = if *b { "1" } else { "0" };
            format!(r#"<c r="{cell_ref}"{attrs} t="b"><v>{v}</v></c>"#)
        },
    }
}

/// Replace an existing `<c>` element in place, preserving its attributes.
///
/// Returns `None` if the cell is not present in `xml`.
fn replace_existing_cell(xml: &str, cell_ref: &str, value: &CellValue) -> Option<String> {
    let start = xml.find(&format!(r#"<c r="{cell_ref}""#))?;
    let rest = &xml[start..];

    // Delimit the OPENING tag first. A cell may be self-closing — `<c r="A1" s="5"/>`,
    // an empty but styled cell, which is what a pre-formatted template row is made of.
    // Searching for `</c>` first would run past such a cell and land on the NEXT cell's
    // closing tag, and the replacement would silently delete the cell in between.
    let tag_end = rest.find('>')?;
    let open_tag = &rest[..=tag_end];

    let end = if open_tag.ends_with("/>") {
        start + tag_end + 1
    } else {
        start + rest.find("</c>")? + 4
    };

    let attrs = preserved_attrs(open_tag);
    let cell_xml = render_cell(cell_ref, &attrs, value);

    let mut result = String::with_capacity(xml.len());
    result.push_str(&xml[..start]);
    result.push_str(&cell_xml);
    result.push_str(&xml[end..]);
    Some(result)
}

/// Set a cell value in worksheet XML by string manipulation.
///
/// If the cell already exists, its value is replaced and its attributes (style, etc.)
/// are preserved. Otherwise, a new cell is inserted.
fn set_cell_in_xml(xml: &str, cell_ref: &str, value: &CellValue) -> String {
    if let Some(replaced) = replace_existing_cell(xml, cell_ref, value) {
        return replaced;
    }

    let cell_xml = render_cell(cell_ref, "", value);

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
    fn set_existing_cell_preserves_style_index() {
        // `s="5"` is the cell's style: its number format, fill and font. Rebuilding the
        // <c> element from scratch dropped it, and the written cell came back unstyled.
        let xml = r#"<sheetData><row r="1"><c r="A1" s="5" t="n"><v>42</v></c></row></sheetData>"#;
        let result = set_cell_in_xml(xml, "A1", &CellValue::Number(99.0));
        assert!(result.contains(r#"<c r="A1" s="5"><v>99</v></c>"#), "result: {result}");
    }

    #[test]
    fn set_existing_string_cell_preserves_style_index() {
        let xml = r#"<sheetData><row r="1"><c r="A1" s="3" t="inlineStr"><is><t>old</t></is></c></row></sheetData>"#;
        let result = set_cell_in_xml(xml, "A1", &CellValue::String("new".into()));
        assert!(
            result.contains(r#"<c r="A1" s="3" t="inlineStr"><is><t>new</t></is></c>"#),
            "result: {result}"
        );
    }

    #[test]
    fn set_existing_cell_preserves_unrelated_attributes() {
        let xml =
            r#"<sheetData><row r="1"><c r="A1" s="2" cm="1" vm="4"><v>1</v></c></row></sheetData>"#;
        let result = set_cell_in_xml(xml, "A1", &CellValue::Number(2.0));
        assert!(result.contains(r#"s="2""#), "result: {result}");
        assert!(result.contains(r#"cm="1""#), "result: {result}");
        assert!(result.contains(r#"vm="4""#), "result: {result}");
    }

    #[test]
    fn set_self_closing_cell_does_not_swallow_the_next_cell() {
        // An empty but STYLED cell is self-closing: `<c r="A1" s="5"/>` — no `</c>`.
        // Searching for `</c>` first found B1's closing tag instead, and the
        // replacement deleted B1 along the way. A pre-formatted template row is made
        // entirely of such cells, so the nominal case triggered the data loss.
        let xml = r#"<sheetData><row r="1"><c r="A1" s="5"/><c r="B1" s="6"><v>7</v></c></row></sheetData>"#;
        let result = set_cell_in_xml(xml, "A1", &CellValue::Number(1.0));

        assert!(result.contains(r#"<c r="A1" s="5"><v>1</v></c>"#), "result: {result}");
        assert!(
            result.contains(r#"<c r="B1" s="6"><v>7</v></c>"#),
            "the neighbouring cell must survive; result: {result}"
        );
    }

    #[test]
    fn set_self_closing_cell_to_empty_stays_self_closing() {
        let xml =
            r#"<sheetData><row r="1"><c r="A1" s="5"/><c r="B1"><v>7</v></c></row></sheetData>"#;
        let result = set_cell_in_xml(xml, "A1", &CellValue::Empty);
        assert!(result.contains(r#"<c r="A1" s="5"/>"#), "result: {result}");
        assert!(result.contains(r#"<c r="B1"><v>7</v></c>"#), "result: {result}");
    }

    #[test]
    fn set_new_cell_carries_no_borrowed_attributes() {
        let xml = r#"<sheetData><row r="1"><c r="A1" s="9"><v>1</v></c></row></sheetData>"#;
        let result = set_cell_in_xml(xml, "B1", &CellValue::Number(2.0));
        assert!(result.contains(r#"<c r="B1"><v>2</v></c>"#), "result: {result}");
    }

    #[test]
    fn parse_cell_ref_valid() {
        assert_eq!(parse_cell_ref("A1"), Some((1, "A")));
        assert_eq!(parse_cell_ref("ZZ100"), Some((100, "ZZ")));
    }
}
