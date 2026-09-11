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
        CellValue::Number(n) if n.is_finite() => {
            format!(r#"<c r="{cell_ref}"{attrs}><v>{n}</v></c>"#)
        },
        // NaN and the infinities have no xsd:double lexical form. The writer
        // maps them to an error cell; the editor must agree.
        CellValue::Number(_) => {
            format!(r#"<c r="{cell_ref}"{attrs} t="e"><v>#NUM!</v></c>"#)
        },
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

/// Byte offset just past the row element that starts at `row_start`, plus the
/// offset at which a new cell should be inserted to keep columns ascending.
///
/// Handles a self-closing `<row .../>`, which is what Excel writes for a row
/// that carries only formatting (a custom height). Searching for `</row>`
/// without checking would run straight past such a row and land on the NEXT
/// row's closing tag, putting the cell in the wrong row entirely.
enum RowShape {
    /// `<row …/>` — must be expanded into an open/close pair.
    SelfClosing { open_start: usize, open_end: usize },
    /// `<row …> … </row>` — the row body spans `body_start..body_end`.
    Paired { body_start: usize, body_end: usize },
}

fn locate_row(xml: &str, row: u32) -> Option<RowShape> {
    let row_start = xml.find(&format!(r#"<row r="{row}""#))?;
    let rest = &xml[row_start..];
    let tag_end = rest.find('>')?;
    let open_tag = &rest[..=tag_end];

    if open_tag.ends_with("/>") {
        return Some(RowShape::SelfClosing {
            open_start: row_start,
            open_end: row_start + tag_end + 1,
        });
    }

    let body_start = row_start + tag_end + 1;
    let body_end = body_start + xml[body_start..].find("</row>")?;
    Some(RowShape::Paired {
        body_start,
        body_end,
    })
}

/// Zero-based column index for a run of column letters (`A` -> 0, `AA` -> 26).
fn col_index(letters: &str) -> u32 {
    letters.bytes().fold(0u32, |acc, b| {
        acc.saturating_mul(26)
            .saturating_add(u32::from(b.to_ascii_uppercase() - b'A') + 1)
    })
}

/// Offset within `row_body` at which a cell in `col` keeps `<c>` ascending.
///
/// [ECMA-376] §18.3.1.73 requires cells in ascending column order; Excel
/// repairs a sheet whose cells are out of order.
fn cell_insert_offset(row_body: &str, col: &str) -> usize {
    let col = col_index(col);
    let mut search = 0usize;
    while let Some(rel) = row_body[search..].find("<c r=\"") {
        let at = search + rel;
        let ref_start = at + 6;
        let Some(len) = row_body[ref_start..].find('"') else {
            break;
        };
        let existing = &row_body[ref_start..ref_start + len];
        match parse_cell_ref(existing) {
            Some((_, existing_col)) if col_index(existing_col) > col => return at,
            _ => {},
        }
        search = ref_start + len;
    }
    row_body.len()
}

/// Offset within `sheetData` at which a new `<row r="N">` keeps rows ascending.
fn row_insert_offset(xml: &str, sd_body_start: usize, sd_end: usize, row: u32) -> usize {
    let body = &xml[sd_body_start..sd_end];
    let mut search = 0usize;
    while let Some(rel) = body[search..].find("<row r=\"") {
        let at = search + rel;
        let num_start = at + 8;
        let Some(len) = body[num_start..].find('"') else {
            break;
        };
        if let Ok(existing) = body[num_start..num_start + len].parse::<u32>() {
            if existing > row {
                return sd_body_start + at;
            }
        }
        search = num_start + len;
    }
    sd_end
}

/// Set a cell value in worksheet XML by string manipulation.
///
/// If the cell already exists, its value is replaced and its attributes (style, etc.)
/// are preserved. Otherwise, a new cell is inserted in column order, into a row
/// inserted in row order.
fn set_cell_in_xml(xml: &str, cell_ref: &str, value: &CellValue) -> String {
    if let Some(replaced) = replace_existing_cell(xml, cell_ref, value) {
        return replaced;
    }

    let cell_xml = render_cell(cell_ref, "", value);

    // Cell doesn't exist — find the right row or create one
    let Some((row, col)) = parse_cell_ref(cell_ref) else {
        return xml.to_string();
    };

    match locate_row(xml, row) {
        Some(RowShape::SelfClosing {
            open_start,
            open_end,
        }) => {
            // `<row …/>` becomes `<row …>cell</row>`.
            let open_tag = &xml[open_start..open_end];
            let reopened = format!("{}>{cell_xml}</row>", &open_tag[..open_tag.len() - 2]);
            let mut result = String::with_capacity(xml.len() + reopened.len());
            result.push_str(&xml[..open_start]);
            result.push_str(&reopened);
            result.push_str(&xml[open_end..]);
            return result;
        },
        Some(RowShape::Paired {
            body_start,
            body_end,
        }) => {
            let at = body_start + cell_insert_offset(&xml[body_start..body_end], col);
            let mut result = String::with_capacity(xml.len() + cell_xml.len());
            result.push_str(&xml[..at]);
            result.push_str(&cell_xml);
            result.push_str(&xml[at..]);
            return result;
        },
        None => {},
    }

    // Row doesn't exist — insert it in ascending row order.
    if let Some(sd_end) = xml.find("</sheetData>") {
        let sd_body_start = xml
            .find("<sheetData")
            .and_then(|i| xml[i..].find('>').map(|j| i + j + 1))
            .unwrap_or(sd_end);
        let at = row_insert_offset(xml, sd_body_start, sd_end, row);
        let row_xml = format!(r#"<row r="{row}">{cell_xml}</row>"#);
        let mut result = String::with_capacity(xml.len() + row_xml.len());
        result.push_str(&xml[..at]);
        result.push_str(&row_xml);
        result.push_str(&xml[at..]);
        return result;
    }

    xml.to_string()
}

fn escape_xml(s: &str) -> String {
    // Strip characters XML cannot represent before escaping the rest;
    // otherwise a control byte makes the whole part unparseable.
    crate::core::xml::sanitize_xml_text(s)
        .replace('&', "&amp;")
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

#[cfg(test)]
mod ordering_tests {
    use super::*;

    /// Excel writes `<row r="N" ht="…" customHeight="1"/>` for a row that
    /// carries only formatting. Delimiting the row by searching for `</row>`
    /// runs past such a row and lands on the NEXT row's closing tag, so the
    /// inserted cell ends up inside the wrong row entirely.
    #[test]
    fn self_closing_row_does_not_swallow_the_cell_into_the_next_row() {
        let xml = concat!(
            r#"<sheetData>"#,
            r#"<row r="1" ht="30" customHeight="1"/>"#,
            r#"<row r="2"><c r="A2" t="inlineStr"><is><t>row2A</t></is></c></row>"#,
            r#"</sheetData>"#
        );
        let out = set_cell_in_xml(xml, "A1", &CellValue::Number(111.0));

        let row1 = out.find(r#"<row r="1""#).unwrap();
        let row2 = out.find(r#"<row r="2""#).unwrap();
        let a1 = out.find(r#"<c r="A1""#).expect("A1 must be written");
        assert!(row1 < a1 && a1 < row2, "A1 must live inside row 1, not row 2:\n{out}");
        assert!(out.contains("row2A"), "row 2 content must survive");
        assert!(out.contains(r#"ht="30""#), "row formatting must survive");
    }

    /// [ECMA-376] §18.3.1.73: cells ascend by column within a row.
    #[test]
    fn new_cell_is_inserted_in_ascending_column_order() {
        let xml = concat!(
            r#"<sheetData><row r="2">"#,
            r#"<c r="B2"><v>2</v></c><c r="D2"><v>4</v></c>"#,
            r#"</row></sheetData>"#
        );
        let out = set_cell_in_xml(xml, "C2", &CellValue::Number(3.0));
        let b = out.find(r#"<c r="B2""#).unwrap();
        let c = out.find(r#"<c r="C2""#).expect("C2 written");
        let d = out.find(r#"<c r="D2""#).unwrap();
        assert!(b < c && c < d, "cells must ascend by column:\n{out}");

        let out = set_cell_in_xml(xml, "A2", &CellValue::Number(1.0));
        let a = out.find(r#"<c r="A2""#).expect("A2 written");
        let b = out.find(r#"<c r="B2""#).unwrap();
        assert!(a < b, "a new first column must come first:\n{out}");
    }

    /// Rows ascend by index within `sheetData`.
    #[test]
    fn new_row_is_inserted_in_ascending_row_order() {
        let xml = r#"<sheetData><row r="2"><c r="A2"><v>2</v></c></row></sheetData>"#;
        let out = set_cell_in_xml(xml, "A1", &CellValue::Number(1.0));
        let r1 = out.find(r#"<row r="1""#).expect("row 1 written");
        let r2 = out.find(r#"<row r="2""#).unwrap();
        assert!(r1 < r2, "rows must ascend:\n{out}");
    }

    /// Control characters make the part unparseable; the writer already
    /// guards this and the editor must agree.
    #[test]
    fn control_characters_are_stripped_from_cell_text() {
        let xml = r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#;
        let out = set_cell_in_xml(xml, "B1", &CellValue::String("ctl\u{1}here".into()));
        assert!(!out.contains('\u{1}'), "control char reached the XML:\n{out}");
        assert!(out.contains("ctlhere"), "surrounding text must survive:\n{out}");
    }

    /// `NaN`/`inf` have no `xsd:double` lexical form.
    #[test]
    fn non_finite_numbers_become_error_cells() {
        let xml = r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#;
        for v in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            let out = set_cell_in_xml(xml, "B1", &CellValue::Number(v));
            assert!(out.contains(r#"t="e""#), "expected an error cell for {v}:\n{out}");
            assert!(out.contains("#NUM!"), "expected #NUM! for {v}");
            assert!(!out.contains("NaN") && !out.contains("inf"), "raw float written:\n{out}");
        }
    }
}
