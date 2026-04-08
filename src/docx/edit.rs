//! DOCX editing via raw XML text replacement.
//!
//! Uses the `EditablePackage` from core to preserve all parts,
//! replacing text in the document.xml body by string substitution
//! in the raw XML `<w:t>` elements.

use crate::core::editable::EditablePackage;
use crate::core::opc::PartName;

use super::Result;

/// An editable DOCX document that supports text replacement and saving.
pub struct EditableDocx {
    package: EditablePackage,
    main_part: PartName,
}

impl EditableDocx {
    /// Open a DOCX file for editing.
    pub fn open(path: impl AsRef<std::path::Path>) -> Result<Self> {
        let package = EditablePackage::open(&path)?;
        let main_part = PartName::new("/word/document.xml")?;
        Ok(Self { package, main_part })
    }

    /// Open from any `Read + Seek` source.
    pub fn from_reader<R: std::io::Read + std::io::Seek>(reader: R) -> Result<Self> {
        let package = EditablePackage::from_reader(reader)?;
        let main_part = PartName::new("/word/document.xml")?;
        Ok(Self { package, main_part })
    }

    /// Replace all occurrences of `find` with `replace` in the document body.
    /// Returns the number of replacements made.
    pub fn replace_text(&mut self, find: &str, replace: &str) -> usize {
        let Some(data) = self.package.get_part(&self.main_part) else {
            return 0;
        };
        let xml_str = String::from_utf8_lossy(data);

        // Replace text within <w:t> elements.
        // Strategy: find text between <w:t...> and </w:t> tags and do replacements there.
        let (new_xml, count) = replace_in_wt_elements(&xml_str, find, replace);

        if count > 0 {
            self.package.set_part(self.main_part.clone(), new_xml.into_bytes());
        }
        count
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

/// Replace text within `<w:t>...</w:t>` elements in a WML XML string.
/// Returns the new string and the count of replacements.
fn replace_in_wt_elements(xml: &str, find: &str, replace: &str) -> (String, usize) {
    let mut result = String::with_capacity(xml.len());
    let mut count = 0;
    let mut pos = 0;

    while pos < xml.len() {
        // Find next <w:t> or <w:t ...>
        if let Some(tag_start) = xml[pos..].find("<w:t") {
            let tag_start = pos + tag_start;

            // Find the end of the opening tag
            let Some(tag_end_offset) = xml[tag_start..].find('>') else {
                result.push_str(&xml[pos..]);
                break;
            };
            let tag_end = tag_start + tag_end_offset + 1;

            // Check if it's a self-closing tag
            if xml[tag_start..tag_end].ends_with("/>") {
                result.push_str(&xml[pos..tag_end]);
                pos = tag_end;
                continue;
            }

            // Find closing </w:t>
            let Some(close_offset) = xml[tag_end..].find("</w:t>") else {
                result.push_str(&xml[pos..]);
                break;
            };
            let close_start = tag_end + close_offset;

            // Extract text content between tags
            let text_content = &xml[tag_end..close_start];
            count += text_content.matches(find).count();
            let replaced = text_content.replace(find, replace);

            // Write: everything before this tag + tag + replaced text + close tag
            result.push_str(&xml[pos..tag_end]);
            result.push_str(&replaced);

            pos = close_start;
        } else {
            result.push_str(&xml[pos..]);
            break;
        }
    }

    (result, count)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn replace_in_wt_simple() {
        let xml = r#"<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>"#;
        let (result, count) = replace_in_wt_elements(xml, "World", "Rust");
        assert_eq!(count, 1);
        assert!(result.contains("<w:t>Hello Rust</w:t>"));
    }

    #[test]
    fn replace_in_wt_multiple() {
        let xml = r#"<w:r><w:t>foo bar foo</w:t></w:r>"#;
        let (result, count) = replace_in_wt_elements(xml, "foo", "baz");
        assert_eq!(count, 2);
        assert!(result.contains("<w:t>baz bar baz</w:t>"));
    }

    #[test]
    fn replace_preserves_attributes() {
        let xml = r#"<w:r><w:t xml:space="preserve"> Hello </w:t></w:r>"#;
        let (result, count) = replace_in_wt_elements(xml, "Hello", "World");
        assert_eq!(count, 1);
        assert!(result.contains(r#"xml:space="preserve"> World </w:t>"#));
    }

    #[test]
    fn no_match_returns_zero() {
        let xml = r#"<w:r><w:t>Hello</w:t></w:r>"#;
        let (result, count) = replace_in_wt_elements(xml, "xyz", "abc");
        assert_eq!(count, 0);
        assert_eq!(result, xml);
    }
}
