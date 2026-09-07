//! PPTX editing via raw XML text replacement.
//!
//! Uses the `EditablePackage` from core to preserve all parts,
//! replacing text in slide XML `<a:t>` elements.

use crate::core::editable::EditablePackage;
use crate::core::opc::PartName;

use super::Result;

/// An editable PPTX document that supports text replacement and saving.
pub struct EditablePptx {
    package: EditablePackage,
}

impl EditablePptx {
    /// Open a PPTX file for editing.
    pub fn open(path: impl AsRef<std::path::Path>) -> Result<Self> {
        let package = EditablePackage::open(&path)?;
        Ok(Self { package })
    }

    /// Open from any `Read + Seek` source.
    pub fn from_reader<R: std::io::Read + std::io::Seek>(reader: R) -> Result<Self> {
        let package = EditablePackage::from_reader(reader)?;
        Ok(Self { package })
    }

    /// Replace all occurrences of `find` with `replace` across all slides.
    /// Returns the total number of replacements made.
    pub fn replace_text(&mut self, find: &str, replace: &str) -> usize {
        let mut total = 0;

        // Find all slide parts
        for i in 1..=100 {
            let part_name = match PartName::new(&format!("/ppt/slides/slide{i}.xml")) {
                Ok(pn) => pn,
                Err(_) => break,
            };
            let Some(data) = self.package.get_part(&part_name) else {
                break;
            };
            let xml_str = String::from_utf8_lossy(data);
            let (new_xml, count) = replace_in_at_elements(&xml_str, find, replace);
            if count > 0 {
                self.package.set_part(part_name, new_xml.into_bytes());
                total += count;
            }
        }

        total
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

/// Replace text within `<a:t>...</a:t>` elements in DrawingML XML.
fn replace_in_at_elements(xml: &str, find: &str, replace: &str) -> (String, usize) {
    crate::core::editable::replace_in_text_elements(xml, "a:t", find, replace)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn replace_in_at_simple() {
        let xml = r#"<a:p><a:r><a:t>Hello World</a:t></a:r></a:p>"#;
        let (result, count) = replace_in_at_elements(xml, "World", "PPTX");
        assert_eq!(count, 1);
        assert!(result.contains("<a:t>Hello PPTX</a:t>"));
    }

    #[test]
    fn replace_in_at_multiple_runs() {
        let xml = r#"<a:r><a:t>foo</a:t></a:r><a:r><a:t>foo</a:t></a:r>"#;
        let (result, count) = replace_in_at_elements(xml, "foo", "bar");
        assert_eq!(count, 2);
        assert_eq!(result.matches("bar").count(), 2);
    }

    #[test]
    fn no_match_returns_zero() {
        let xml = r#"<a:r><a:t>Hello</a:t></a:r>"#;
        let (result, count) = replace_in_at_elements(xml, "xyz", "abc");
        assert_eq!(count, 0);
        assert_eq!(result, xml);
    }
}
