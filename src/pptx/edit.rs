//! PPTX editing via raw XML text replacement.
//!
//! Uses the `EditablePackage` from core to preserve all parts,
//! replacing text in slide XML `<a:t>` elements.

use crate::core::editable::EditablePackage;
use crate::core::opc::PartName;

use super::Result;

/// Content types (ECMA-376 Part 1 §13.3) of the PresentationML parts whose
/// `<a:t>` text `replace_text` rewrites.
const TEXT_PART_CONTENT_TYPES: &[&str] = &[
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
    "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
    "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
];

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

    /// Replace all occurrences of `find` with `replace` in the text of every
    /// slide, notes slide, slide layout and slide master part.
    /// Returns the total number of replacements made.
    ///
    /// Parts are found by content type, not by guessing names: slide part
    /// numbering has gaps after a deletion, and a deck may have any number
    /// of slides.
    pub fn replace_text(&mut self, find: &str, replace: &str) -> usize {
        let mut total = 0;

        let mut targets: Vec<PartName> = self
            .package
            .content_types()
            .overrides()
            .iter()
            .filter(|(_, ct)| TEXT_PART_CONTENT_TYPES.contains(&ct.as_str()))
            .map(|(pn, _)| pn.clone())
            .collect();
        targets.sort_by(|a, b| a.as_str().cmp(b.as_str()));

        for part_name in targets {
            let Some(data) = self.package.get_part(&part_name) else {
                continue;
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

    /// A package with slide parts numbered with a gap (slide1, slide3 — the
    /// shape a deck has after a slide is deleted), plus a notes slide, a
    /// layout and a master, each carrying the search text once.
    fn gapped_deck() -> Vec<u8> {
        use crate::core::opc::OpcWriter;
        use crate::core::relationships::rel_types;
        const NS: &str = r#"xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#;
        let tree = |root: &str| {
            format!(
                r#"<?xml version="1.0"?><p:{root} {NS}><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>foo</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:{root}>"#
            )
        };
        let mut w = OpcWriter::new(std::io::Cursor::new(Vec::new())).unwrap();
        let pres = PartName::new("/ppt/presentation.xml").unwrap();
        w.add_package_rel(rel_types::OFFICE_DOCUMENT, "ppt/presentation.xml");
        w.add_part(
            &pres,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            format!(r#"<?xml version="1.0"?><p:presentation {NS}/>"#).as_bytes(),
        )
        .unwrap();
        let ct = "application/vnd.openxmlformats-officedocument.presentationml.";
        for (name, kind, root) in [
            ("/ppt/slides/slide1.xml", "slide+xml", "sld"),
            ("/ppt/slides/slide3.xml", "slide+xml", "sld"),
            ("/ppt/notesSlides/notesSlide3.xml", "notesSlide+xml", "notes"),
            ("/ppt/slideLayouts/slideLayout1.xml", "slideLayout+xml", "sldLayout"),
            ("/ppt/slideMasters/slideMaster1.xml", "slideMaster+xml", "sldMaster"),
        ] {
            let part = PartName::new(name).unwrap();
            w.add_part(&part, &format!("{ct}{kind}"), tree(root).as_bytes())
                .unwrap();
        }
        w.finish().unwrap().into_inner()
    }

    /// Slide numbering has gaps after a deletion; a positional
    /// `slide1..=slideN` walk stopped at the first missing number, so every
    /// slide after the gap — and every notes slide, layout and master — was
    /// silently left unchanged while the caller got a partial count.
    #[test]
    fn test_replace_text_reaches_every_slide_notes_layout_and_master_part() {
        let mut ed = EditablePptx::from_reader(std::io::Cursor::new(gapped_deck())).unwrap();
        assert_eq!(ed.replace_text("foo", "bar"), 5);
        for name in [
            "/ppt/slides/slide1.xml",
            "/ppt/slides/slide3.xml",
            "/ppt/notesSlides/notesSlide3.xml",
            "/ppt/slideLayouts/slideLayout1.xml",
            "/ppt/slideMasters/slideMaster1.xml",
        ] {
            let data = ed.package.get_part(&PartName::new(name).unwrap()).unwrap();
            let s = String::from_utf8_lossy(data);
            assert!(s.contains("<a:t>bar</a:t>"), "{name} not rewritten: {s}");
        }
    }

    #[test]
    fn test_replace_in_at_simple() {
        let xml = r#"<a:p><a:r><a:t>Hello World</a:t></a:r></a:p>"#;
        let (result, count) = replace_in_at_elements(xml, "World", "PPTX");
        assert_eq!(count, 1);
        assert!(result.contains("<a:t>Hello PPTX</a:t>"));
    }

    #[test]
    fn test_replace_in_at_multiple_runs() {
        let xml = r#"<a:r><a:t>foo</a:t></a:r><a:r><a:t>foo</a:t></a:r>"#;
        let (result, count) = replace_in_at_elements(xml, "foo", "bar");
        assert_eq!(count, 2);
        assert_eq!(result.matches("bar").count(), 2);
    }

    #[test]
    fn test_no_match_returns_zero() {
        let xml = r#"<a:r><a:t>Hello</a:t></a:r>"#;
        let (result, count) = replace_in_at_elements(xml, "xyz", "abc");
        assert_eq!(count, 0);
        assert_eq!(result, xml);
    }
}
