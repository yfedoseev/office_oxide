//! Unified document editing API.
//!
//! Provides a read-modify-write workflow that preserves all unmodified
//! parts (images, charts, custom XML, etc.) through the `EditablePackage`.

use std::io::{Read, Seek, Write};
use std::path::Path;

use crate::Result;
use crate::format::DocumentFormat;

/// An editable document that supports text replacement and saving.
///
/// The document is loaded in its entirety (all OPC parts) so that
/// unmodified parts are preserved verbatim on save.
pub struct EditableDocument {
    inner: EditableInner,
}

enum EditableInner {
    Docx(crate::docx::edit::EditableDocx),
    Xlsx(crate::xlsx::edit::EditableXlsx),
    Pptx(crate::pptx::edit::EditablePptx),
}

impl EditableDocument {
    /// Open a document for editing. Format is detected from the file extension.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let path = path.as_ref();
        let format = DocumentFormat::from_path(path).ok_or_else(|| {
            crate::OfficeError::UnsupportedFormat(
                path.extension()
                    .and_then(|e| e.to_str())
                    .unwrap_or("(none)")
                    .to_string(),
            )
        })?;
        match format {
            DocumentFormat::Docx => {
                let doc = crate::docx::edit::EditableDocx::open(path)?;
                Ok(Self {
                    inner: EditableInner::Docx(doc),
                })
            },
            DocumentFormat::Xlsx => {
                let doc = crate::xlsx::edit::EditableXlsx::open(path)?;
                Ok(Self {
                    inner: EditableInner::Xlsx(doc),
                })
            },
            DocumentFormat::Pptx => {
                let doc = crate::pptx::edit::EditablePptx::open(path)?;
                Ok(Self {
                    inner: EditableInner::Pptx(doc),
                })
            },
            _ => Err(crate::OfficeError::UnsupportedFormat(format!("{format:?}"))),
        }
    }

    /// Open a document for editing from a reader with explicit format.
    pub fn from_reader<R: Read + Seek>(reader: R, format: DocumentFormat) -> Result<Self> {
        match format {
            DocumentFormat::Docx => {
                let doc = crate::docx::edit::EditableDocx::from_reader(reader)?;
                Ok(Self {
                    inner: EditableInner::Docx(doc),
                })
            },
            DocumentFormat::Xlsx => {
                let doc = crate::xlsx::edit::EditableXlsx::from_reader(reader)?;
                Ok(Self {
                    inner: EditableInner::Xlsx(doc),
                })
            },
            DocumentFormat::Pptx => {
                let doc = crate::pptx::edit::EditablePptx::from_reader(reader)?;
                Ok(Self {
                    inner: EditableInner::Pptx(doc),
                })
            },
            _ => Err(crate::OfficeError::UnsupportedFormat(format!("{format:?}"))),
        }
    }

    /// Replace all occurrences of `find` with `replace` in text content.
    /// Returns the number of replacements made.
    ///
    /// For DOCX: replaces text in `<w:t>` elements.
    /// For PPTX: replaces text in `<a:t>` elements across all slides.
    ///
    /// # Errors
    ///
    /// Returns [`OfficeError::UnsupportedFormat`] for XLSX, which has no text
    /// replacement — use [`Self::set_cell`]. Returning `0` instead was
    /// indistinguishable from "the text was not present", so a caller (and
    /// every agent driving the CLI or MCP server) was told the edit
    /// succeeded when the operation is not implemented at all.
    pub fn replace_text(&mut self, find: &str, replace: &str) -> Result<usize> {
        match &mut self.inner {
            EditableInner::Docx(doc) => Ok(doc.replace_text(find, replace)),
            EditableInner::Pptx(doc) => Ok(doc.replace_text(find, replace)),
            EditableInner::Xlsx(_) => Err(crate::OfficeError::UnsupportedFormat(
                "replace_text is not supported for XLSX; use set_cell".to_string(),
            )),
        }
    }

    /// Set a cell value in an XLSX document.
    /// Returns an error if this is not an XLSX document.
    pub fn set_cell(
        &mut self,
        sheet_index: usize,
        cell_ref: &str,
        value: crate::xlsx::edit::CellValue,
    ) -> Result<()> {
        match &mut self.inner {
            EditableInner::Xlsx(doc) => {
                doc.set_cell(sheet_index, cell_ref, value)?;
                Ok(())
            },
            _ => Err(crate::OfficeError::UnsupportedFormat(
                "set_cell is only supported for XLSX".to_string(),
            )),
        }
    }

    /// Save the edited document to a file.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        match &self.inner {
            EditableInner::Docx(doc) => doc.save(path)?,
            EditableInner::Xlsx(doc) => doc.save(path)?,
            EditableInner::Pptx(doc) => doc.save(path)?,
        }
        Ok(())
    }

    /// Write the edited document to any `Write + Seek` destination.
    pub fn write_to<W: Write + Seek>(&self, writer: W) -> Result<()> {
        match &self.inner {
            EditableInner::Docx(doc) => doc.write_to(writer)?,
            EditableInner::Xlsx(doc) => doc.write_to(writer)?,
            EditableInner::Pptx(doc) => doc.write_to(writer)?,
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use super::*;

    fn make_docx_bytes() -> Vec<u8> {
        let mut w = crate::docx::write::DocxWriter::new();
        w.add_paragraph("Hello world");
        let mut buf = Cursor::new(Vec::new());
        w.write_to(&mut buf).unwrap();
        buf.into_inner()
    }

    fn make_xlsx_bytes() -> Vec<u8> {
        let mut w = crate::xlsx::write::XlsxWriter::new();
        {
            let mut sheet = w.add_sheet("Sheet1");
            sheet.set_cell(0, 0, crate::xlsx::write::CellData::String("value".into()));
        }
        let mut buf = Cursor::new(Vec::new());
        w.write_to(&mut buf).unwrap();
        buf.into_inner()
    }

    fn make_pptx_bytes() -> Vec<u8> {
        let mut w = crate::pptx::write::PptxWriter::new();
        let slide = w.add_slide();
        slide.set_title("Slide 1");
        let mut buf = Cursor::new(Vec::new());
        w.write_to(&mut buf).unwrap();
        buf.into_inner()
    }

    #[test]
    fn docx_replace_text_roundtrip() {
        let data = make_docx_bytes();
        let mut doc =
            EditableDocument::from_reader(Cursor::new(data), DocumentFormat::Docx).unwrap();
        let count = doc
            .replace_text("Hello", "Hi")
            .expect("docx supports replace");
        assert!(count >= 1);
        let mut out = Cursor::new(Vec::new());
        doc.write_to(&mut out).unwrap();
        assert!(!out.into_inner().is_empty());
    }

    #[test]
    fn xlsx_set_cell_roundtrip() {
        let data = make_xlsx_bytes();
        let mut doc =
            EditableDocument::from_reader(Cursor::new(data), DocumentFormat::Xlsx).unwrap();
        doc.set_cell(0, "A1", crate::xlsx::edit::CellValue::String("updated".into()))
            .unwrap();
        let mut out = Cursor::new(Vec::new());
        doc.write_to(&mut out).unwrap();
        assert!(!out.into_inner().is_empty());
    }

    #[test]
    fn pptx_replace_text_roundtrip() {
        let data = make_pptx_bytes();
        let mut doc =
            EditableDocument::from_reader(Cursor::new(data), DocumentFormat::Pptx).unwrap();
        doc.replace_text("Slide 1", "Updated")
            .expect("pptx supports replace");
        let mut out = Cursor::new(Vec::new());
        doc.write_to(&mut out).unwrap();
        assert!(!out.into_inner().is_empty());
    }

    #[test]
    fn xlsx_replace_text_is_an_error_not_a_silent_zero() {
        // Returning 0 was indistinguishable from "the text was not present",
        // so the CLI and MCP server reported success and rewrote the file for
        // an operation that cannot work at all.
        let data = make_xlsx_bytes();
        let mut doc =
            EditableDocument::from_reader(Cursor::new(data), DocumentFormat::Xlsx).unwrap();
        let err = doc
            .replace_text("anything", "other")
            .expect_err("XLSX has no text replacement and must say so");
        let msg = err.to_string().to_lowercase();
        assert!(
            msg.contains("xlsx") || msg.contains("unsupported"),
            "the error must name the unsupported format: {err}"
        );
    }

    #[test]
    fn set_cell_on_docx_returns_error() {
        let data = make_docx_bytes();
        let mut doc =
            EditableDocument::from_reader(Cursor::new(data), DocumentFormat::Docx).unwrap();
        assert!(
            doc.set_cell(0, "A1", crate::xlsx::edit::CellValue::String("x".into()))
                .is_err()
        );
    }

    #[test]
    fn from_reader_unsupported_format_returns_error() {
        let data = vec![0u8; 16];
        assert!(EditableDocument::from_reader(Cursor::new(data), DocumentFormat::Doc).is_err());
    }
}
