//! Unified document editing API.
//!
//! Provides a read-modify-write workflow that preserves all unmodified
//! parts (images, charts, custom XML, etc.) through the `EditablePackage`.

use std::io::{Read, Seek, Write};
use std::path::Path;

use crate::format::DocumentFormat;
use crate::Result;

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
                Ok(Self { inner: EditableInner::Docx(doc) })
            }
            DocumentFormat::Xlsx => {
                let doc = crate::xlsx::edit::EditableXlsx::open(path)?;
                Ok(Self { inner: EditableInner::Xlsx(doc) })
            }
            DocumentFormat::Pptx => {
                let doc = crate::pptx::edit::EditablePptx::open(path)?;
                Ok(Self { inner: EditableInner::Pptx(doc) })
            }
            _ => Err(crate::OfficeError::UnsupportedFormat(format!("{format:?}"))),
        }
    }

    /// Open a document for editing from a reader with explicit format.
    pub fn from_reader<R: Read + Seek>(reader: R, format: DocumentFormat) -> Result<Self> {
        match format {
            DocumentFormat::Docx => {
                let doc = crate::docx::edit::EditableDocx::from_reader(reader)?;
                Ok(Self { inner: EditableInner::Docx(doc) })
            }
            DocumentFormat::Xlsx => {
                let doc = crate::xlsx::edit::EditableXlsx::from_reader(reader)?;
                Ok(Self { inner: EditableInner::Xlsx(doc) })
            }
            DocumentFormat::Pptx => {
                let doc = crate::pptx::edit::EditablePptx::from_reader(reader)?;
                Ok(Self { inner: EditableInner::Pptx(doc) })
            }
            _ => Err(crate::OfficeError::UnsupportedFormat(format!("{format:?}"))),
        }
    }

    /// Replace all occurrences of `find` with `replace` in text content.
    /// Returns the number of replacements made.
    ///
    /// For DOCX: replaces text in `<w:t>` elements.
    /// For PPTX: replaces text in `<a:t>` elements across all slides.
    /// For XLSX: not applicable (use `set_cell` instead) — returns 0.
    pub fn replace_text(&mut self, find: &str, replace: &str) -> usize {
        match &mut self.inner {
            EditableInner::Docx(doc) => doc.replace_text(find, replace),
            EditableInner::Pptx(doc) => doc.replace_text(find, replace),
            EditableInner::Xlsx(_) => 0,
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
            }
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
