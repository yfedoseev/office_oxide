//! Common traits for all Office document format crates.

/// A parsed Office document that supports text extraction.
///
/// All format crates (`docx_oxide`, `xlsx_oxide`, `pptx_oxide`,
/// `doc_oxide`, `xls_oxide`, `ppt_oxide`) implement this trait
/// on their main document type.
pub trait OfficeDocument {
    /// Extract plain text from the document.
    fn plain_text(&self) -> String;

    /// Convert the document to a Markdown representation.
    fn to_markdown(&self) -> String;
}
