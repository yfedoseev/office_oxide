//! Pure Rust reader for legacy Word Binary (.doc) files.
//!
//! # Example
//!
//! ```no_run
//! use office_oxide::doc::DocDocument;
//!
//! let doc = DocDocument::open("document.doc").unwrap();
//! println!("{}", doc.plain_text());
//! ```

mod codepage;
mod document;
mod error;
mod fib;
pub mod images;
mod list_format;
mod papx;
mod piece_table;
mod sprm;

pub use crate::core::OfficeDocument;
pub use document::{DocDocument, SubDocument, SubDocumentKind};
pub use error::{DocError, Result};
pub use images::{DocImage, ImageFormat};
pub(crate) use list_format::ListFormatting;
pub(crate) use papx::DocParagraph;
pub(crate) use piece_table::HyperlinkSpan;
pub(crate) use sprm::{TapCellInfo, TapInfo};
// `PapProps` and `ListLevel` are only needed by unit tests inside this
// crate, so their re-exports are test-gated to avoid an unused-import
// warning in non-test builds.
#[cfg(test)]
pub(crate) use list_format::ListLevel;
#[cfg(test)]
pub(crate) use sprm::PapProps;
