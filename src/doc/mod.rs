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

mod document;
mod error;
mod fib;
pub mod images;
mod papx;
mod piece_table;
mod sprm;
pub mod styles;

/// The deepest outline level MS-DOC stores: `Heading 1`–`Heading 9`. The
/// `StdfBase.sti` and a user-defined `Heading N` style name use this range,
/// so it lives here — at the `.doc` format root rather than in any one
/// submodule — and the style-sheet sites read this same bound.
///
/// Note `sprmPOutLvl` (0x2640) is **not** in this range: it encodes the level
/// zero-based, 0x00–0x08 for Heading 1–9, with 0x09 meaning body text.
pub const MAX_OUTLINE_LEVEL: u8 = 9;

pub use crate::core::OfficeDocument;
pub use document::{DocDocument, SubDocument, SubDocumentKind};
pub use error::{DocError, Result};
pub use images::{DocImage, ImageFormat};
pub(crate) use papx::DocParagraph;
pub(crate) use sprm::{TapCellInfo, TapInfo};
// `PapProps` is only needed by unit tests inside this crate, so the re-export
// is test-gated to avoid an unused-import warning in non-test builds.
#[cfg(test)]
pub(crate) use sprm::PapProps;
