//! Pure Rust reader for legacy PowerPoint Binary (.ppt) files.
//!
//! # Example
//!
//! ```no_run
//! use ppt_oxide::PptDocument;
//!
//! let doc = PptDocument::open("presentation.ppt").unwrap();
//! println!("{}", doc.plain_text());
//! ```

mod document;
mod error;
pub mod images;
mod records;
mod text;

pub use document::PptDocument;
pub use error::{PptError, Result};
pub use images::{ImageFormat, PptImage};
pub use office_core::OfficeDocument;
pub use text::{SlideText, TextRun, TextType};
