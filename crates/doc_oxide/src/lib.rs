//! Pure Rust reader for legacy Word Binary (.doc) files.
//!
//! # Example
//!
//! ```no_run
//! use doc_oxide::DocDocument;
//!
//! let doc = DocDocument::open("document.doc").unwrap();
//! println!("{}", doc.plain_text());
//! ```

mod document;
mod error;
mod fib;
pub mod images;
mod piece_table;

pub use document::DocDocument;
pub use error::{DocError, Result};
pub use images::{DocImage, ImageFormat};
pub use office_core::OfficeDocument;
