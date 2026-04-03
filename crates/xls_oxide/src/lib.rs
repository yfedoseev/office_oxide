//! Pure Rust reader for legacy Excel Binary (.xls) BIFF8 files.
//!
//! # Example
//!
//! ```no_run
//! use xls_oxide::XlsDocument;
//!
//! let doc = XlsDocument::open("spreadsheet.xls").unwrap();
//! println!("{}", doc.plain_text());
//! ```

mod cell;
mod error;
pub mod images;
mod records;
mod sst;
mod workbook;

pub use cell::{Cell, CellValue};
pub use error::{XlsError, Result};
pub use images::{ImageFormat, XlsImage};
pub use office_core::OfficeDocument;
pub use workbook::{Sheet, XlsDocument};
