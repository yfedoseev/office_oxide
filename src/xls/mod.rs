//! Pure Rust reader for legacy Excel Binary (.xls) BIFF8 files.
//!
//! # Example
//!
//! ```no_run
//! use office_oxide::xls::XlsDocument;
//!
//! let doc = XlsDocument::open("spreadsheet.xls").unwrap();
//! println!("{}", doc.plain_text());
//! ```

mod biff_old;
mod cell;
mod codepage;
pub mod comment;
pub(crate) mod condfmt;
mod data_validation;
mod error;
pub mod hyperlink;
pub mod images;
mod records;
mod sst;
mod workbook;

pub use crate::core::OfficeDocument;
pub use cell::{Cell, CellValue, decode_rk};
pub use comment::XlsComment;
pub use error::{Result, XlsError};
pub use hyperlink::XlsHyperlink;
pub use images::{ImageFormat, XlsImage};
pub use workbook::{Sheet, XlsDocument};
