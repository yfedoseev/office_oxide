//! Pure Rust reader for Compound Binary File Format (CFBF/OLE2) containers.
//!
//! This crate provides read access to the OLE2 structured storage format used by
//! legacy Microsoft Office files (.doc, .xls, .ppt) and other applications.
//!
//! # Example
//!
//! ```no_run
//! use std::fs::File;
//! use office_oxide::cfb::CfbReader;
//!
//! let file = File::open("spreadsheet.xls").unwrap();
//! let mut reader = CfbReader::new(file).unwrap();
//! let workbook = reader.open_stream("Workbook").unwrap();
//! ```

pub mod blip;
mod directory;
mod error;
mod header;
mod reader;

pub use blip::{BlipFormat, BlipImage, extract_blip_images};
pub use directory::{DirEntry, EntryType};
pub use error::{CfbError, Result};
pub use header::{
    CFB_SIGNATURE, CfbHeader, DIFAT_SECT, END_OF_CHAIN, FAT_SECT, FREE_SECT, MAX_REG_SECT,
};
pub use reader::CfbReader;
