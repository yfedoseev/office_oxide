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

/// `true` when `reader` begins with the CFB/OLE2 magic signature — the
/// wrapper Office uses for a password-protected (encrypted) OOXML
/// package, which is not a ZIP at all. Leaves the reader rewound to the
/// start.
///
/// Shared by `Document::from_reader` and each OOXML format's own
/// `from_reader` (`DocxDocument`/`XlsxDocument`/`PptxDocument`) so a
/// caller using the format-specific reader directly gets the same
/// friendly "password-protected" error instead of a confusing low-level
/// zip error ("Could not find EOCD") that says nothing about the real
/// cause (issue #232).
pub fn is_cfb_container<R: std::io::Read + std::io::Seek>(
    reader: &mut R,
) -> std::io::Result<bool> {
    use std::io::SeekFrom;
    let mut magic = [0u8; 8];
    reader.seek(SeekFrom::Start(0))?;
    let mut filled = 0;
    while filled < magic.len() {
        match reader.read(&mut magic[filled..]) {
            Ok(0) => break,
            Ok(n) => filled += n,
            Err(e) if e.kind() == std::io::ErrorKind::Interrupted => continue,
            Err(e) => return Err(e),
        }
    }
    reader.seek(SeekFrom::Start(0))?;
    Ok(filled == 8 && magic == CFB_SIGNATURE)
}
