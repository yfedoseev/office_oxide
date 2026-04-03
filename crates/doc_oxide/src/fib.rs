//! File Information Block (FIB) parsing.
//!
//! The FIB is at the start of the WordDocument stream. It contains metadata
//! and pointers to other structures in the Table stream.

use crate::error::{DocError, Result};

/// Parsed FIB fields needed for text extraction.
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct Fib {
    /// Word version identifier.
    pub version: u16,
    /// Which table stream to use: true = "1Table", false = "0Table".
    pub use_table1: bool,
    /// Offset of the CLX (piece table) in the Table stream.
    pub clx_offset: u32,
    /// Size of the CLX in the Table stream.
    pub clx_size: u32,
    /// Total length of text in the main document (in characters).
    pub text_len: u32,
    /// Length of footnote text.
    pub footnote_len: u32,
    /// Length of header/footer text.
    pub header_len: u32,
    /// Length of comment text.
    pub comment_len: u32,
    /// Length of endnote text.
    pub endnote_len: u32,
    /// Length of textbox text.
    pub textbox_len: u32,
    /// Length of header textbox text.
    pub header_textbox_len: u32,
}

impl Fib {
    /// Parse the FIB from the WordDocument stream.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 68 {
            return Err(DocError::InvalidFib(format!(
                "WordDocument stream too short: {} bytes",
                data.len()
            )));
        }

        let wident = u16::from_le_bytes([data[0], data[1]]);
        // 0xA5EC = Word 97 and later. 0xA5DC = Word 6/95. Others may appear.
        if wident != 0xA5EC && wident != 0xA5DC {
            return Err(DocError::InvalidFib(format!(
                "unknown wIdent: 0x{wident:04X}"
            )));
        }

        let version = u16::from_le_bytes([data[2], data[3]]);

        // Flags at offset 0x0A (u16): bit 9 = fWhichTblStm
        let flags = u16::from_le_bytes([data[0x0A], data[0x0B]]);
        let use_table1 = (flags & (1 << 9)) != 0;

        // FibRgLw97 starts at offset 0x22, its size field at 0x22 (u16, should be 0x16).
        // Text lengths in FibRgLw97:
        // ccpText at 0x4C (offset from FIB start)
        let text_len = read_u32(data, 0x4C);
        let footnote_len = read_u32(data, 0x50);
        let header_len = read_u32(data, 0x54);
        let comment_len = read_u32(data, 0x58);
        let endnote_len = read_u32(data, 0x5C);
        let textbox_len = read_u32(data, 0x60);
        let header_textbox_len = read_u32(data, 0x64);

        // FibRgFcLcb starts at variable offset. For Word 97 (nFib=0x00C1),
        // the CLX is stored as fcClx/lcbClx.
        //
        // In practice, the FibRgFcLcb97 starts at offset 0x9A.
        // fcClx is at FibRgFcLcb97 offset 0x01A2 (relative to 0x9A).
        // = absolute offset 0x9A + 0x01A2 = 0x023C
        // lcbClx is at 0x9A + 0x01A6 = 0x0240
        //
        // But these offsets vary by nFib. Let's use the standard approach:
        // FibRgFcLcb97 starts at 0x9A.
        // fcClx = FibRgFcLcb[0x1A2..0x1A6] relative = absolute 0x23C
        // lcbClx = FibRgFcLcb[0x1A6..0x1AA] relative = absolute 0x240

        // CLX offset varies by FIB version. Try standard Word 97+ location first.
        let (clx_offset, clx_size) = if data.len() > 0x01AA {
            (read_u32(data, 0x01A2), read_u32(data, 0x01A6))
        } else {
            // FIB too short for standard CLX location — return zeros.
            (0, 0)
        };

        Ok(Self {
            version,
            use_table1,
            clx_offset,
            clx_size,
            text_len,
            footnote_len,
            header_len,
            comment_len,
            endnote_len,
            textbox_len,
            header_textbox_len,
        })
    }
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    if offset + 4 <= data.len() {
        u32::from_le_bytes([
            data[offset],
            data[offset + 1],
            data[offset + 2],
            data[offset + 3],
        ])
    } else {
        0
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn build_minimal_fib() -> Vec<u8> {
        let mut data = vec![0u8; 1024];
        // wIdent = Word 97
        data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        // nFib (version)
        data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        // flags: use 1Table (bit 9)
        data[0x0A..0x0C].copy_from_slice(&(1u16 << 9).to_le_bytes());
        // ccpText = 100
        data[0x4C..0x50].copy_from_slice(&100u32.to_le_bytes());
        // fcClx
        data[0x01A2..0x01A6].copy_from_slice(&512u32.to_le_bytes());
        // lcbClx
        data[0x01A6..0x01AA].copy_from_slice(&64u32.to_le_bytes());
        data
    }

    #[test]
    fn parse_valid_fib() {
        let data = build_minimal_fib();
        let fib = Fib::parse(&data).unwrap();
        assert_eq!(fib.version, 0x00C1);
        assert!(fib.use_table1);
        assert_eq!(fib.text_len, 100);
        assert_eq!(fib.clx_offset, 512);
        assert_eq!(fib.clx_size, 64);
    }

    #[test]
    fn bad_wident_rejected() {
        let mut data = build_minimal_fib();
        data[0..2].copy_from_slice(&0x1234u16.to_le_bytes());
        assert!(Fib::parse(&data).is_err());
    }

    #[test]
    fn too_short_rejected() {
        let data = vec![0u8; 100];
        assert!(Fib::parse(&data).is_err());
    }

    #[test]
    fn use_table0() {
        let mut data = build_minimal_fib();
        data[0x0A..0x0C].copy_from_slice(&0u16.to_le_bytes()); // clear bit 9
        let fib = Fib::parse(&data).unwrap();
        assert!(!fib.use_table1);
    }
}
