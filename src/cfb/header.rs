use super::error::{CfbError, Result};

/// CFB file magic signature.
pub const CFB_SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

/// Special sector IDs.
pub const FREE_SECT: u32 = 0xFFFFFFFF;
pub const END_OF_CHAIN: u32 = 0xFFFFFFFE;
pub const FAT_SECT: u32 = 0xFFFFFFFD;
pub const DIFAT_SECT: u32 = 0xFFFFFFFC;
pub const MAX_REG_SECT: u32 = 0xFFFFFFFA;

/// Parsed CFB header (first 512 bytes).
#[derive(Debug, Clone)]
pub struct CfbHeader {
    /// Major version (3 or 4).
    pub major_version: u16,
    /// Minor version (typically 0x003E).
    pub minor_version: u16,
    /// Sector size in bytes (512 for v3, 4096 for v4).
    pub sector_size: usize,
    /// Mini-sector size in bytes (always 64).
    pub mini_sector_size: usize,
    /// Total number of FAT sectors.
    pub fat_sector_count: u32,
    /// First directory sector.
    pub first_dir_sector: u32,
    /// Mini-stream cutoff size (typically 4096).
    pub mini_stream_cutoff: u32,
    /// First mini-FAT sector.
    pub first_mini_fat_sector: u32,
    /// Total number of mini-FAT sectors.
    pub mini_fat_sector_count: u32,
    /// First DIFAT sector (0xFFFFFFFE if none).
    pub first_difat_sector: u32,
    /// Total number of DIFAT sectors.
    pub difat_sector_count: u32,
    /// The 109 DIFAT entries stored in the header.
    pub header_difat: Vec<u32>,
}

impl CfbHeader {
    /// Parse a CFB header from a 512-byte buffer.
    pub fn parse(buf: &[u8]) -> Result<Self> {
        if buf.len() < 512 {
            return Err(CfbError::InvalidHeader("buffer too short".into()));
        }
        // Validate signature
        if buf[0..8] != CFB_SIGNATURE {
            return Err(CfbError::InvalidHeader("bad magic signature".into()));
        }

        let minor_version = u16::from_le_bytes([buf[0x18], buf[0x19]]);
        let major_version = u16::from_le_bytes([buf[0x1A], buf[0x1B]]);

        if major_version != 3 && major_version != 4 {
            return Err(CfbError::InvalidHeader(format!(
                "unsupported major version: {major_version}"
            )));
        }

        // Byte order must be little-endian (0xFFFE).
        let byte_order = u16::from_le_bytes([buf[0x1C], buf[0x1D]]);
        if byte_order != 0xFFFE {
            return Err(CfbError::InvalidHeader(format!(
                "unsupported byte order: 0x{byte_order:04X}"
            )));
        }

        let sector_power = u16::from_le_bytes([buf[0x1E], buf[0x1F]]);
        let sector_size = 1usize << sector_power;

        // v3 must be 512, v4 must be 4096
        if major_version == 3 && sector_size != 512 {
            return Err(CfbError::InvalidHeader("v3 sector size must be 512".into()));
        }
        if major_version == 4 && sector_size != 4096 {
            return Err(CfbError::InvalidHeader("v4 sector size must be 4096".into()));
        }

        let mini_sector_power = u16::from_le_bytes([buf[0x20], buf[0x21]]);
        let mini_sector_size = 1usize << mini_sector_power;

        let fat_sector_count = u32::from_le_bytes([buf[0x2C], buf[0x2D], buf[0x2E], buf[0x2F]]);
        let first_dir_sector = u32::from_le_bytes([buf[0x30], buf[0x31], buf[0x32], buf[0x33]]);
        let mini_stream_cutoff = u32::from_le_bytes([buf[0x38], buf[0x39], buf[0x3A], buf[0x3B]]);
        let first_mini_fat_sector =
            u32::from_le_bytes([buf[0x3C], buf[0x3D], buf[0x3E], buf[0x3F]]);
        let mini_fat_sector_count =
            u32::from_le_bytes([buf[0x40], buf[0x41], buf[0x42], buf[0x43]]);
        let first_difat_sector = u32::from_le_bytes([buf[0x44], buf[0x45], buf[0x46], buf[0x47]]);
        let difat_sector_count = u32::from_le_bytes([buf[0x48], buf[0x49], buf[0x4A], buf[0x4B]]);

        // Read the 109 DIFAT entries from header (offsets 0x4C..0x200).
        let mut header_difat = Vec::with_capacity(109);
        for i in 0..109 {
            let off = 0x4C + i * 4;
            let val = u32::from_le_bytes([buf[off], buf[off + 1], buf[off + 2], buf[off + 3]]);
            header_difat.push(val);
        }

        Ok(Self {
            major_version,
            minor_version,
            sector_size,
            mini_sector_size,
            fat_sector_count,
            first_dir_sector,
            mini_stream_cutoff,
            first_mini_fat_sector,
            mini_fat_sector_count,
            first_difat_sector,
            difat_sector_count,
            header_difat,
        })
    }

    /// Byte offset of a given sector index in the file.
    #[inline]
    pub fn sector_offset(&self, sector: u32) -> u64 {
        // Sector 0 starts right after the 512-byte header.
        // For v3: offset = 512 + sector * 512
        // For v4: the header occupies the first 4096-byte sector, so offset = 4096 + sector * 4096
        //
        // General formula: (sector + 1) * sector_size works for v3 (512-byte sectors).
        // For v4, header is padded to 4096 bytes, so same formula works.
        (sector as u64 + 1) * self.sector_size as u64
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Build a minimal valid v3 CFB header.
    fn build_v3_header() -> Vec<u8> {
        let mut buf = vec![0u8; 512];
        // Signature
        buf[0..8].copy_from_slice(&CFB_SIGNATURE);
        // Minor version
        buf[0x18..0x1A].copy_from_slice(&0x003Eu16.to_le_bytes());
        // Major version = 3
        buf[0x1A..0x1C].copy_from_slice(&3u16.to_le_bytes());
        // Byte order = little-endian
        buf[0x1C..0x1E].copy_from_slice(&0xFFFEu16.to_le_bytes());
        // Sector size power = 9 (512)
        buf[0x1E..0x20].copy_from_slice(&9u16.to_le_bytes());
        // Mini sector size power = 6 (64)
        buf[0x20..0x22].copy_from_slice(&6u16.to_le_bytes());
        // FAT sector count = 1
        buf[0x2C..0x30].copy_from_slice(&1u32.to_le_bytes());
        // First directory sector = 0
        buf[0x30..0x34].copy_from_slice(&0u32.to_le_bytes());
        // Mini-stream cutoff = 4096
        buf[0x38..0x3C].copy_from_slice(&4096u32.to_le_bytes());
        // First mini-FAT sector = END_OF_CHAIN (none)
        buf[0x3C..0x40].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
        // DIFAT entries all FREE_SECT
        for i in 0..109 {
            let off = 0x4C + i * 4;
            buf[off..off + 4].copy_from_slice(&FREE_SECT.to_le_bytes());
        }
        // First DIFAT entry = sector 1 (where the FAT lives)
        buf[0x4C..0x50].copy_from_slice(&1u32.to_le_bytes());
        buf
    }

    #[test]
    fn parse_valid_v3_header() {
        let buf = build_v3_header();
        let header = CfbHeader::parse(&buf).unwrap();
        assert_eq!(header.major_version, 3);
        assert_eq!(header.sector_size, 512);
        assert_eq!(header.mini_sector_size, 64);
        assert_eq!(header.fat_sector_count, 1);
        assert_eq!(header.first_dir_sector, 0);
        assert_eq!(header.mini_stream_cutoff, 4096);
        assert_eq!(header.first_mini_fat_sector, END_OF_CHAIN);
    }

    #[test]
    fn bad_signature_rejected() {
        let mut buf = build_v3_header();
        buf[0] = 0x00;
        assert!(CfbHeader::parse(&buf).is_err());
    }

    #[test]
    fn bad_version_rejected() {
        let mut buf = build_v3_header();
        buf[0x1A..0x1C].copy_from_slice(&5u16.to_le_bytes());
        assert!(CfbHeader::parse(&buf).is_err());
    }

    #[test]
    fn bad_byte_order_rejected() {
        let mut buf = build_v3_header();
        buf[0x1C..0x1E].copy_from_slice(&0xFFFFu16.to_le_bytes());
        assert!(CfbHeader::parse(&buf).is_err());
    }

    #[test]
    fn sector_offset_v3() {
        let buf = build_v3_header();
        let header = CfbHeader::parse(&buf).unwrap();
        // Sector 0 starts at byte 512
        assert_eq!(header.sector_offset(0), 512);
        assert_eq!(header.sector_offset(1), 1024);
        assert_eq!(header.sector_offset(5), 3072);
    }

    #[test]
    fn too_short_buffer_rejected() {
        let buf = vec![0u8; 100];
        assert!(CfbHeader::parse(&buf).is_err());
    }
}
