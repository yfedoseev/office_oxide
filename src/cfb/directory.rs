use super::error::{CfbError, Result};

/// Type of a directory entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EntryType {
    /// Empty / unused slot.
    Empty,
    /// Storage (like a directory).
    Storage,
    /// Stream (like a file).
    Stream,
    /// Root storage (exactly one per file).
    RootStorage,
}

/// A parsed directory entry (128 bytes each in the CFB).
#[derive(Debug, Clone)]
pub struct DirEntry {
    /// Entry name decoded from UTF-16LE.
    pub name: String,
    /// Entry type.
    pub entry_type: EntryType,
    /// Color in the red-black tree (0 = red, 1 = black).
    pub color: u8,
    /// Left sibling directory ID (0xFFFFFFFF = none).
    pub left_sibling: u32,
    /// Right sibling directory ID.
    pub right_sibling: u32,
    /// Child directory ID (for storages).
    pub child: u32,
    /// Starting sector of this entry's data.
    pub start_sector: u32,
    /// Size of the stream in bytes.
    pub stream_size: u64,
}

/// Sentinel: no sibling/child.
pub const NO_ENTRY: u32 = 0xFFFFFFFF;

impl DirEntry {
    /// Parse a single 128-byte directory entry.
    pub fn parse(buf: &[u8], major_version: u16) -> Result<Self> {
        if buf.len() < 128 {
            return Err(CfbError::InvalidDirectory("entry too short".into()));
        }

        // Name: UTF-16LE in first 64 bytes. Name size (in bytes, including null) at offset 0x40.
        let name_size = u16::from_le_bytes([buf[0x40], buf[0x41]]) as usize;
        let name = if name_size >= 2 {
            // Exclude the null terminator (2 bytes). Clamp to max 64 bytes (the name field size).
            let name_bytes = (name_size - 2).min(62);
            let chars: Vec<u16> = (0..name_bytes / 2)
                .map(|i| u16::from_le_bytes([buf[i * 2], buf[i * 2 + 1]]))
                .collect();
            String::from_utf16_lossy(&chars)
        } else {
            String::new()
        };

        let entry_type = match buf[0x42] {
            0 | 3 | 4 => EntryType::Empty, // 0=unknown, 3/4=reserved, treat as empty
            1 => EntryType::Storage,
            2 => EntryType::Stream,
            5 => EntryType::RootStorage,
            _ => EntryType::Empty, // tolerate unknown types
        };

        let color = buf[0x43];
        let left_sibling = u32::from_le_bytes([buf[0x44], buf[0x45], buf[0x46], buf[0x47]]);
        let right_sibling = u32::from_le_bytes([buf[0x48], buf[0x49], buf[0x4A], buf[0x4B]]);
        let child = u32::from_le_bytes([buf[0x4C], buf[0x4D], buf[0x4E], buf[0x4F]]);

        let start_sector = u32::from_le_bytes([buf[0x74], buf[0x75], buf[0x76], buf[0x77]]);

        // Stream size: v3 uses only low 32 bits, v4 uses full 64 bits.
        let size_low = u32::from_le_bytes([buf[0x78], buf[0x79], buf[0x7A], buf[0x7B]]) as u64;
        let stream_size = if major_version == 4 {
            let size_high =
                u32::from_le_bytes([buf[0x7C], buf[0x7D], buf[0x7E], buf[0x7F]]) as u64;
            (size_high << 32) | size_low
        } else {
            size_low
        };

        Ok(Self {
            name,
            entry_type,
            color,
            left_sibling,
            right_sibling,
            child,
            start_sector,
            stream_size,
        })
    }
}

/// Parse all directory entries from a concatenated directory stream.
pub fn parse_directory(data: &[u8], major_version: u16) -> Result<Vec<DirEntry>> {
    let count = data.len() / 128;
    let mut entries = Vec::with_capacity(count);
    for i in 0..count {
        let entry = DirEntry::parse(&data[i * 128..(i + 1) * 128], major_version)?;
        entries.push(entry);
    }
    Ok(entries)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn build_entry(name: &str, entry_type: u8) -> Vec<u8> {
        let mut buf = vec![0u8; 128];
        // Encode name as UTF-16LE with null terminator.
        let utf16: Vec<u16> = name.encode_utf16().collect();
        for (i, &ch) in utf16.iter().enumerate() {
            let bytes = ch.to_le_bytes();
            buf[i * 2] = bytes[0];
            buf[i * 2 + 1] = bytes[1];
        }
        // Name size in bytes (including null terminator).
        let name_size = ((utf16.len() + 1) * 2) as u16;
        buf[0x40..0x42].copy_from_slice(&name_size.to_le_bytes());
        buf[0x42] = entry_type;
        buf[0x43] = 1; // black
        // No siblings/children.
        buf[0x44..0x48].copy_from_slice(&super::NO_ENTRY.to_le_bytes());
        buf[0x48..0x4C].copy_from_slice(&super::NO_ENTRY.to_le_bytes());
        buf[0x4C..0x50].copy_from_slice(&super::NO_ENTRY.to_le_bytes());
        buf
    }

    #[test]
    fn parse_root_entry() {
        let mut buf = build_entry("Root Entry", 5);
        // start sector = 0, size = 1024
        buf[0x74..0x78].copy_from_slice(&0u32.to_le_bytes());
        buf[0x78..0x7C].copy_from_slice(&1024u32.to_le_bytes());
        let entry = DirEntry::parse(&buf, 3).unwrap();
        assert_eq!(entry.name, "Root Entry");
        assert_eq!(entry.entry_type, EntryType::RootStorage);
        assert_eq!(entry.stream_size, 1024);
    }

    #[test]
    fn parse_stream_entry() {
        let mut buf = build_entry("Workbook", 2);
        buf[0x74..0x78].copy_from_slice(&5u32.to_le_bytes());
        buf[0x78..0x7C].copy_from_slice(&8192u32.to_le_bytes());
        let entry = DirEntry::parse(&buf, 3).unwrap();
        assert_eq!(entry.name, "Workbook");
        assert_eq!(entry.entry_type, EntryType::Stream);
        assert_eq!(entry.start_sector, 5);
        assert_eq!(entry.stream_size, 8192);
    }

    #[test]
    fn parse_empty_entry() {
        let buf = build_entry("", 0);
        let entry = DirEntry::parse(&buf, 3).unwrap();
        assert_eq!(entry.entry_type, EntryType::Empty);
    }

    #[test]
    fn v4_stream_size_64bit() {
        let mut buf = build_entry("BigStream", 2);
        // 5 GB stream: low = 0x40000000, high = 0x01
        buf[0x78..0x7C].copy_from_slice(&0x40000000u32.to_le_bytes());
        buf[0x7C..0x80].copy_from_slice(&0x01u32.to_le_bytes());
        let entry = DirEntry::parse(&buf, 4).unwrap();
        assert_eq!(entry.stream_size, 0x0000_0001_4000_0000);
    }

    #[test]
    fn parse_multiple_entries() {
        let mut data = build_entry("Root Entry", 5);
        data.extend_from_slice(&build_entry("Workbook", 2));
        data.extend_from_slice(&build_entry("", 0));
        let entries = parse_directory(&data, 3).unwrap();
        assert_eq!(entries.len(), 3);
        assert_eq!(entries[0].entry_type, EntryType::RootStorage);
        assert_eq!(entries[1].name, "Workbook");
        assert_eq!(entries[2].entry_type, EntryType::Empty);
    }
}
