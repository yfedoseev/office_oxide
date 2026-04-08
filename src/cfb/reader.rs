use std::io::{Read, Seek, SeekFrom};

use super::directory::{parse_directory, DirEntry, EntryType, NO_ENTRY};
use super::error::{CfbError, Result};
use super::header::{CfbHeader, MAX_REG_SECT};

/// A reader for Compound Binary File (OLE2/CFBF) containers.
///
/// Provides random access to streams within the file.
pub struct CfbReader<R> {
    reader: R,
    header: CfbHeader,
    /// The full FAT: maps each sector → next sector in chain.
    fat: Vec<u32>,
    /// The mini-FAT: maps each mini-sector → next mini-sector.
    mini_fat: Vec<u32>,
    /// Directory entries.
    entries: Vec<DirEntry>,
    /// The mini-stream data (read from the root entry's stream chain).
    mini_stream: Vec<u8>,
}

impl<R: Read + Seek> CfbReader<R> {
    /// Open and parse a CFB file.
    pub fn new(mut reader: R) -> Result<Self> {
        // Read header.
        let mut header_buf = [0u8; 512];
        reader.read_exact(&mut header_buf)?;
        let header = CfbHeader::parse(&header_buf)?;

        // Build the FAT.
        let fat = Self::read_fat(&mut reader, &header)?;

        // Read directory entries.
        let dir_data = Self::read_chain(&mut reader, &header, &fat, header.first_dir_sector)?;
        let entries = parse_directory(&dir_data, header.major_version)?;

        // Read mini-FAT.
        let mini_fat = if header.first_mini_fat_sector <= MAX_REG_SECT {
            let mini_fat_data =
                Self::read_chain(&mut reader, &header, &fat, header.first_mini_fat_sector)?;
            mini_fat_data
                .chunks_exact(4)
                .map(|c| u32::from_le_bytes([c[0], c[1], c[2], c[3]]))
                .collect()
        } else {
            Vec::new()
        };

        // Read mini-stream (data from root entry's stream chain).
        let mini_stream = if !entries.is_empty()
            && entries[0].entry_type == EntryType::RootStorage
            && entries[0].start_sector <= MAX_REG_SECT
        {
            Self::read_chain(&mut reader, &header, &fat, entries[0].start_sector)?
        } else {
            Vec::new()
        };

        Ok(Self {
            reader,
            header,
            fat,
            mini_fat,
            entries,
            mini_stream,
        })
    }

    /// Get all directory entries.
    pub fn entries(&self) -> &[DirEntry] {
        &self.entries
    }

    /// Get the header.
    pub fn header(&self) -> &CfbHeader {
        &self.header
    }

    /// Find a stream entry by name (case-insensitive).
    pub fn find_entry(&self, name: &str) -> Option<usize> {
        let lower = name.to_ascii_lowercase();
        self.entries.iter().position(|e| {
            e.entry_type == EntryType::Stream && e.name.to_ascii_lowercase() == lower
        })
    }

    /// Find an entry by path (e.g., "Storage1/StreamName"), case-insensitive.
    pub fn find_entry_by_path(&self, path: &str) -> Option<usize> {
        let parts: Vec<&str> = path.split('/').collect();
        if parts.is_empty() {
            return None;
        }

        // Start from root entry (index 0).
        if self.entries.is_empty() || self.entries[0].entry_type != EntryType::RootStorage {
            return None;
        }

        let mut current_child = self.entries[0].child;

        for (i, part) in parts.iter().enumerate() {
            let is_last = i == parts.len() - 1;
            let found = self.find_in_tree(current_child, part)?;

            if is_last {
                return Some(found);
            }

            // Must be a storage to traverse into.
            let entry = &self.entries[found];
            if entry.entry_type != EntryType::Storage && entry.entry_type != EntryType::RootStorage
            {
                return None;
            }
            current_child = entry.child;
        }

        None
    }

    /// Search the red-black tree rooted at `node_id` for an entry matching `name`.
    fn find_in_tree(&self, node_id: u32, name: &str) -> Option<usize> {
        if node_id == NO_ENTRY || node_id as usize >= self.entries.len() {
            return None;
        }

        let entry = &self.entries[node_id as usize];
        let lower = name.to_ascii_lowercase();

        if entry.name.to_ascii_lowercase() == lower {
            return Some(node_id as usize);
        }

        // Search both subtrees (the tree may not be well-ordered in malformed files).
        self.find_in_tree(entry.left_sibling, name)
            .or_else(|| self.find_in_tree(entry.right_sibling, name))
    }

    /// Read a stream by directory entry index.
    pub fn read_stream_by_index(&mut self, index: usize) -> Result<Vec<u8>> {
        let entry = self.entries.get(index).ok_or_else(|| {
            CfbError::StreamNotFound(format!("no entry at index {index}"))
        })?;

        let size = entry.stream_size as usize;
        let start = entry.start_sector;

        if size == 0 {
            return Ok(Vec::new());
        }

        // Decide: regular stream or mini-stream?
        // Use mini-stream only if: size < cutoff, not root, and mini-stream exists.
        if size < self.header.mini_stream_cutoff as usize
            && entry.entry_type != EntryType::RootStorage
            && !self.mini_stream.is_empty()
        {
            self.read_mini_stream(start, size)
        } else {
            let data = Self::read_chain(&mut self.reader, &self.header, &self.fat, start)?;
            Ok(data[..size.min(data.len())].to_vec())
        }
    }

    /// Open a stream by name (case-insensitive).
    pub fn open_stream(&mut self, name: &str) -> Result<Vec<u8>> {
        let idx = self
            .find_entry(name)
            .ok_or_else(|| CfbError::StreamNotFound(name.to_string()))?;
        self.read_stream_by_index(idx)
    }

    /// Open a stream by path (e.g. "ObjectPool/MyObj/\x01CompObj").
    pub fn open_stream_by_path(&mut self, path: &str) -> Result<Vec<u8>> {
        let idx = self
            .find_entry_by_path(path)
            .ok_or_else(|| CfbError::StreamNotFound(path.to_string()))?;
        self.read_stream_by_index(idx)
    }

    /// Check if a stream with the given name exists.
    pub fn has_stream(&self, name: &str) -> bool {
        self.find_entry(name).is_some()
    }

    // ── Internal helpers ──

    /// Build the complete FAT from DIFAT entries (header + DIFAT chain).
    fn read_fat(reader: &mut R, header: &CfbHeader) -> Result<Vec<u32>> {
        // Collect all FAT sector locations from DIFAT.
        let mut fat_sectors: Vec<u32> = header
            .header_difat
            .iter()
            .copied()
            .filter(|&s| s <= MAX_REG_SECT)
            .collect();

        // Follow the DIFAT chain for large files.
        let mut difat_sector = header.first_difat_sector;
        let entries_per_difat = header.sector_size / 4 - 1; // last u32 is next DIFAT sector
        while difat_sector <= MAX_REG_SECT {
            let mut sector_buf = vec![0u8; header.sector_size];
            reader.seek(SeekFrom::Start(header.sector_offset(difat_sector)))?;
            let n = read_fully(reader, &mut sector_buf)?;
            if n < header.sector_size {
                sector_buf[n..].fill(0xFF);
            }

            for i in 0..entries_per_difat {
                let off = i * 4;
                let val = u32::from_le_bytes([
                    sector_buf[off],
                    sector_buf[off + 1],
                    sector_buf[off + 2],
                    sector_buf[off + 3],
                ]);
                if val <= MAX_REG_SECT {
                    fat_sectors.push(val);
                }
            }

            // Next DIFAT sector.
            let next_off = entries_per_difat * 4;
            difat_sector = u32::from_le_bytes([
                sector_buf[next_off],
                sector_buf[next_off + 1],
                sector_buf[next_off + 2],
                sector_buf[next_off + 3],
            ]);
        }

        // Read each FAT sector and concatenate entries.
        let entries_per_fat_sector = header.sector_size / 4;
        let mut fat =
            Vec::with_capacity(fat_sectors.len() * entries_per_fat_sector);
        let mut sector_buf = vec![0u8; header.sector_size];

        for &fat_sec in &fat_sectors {
            reader.seek(SeekFrom::Start(header.sector_offset(fat_sec)))?;
            let n = read_fully(reader, &mut sector_buf)?;
            if n < header.sector_size {
                // Zero-fill remainder for truncated sectors.
                sector_buf[n..].fill(0xFF); // FREE_SECT
            }
            for i in 0..entries_per_fat_sector {
                let off = i * 4;
                fat.push(u32::from_le_bytes([
                    sector_buf[off],
                    sector_buf[off + 1],
                    sector_buf[off + 2],
                    sector_buf[off + 3],
                ]));
            }
        }

        Ok(fat)
    }

    /// Read a chain of sectors starting at `start` and return the concatenated data.
    fn read_chain(reader: &mut R, header: &CfbHeader, fat: &[u32], start: u32) -> Result<Vec<u8>> {
        let mut data = Vec::new();
        let mut sector = start;
        let mut visited = 0u32;
        let max_sectors = fat.len() as u32 + 1; // safety limit

        while sector <= MAX_REG_SECT {
            if visited > max_sectors {
                return Err(CfbError::CorruptedStream("FAT chain cycle detected".into()));
            }

            let offset = header.sector_offset(sector);
            let mut buf = vec![0u8; header.sector_size];
            reader.seek(SeekFrom::Start(offset))?;
            // Tolerate truncated files: read as much as available.
            let n = read_fully(reader, &mut buf)?;
            if n == 0 {
                break;
            }
            data.extend_from_slice(&buf[..n]);

            // Follow chain.
            if (sector as usize) < fat.len() {
                sector = fat[sector as usize];
            } else {
                break;
            }
            visited += 1;
        }

        Ok(data)
    }

    /// Read from the mini-stream using mini-FAT chain.
    fn read_mini_stream(&self, start: u32, size: usize) -> Result<Vec<u8>> {
        let mut data = Vec::with_capacity(size);
        let mut sector = start;
        let mut remaining = size;
        let mini_sector_size = self.header.mini_sector_size;
        let max_sectors = self.mini_fat.len() as u32 + 1;
        let mut visited = 0u32;

        while sector <= MAX_REG_SECT && remaining > 0 {
            if visited > max_sectors {
                return Err(CfbError::CorruptedStream(
                    "mini-FAT chain cycle detected".into(),
                ));
            }

            let offset = sector as usize * mini_sector_size;
            let to_read = remaining.min(mini_sector_size);

            if offset + to_read <= self.mini_stream.len() {
                data.extend_from_slice(&self.mini_stream[offset..offset + to_read]);
            } else {
                // Tolerate truncated mini-stream.
                let available = self.mini_stream.len().saturating_sub(offset);
                if available > 0 {
                    data.extend_from_slice(&self.mini_stream[offset..offset + available]);
                }
                break;
            }

            remaining -= to_read;

            if (sector as usize) < self.mini_fat.len() {
                sector = self.mini_fat[sector as usize];
            } else {
                break;
            }
            visited += 1;
        }

        Ok(data)
    }
}

/// Read as much as possible into `buf`, returning the number of bytes read.
/// Unlike `read_exact`, this does not error on truncated input.
fn read_fully<R: Read>(reader: &mut R, buf: &mut [u8]) -> super::error::Result<usize> {
    let mut total = 0;
    while total < buf.len() {
        match reader.read(&mut buf[total..]) {
            Ok(0) => break,
            Ok(n) => total += n,
            Err(ref e) if e.kind() == std::io::ErrorKind::Interrupted => continue,
            Err(e) => return Err(e.into()),
        }
    }
    Ok(total)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cfb::header::{END_OF_CHAIN, FAT_SECT, FREE_SECT};
    use std::io::Cursor;

    /// Build a complete minimal CFB v3 file in memory with one stream.
    ///
    /// Layout (512-byte sectors):
    /// - Header (512 bytes)
    /// - Sector 0: Directory (4 entries × 128 bytes = 512 bytes)
    /// - Sector 1: FAT (128 entries × 4 bytes = 512 bytes)
    /// - Sector 2: Stream data ("Hello, CFB!")
    fn build_minimal_cfb() -> Vec<u8> {
        let sector_size = 512usize;

        // We'll have 3 sectors.
        let mut file = vec![0u8; 512 + 3 * sector_size]; // header + 3 sectors

        // ── Header ──
        // Signature
        file[0..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
        // Minor version
        file[0x18..0x1A].copy_from_slice(&0x003Eu16.to_le_bytes());
        // Major version = 3
        file[0x1A..0x1C].copy_from_slice(&3u16.to_le_bytes());
        // Byte order
        file[0x1C..0x1E].copy_from_slice(&0xFFFEu16.to_le_bytes());
        // Sector size power = 9 (512)
        file[0x1E..0x20].copy_from_slice(&9u16.to_le_bytes());
        // Mini sector size power = 6 (64)
        file[0x20..0x22].copy_from_slice(&6u16.to_le_bytes());
        // FAT sector count = 1
        file[0x2C..0x30].copy_from_slice(&1u32.to_le_bytes());
        // First directory sector = 0
        file[0x30..0x34].copy_from_slice(&0u32.to_le_bytes());
        // Mini-stream cutoff = 4096
        file[0x38..0x3C].copy_from_slice(&4096u32.to_le_bytes());
        // First mini-FAT sector = END_OF_CHAIN (no mini-FAT)
        file[0x3C..0x40].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
        // Mini-FAT sector count = 0
        file[0x40..0x44].copy_from_slice(&0u32.to_le_bytes());
        // First DIFAT sector = END_OF_CHAIN (no DIFAT chain)
        file[0x44..0x48].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
        // DIFAT sector count = 0
        file[0x48..0x4C].copy_from_slice(&0u32.to_le_bytes());
        // DIFAT[0] = sector 1 (FAT)
        file[0x4C..0x50].copy_from_slice(&1u32.to_le_bytes());
        // DIFAT[1..109] = FREE_SECT
        for i in 1..109 {
            let off = 0x4C + i * 4;
            file[off..off + 4].copy_from_slice(&FREE_SECT.to_le_bytes());
        }

        // ── Sector 0: Directory ──
        let dir_offset = 512;

        // Entry 0: Root Entry
        write_dir_entry(
            &mut file[dir_offset..dir_offset + 128],
            "Root Entry",
            5,  // root storage
            1,  // child = entry 1
            END_OF_CHAIN,
            0,
        );

        // Entry 1: "TestStream" (stream)
        write_dir_entry(
            &mut file[dir_offset + 128..dir_offset + 256],
            "TestStream",
            2, // stream
            NO_ENTRY,
            2,  // start sector = 2
            11, // size = 11 ("Hello, CFB!")
        );

        // Entry 2-3: Empty
        file[dir_offset + 256 + 0x42] = 0; // empty
        file[dir_offset + 384 + 0x42] = 0; // empty

        // ── Sector 1: FAT ──
        let fat_offset = 512 + sector_size;
        // Sector 0: END_OF_CHAIN (directory, single sector)
        write_fat_entry(&mut file, fat_offset, 0, END_OF_CHAIN);
        // Sector 1: FAT_SECT (this sector is a FAT sector)
        write_fat_entry(&mut file, fat_offset, 1, FAT_SECT);
        // Sector 2: END_OF_CHAIN (stream data)
        write_fat_entry(&mut file, fat_offset, 2, END_OF_CHAIN);
        // Rest: FREE_SECT
        for i in 3..128 {
            write_fat_entry(&mut file, fat_offset, i, FREE_SECT);
        }

        // ── Sector 2: Stream data ──
        let data_offset = 512 + 2 * sector_size;
        let stream_data = b"Hello, CFB!";
        file[data_offset..data_offset + stream_data.len()].copy_from_slice(stream_data);

        file
    }

    fn write_dir_entry(
        buf: &mut [u8],
        name: &str,
        entry_type: u8,
        child: u32,
        start_sector: u32,
        stream_size: u32,
    ) {
        let utf16: Vec<u16> = name.encode_utf16().collect();
        for (i, &ch) in utf16.iter().enumerate() {
            let bytes = ch.to_le_bytes();
            buf[i * 2] = bytes[0];
            buf[i * 2 + 1] = bytes[1];
        }
        let name_size = ((utf16.len() + 1) * 2) as u16;
        buf[0x40..0x42].copy_from_slice(&name_size.to_le_bytes());
        buf[0x42] = entry_type;
        buf[0x43] = 1; // black
        buf[0x44..0x48].copy_from_slice(&NO_ENTRY.to_le_bytes()); // left
        buf[0x48..0x4C].copy_from_slice(&NO_ENTRY.to_le_bytes()); // right
        buf[0x4C..0x50].copy_from_slice(&child.to_le_bytes());
        buf[0x74..0x78].copy_from_slice(&start_sector.to_le_bytes());
        buf[0x78..0x7C].copy_from_slice(&stream_size.to_le_bytes());
    }

    fn write_fat_entry(file: &mut [u8], fat_offset: usize, index: usize, value: u32) {
        let off = fat_offset + index * 4;
        file[off..off + 4].copy_from_slice(&value.to_le_bytes());
    }

    #[test]
    fn open_minimal_cfb() {
        let data = build_minimal_cfb();
        let cursor = Cursor::new(data);
        let reader = CfbReader::new(cursor).unwrap();
        assert_eq!(reader.header().major_version, 3);
        assert_eq!(reader.entries().len(), 4);
        assert_eq!(reader.entries()[0].name, "Root Entry");
        assert_eq!(reader.entries()[1].name, "TestStream");
    }

    #[test]
    fn read_stream_by_name() {
        let data = build_minimal_cfb();
        let cursor = Cursor::new(data);
        let mut reader = CfbReader::new(cursor).unwrap();
        let stream = reader.open_stream("TestStream").unwrap();
        assert_eq!(&stream, b"Hello, CFB!");
    }

    #[test]
    fn read_stream_case_insensitive() {
        let data = build_minimal_cfb();
        let cursor = Cursor::new(data);
        let mut reader = CfbReader::new(cursor).unwrap();
        let stream = reader.open_stream("teststream").unwrap();
        assert_eq!(&stream, b"Hello, CFB!");
    }

    #[test]
    fn stream_not_found() {
        let data = build_minimal_cfb();
        let cursor = Cursor::new(data);
        let mut reader = CfbReader::new(cursor).unwrap();
        assert!(reader.open_stream("NonExistent").is_err());
    }

    #[test]
    fn has_stream() {
        let data = build_minimal_cfb();
        let cursor = Cursor::new(data);
        let reader = CfbReader::new(cursor).unwrap();
        assert!(reader.has_stream("TestStream"));
        assert!(reader.has_stream("teststream"));
        assert!(!reader.has_stream("Missing"));
    }

    /// Build a CFB with a small stream that goes into the mini-stream.
    fn build_cfb_with_mini_stream() -> Vec<u8> {
        let sector_size = 512usize;
        // Layout:
        // Header (512)
        // Sector 0: Directory
        // Sector 1: FAT
        // Sector 2: Mini-stream container (Root Entry data, holds mini-stream data)
        // Sector 3: Mini-FAT
        let mut file = vec![0u8; 512 + 4 * sector_size];

        // ── Header ──
        file[0..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
        file[0x18..0x1A].copy_from_slice(&0x003Eu16.to_le_bytes());
        file[0x1A..0x1C].copy_from_slice(&3u16.to_le_bytes());
        file[0x1C..0x1E].copy_from_slice(&0xFFFEu16.to_le_bytes());
        file[0x1E..0x20].copy_from_slice(&9u16.to_le_bytes());
        file[0x20..0x22].copy_from_slice(&6u16.to_le_bytes());
        file[0x2C..0x30].copy_from_slice(&1u32.to_le_bytes());
        file[0x30..0x34].copy_from_slice(&0u32.to_le_bytes());
        file[0x38..0x3C].copy_from_slice(&4096u32.to_le_bytes());
        // First mini-FAT sector = 3
        file[0x3C..0x40].copy_from_slice(&3u32.to_le_bytes());
        file[0x40..0x44].copy_from_slice(&1u32.to_le_bytes()); // mini-FAT count = 1
        file[0x44..0x48].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
        file[0x48..0x4C].copy_from_slice(&0u32.to_le_bytes());
        // DIFAT[0] = sector 1
        file[0x4C..0x50].copy_from_slice(&1u32.to_le_bytes());
        for i in 1..109 {
            let off = 0x4C + i * 4;
            file[off..off + 4].copy_from_slice(&FREE_SECT.to_le_bytes());
        }

        let dir_offset = 512;
        // Root Entry: start_sector=2 (mini-stream container), stream_size=512 (container size)
        write_dir_entry(
            &mut file[dir_offset..dir_offset + 128],
            "Root Entry",
            5,
            1, // child = entry 1
            2, // start sector (mini-stream container)
            512, // mini-stream container size
        );
        // Entry 1: "SmallStream" — small stream, goes to mini-stream
        // start_sector = 0 (mini-sector 0), size = 5
        write_dir_entry(
            &mut file[dir_offset + 128..dir_offset + 256],
            "SmallStream",
            2,
            NO_ENTRY,
            0, // start mini-sector
            5, // 5 bytes
        );
        // Empty entries
        file[dir_offset + 256 + 0x42] = 0;
        file[dir_offset + 384 + 0x42] = 0;

        // ── Sector 1: FAT ──
        let fat_offset = 512 + sector_size;
        write_fat_entry(&mut file, fat_offset, 0, END_OF_CHAIN); // dir
        write_fat_entry(&mut file, fat_offset, 1, FAT_SECT);     // FAT
        write_fat_entry(&mut file, fat_offset, 2, END_OF_CHAIN); // mini-stream container
        write_fat_entry(&mut file, fat_offset, 3, END_OF_CHAIN); // mini-FAT
        for i in 4..128 {
            write_fat_entry(&mut file, fat_offset, i, FREE_SECT);
        }

        // ── Sector 2: Mini-stream container ──
        let ms_offset = 512 + 2 * sector_size;
        file[ms_offset..ms_offset + 5].copy_from_slice(b"Small");

        // ── Sector 3: Mini-FAT ──
        let mf_offset = 512 + 3 * sector_size;
        // Mini-sector 0: END_OF_CHAIN
        mf_offset_write(&mut file, mf_offset, 0, END_OF_CHAIN);
        for i in 1..128 {
            mf_offset_write(&mut file, mf_offset, i, FREE_SECT);
        }

        file
    }

    fn mf_offset_write(file: &mut [u8], base: usize, index: usize, value: u32) {
        let off = base + index * 4;
        file[off..off + 4].copy_from_slice(&value.to_le_bytes());
    }

    #[test]
    fn read_mini_stream() {
        let data = build_cfb_with_mini_stream();
        let cursor = Cursor::new(data);
        let mut reader = CfbReader::new(cursor).unwrap();
        let stream = reader.open_stream("SmallStream").unwrap();
        assert_eq!(&stream, b"Small");
    }

    #[test]
    fn find_entry_by_path_simple() {
        let data = build_minimal_cfb();
        let cursor = Cursor::new(data);
        let reader = CfbReader::new(cursor).unwrap();
        // "TestStream" is a child of root.
        let idx = reader.find_entry_by_path("TestStream");
        assert_eq!(idx, Some(1));
    }
}
