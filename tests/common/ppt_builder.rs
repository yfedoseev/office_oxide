//! Synthetic legacy PowerPoint (`.ppt`) files and multi-stream compound
//! files, built in code for integration tests.
//!
//! Included with `#[path = "common/ppt_builder.rs"] mod ppt_builder;` by
//! the test files that need it.

/// A multi-stream CFB v3 container: every stream sits at the root, in
/// consecutive 512-byte sectors (no mini stream), chained as right
/// siblings of the first.
pub fn cfb_with_streams(streams: &[(&str, &[u8])]) -> Vec<u8> {
    const END_OF_CHAIN: u32 = 0xFFFF_FFFE;
    const FAT_SECT: u32 = 0xFFFF_FFFD;
    const FREE_SECT: u32 = 0xFFFF_FFFF;
    const NO_ENTRY: u32 = 0xFFFF_FFFF;

    let dir_sectors = (streams.len() + 1).div_ceil(4);
    let data_sectors: Vec<usize> = streams
        .iter()
        .map(|(_, d)| d.len().div_ceil(512).max(1))
        .collect();
    let total_data: usize = data_sectors.iter().sum();
    // FAT sectors must cover themselves too.
    let mut fat_sectors = 1;
    while fat_sectors * 128 < dir_sectors + fat_sectors + total_data {
        fat_sectors += 1;
    }
    assert!(fat_sectors <= 109, "test builder has no DIFAT support");
    let total = dir_sectors + fat_sectors + total_data;
    let mut file = vec![0u8; 512 * (1 + total)];
    file[0..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
    file[0x18..0x1A].copy_from_slice(&0x003Eu16.to_le_bytes());
    file[0x1A..0x1C].copy_from_slice(&3u16.to_le_bytes());
    file[0x1C..0x1E].copy_from_slice(&0xFFFEu16.to_le_bytes());
    file[0x1E..0x20].copy_from_slice(&9u16.to_le_bytes());
    file[0x20..0x22].copy_from_slice(&6u16.to_le_bytes());
    file[0x2C..0x30].copy_from_slice(&(fat_sectors as u32).to_le_bytes());
    file[0x30..0x34].copy_from_slice(&0u32.to_le_bytes()); // first dir sector
    file[0x38..0x3C].copy_from_slice(&4096u32.to_le_bytes());
    file[0x3C..0x40].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
    file[0x44..0x48].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
    for i in 0..109 {
        let v = if i < fat_sectors {
            (dir_sectors + i) as u32
        } else {
            FREE_SECT
        };
        file[0x4C + i * 4..0x50 + i * 4].copy_from_slice(&v.to_le_bytes());
    }

    let mut fat = vec![FREE_SECT; fat_sectors * 128];
    for s in 0..dir_sectors {
        fat[s] = if s + 1 == dir_sectors {
            END_OF_CHAIN
        } else {
            (s + 1) as u32
        };
    }
    for s in 0..fat_sectors {
        fat[dir_sectors + s] = FAT_SECT;
    }

    let write_entry = |file: &mut [u8],
                       n: usize,
                       name: &str,
                       kind: u8,
                       right: u32,
                       child: u32,
                       start: u32,
                       size: u32| {
        let off = 512 + n * 128;
        let utf16: Vec<u16> = name.encode_utf16().collect();
        for (i, ch) in utf16.iter().enumerate() {
            file[off + i * 2..off + i * 2 + 2].copy_from_slice(&ch.to_le_bytes());
        }
        file[off + 0x40..off + 0x42]
            .copy_from_slice(&(((utf16.len() + 1) * 2) as u16).to_le_bytes());
        file[off + 0x42] = kind;
        file[off + 0x43] = 1;
        file[off + 0x44..off + 0x48].copy_from_slice(&NO_ENTRY.to_le_bytes());
        file[off + 0x48..off + 0x4C].copy_from_slice(&right.to_le_bytes());
        file[off + 0x4C..off + 0x50].copy_from_slice(&child.to_le_bytes());
        file[off + 0x74..off + 0x78].copy_from_slice(&start.to_le_bytes());
        file[off + 0x78..off + 0x7C].copy_from_slice(&size.to_le_bytes());
    };
    // Unused directory slots must be free entries (all zero is not: the
    // sibling fields must be NOSTREAM).
    for n in 0..dir_sectors * 4 {
        let off = 512 + n * 128;
        file[off + 0x44..off + 0x50].copy_from_slice(&[0xFF; 12]);
    }
    let root_child = if streams.is_empty() { NO_ENTRY } else { 1 };
    write_entry(&mut file, 0, "Root Entry", 5, NO_ENTRY, root_child, END_OF_CHAIN, 0);

    let mut next = dir_sectors + fat_sectors;
    for (i, ((name, data), &count)) in streams.iter().zip(&data_sectors).enumerate() {
        let start = next;
        for k in 0..count {
            let s = start + k;
            fat[s] = if k + 1 == count {
                END_OF_CHAIN
            } else {
                (s + 1) as u32
            };
        }
        let off = 512 + start * 512;
        file[off..off + data.len()].copy_from_slice(data);
        let right = if i + 1 < streams.len() {
            (i + 2) as u32
        } else {
            NO_ENTRY
        };
        write_entry(&mut file, i + 1, name, 2, right, NO_ENTRY, start as u32, data.len() as u32);
        next += count;
    }
    for (s, v) in fat.iter().enumerate() {
        let sector = dir_sectors + s / 128;
        let off = 512 + sector * 512 + (s % 128) * 4;
        file[off..off + 4].copy_from_slice(&v.to_le_bytes());
    }
    file
}

/// One MS-PPT record: `recVer`/`recInstance` packed in the first u16
/// (MS-PPT §2.3.1 RecordHeader), then the type and the body length.
pub fn ppt_rec(rec_type: u16, ver_instance: u16, data: &[u8]) -> Vec<u8> {
    let mut b = ver_instance.to_le_bytes().to_vec();
    b.extend_from_slice(&rec_type.to_le_bytes());
    b.extend_from_slice(&(data.len() as u32).to_le_bytes());
    b.extend_from_slice(data);
    b
}

/// `RT_Slide` container (MS-PPT §2.13.24 RecordType 0x03EE).
pub const RT_SLIDE: u16 = 0x03EE;
/// `RT_TextHeaderAtom` (MS-PPT §2.13.24 RecordType 0x0F9F).
pub const RT_TEXT_HEADER: u16 = 0x0F9F;
/// `RT_TextBytesAtom` (MS-PPT §2.13.24 RecordType 0x0FA8).
pub const RT_TEXT_BYTES: u16 = 0x0FA8;
/// `OfficeArtClientTextbox` (MS-ODRAW record type 0xF00D; MS-PPT
/// `OfficeArtClientTextbox` wraps a shape's text records in it).
pub const RT_CLIENT_TEXTBOX: u16 = 0xF00D;

/// A `Slide` container holding one client text box with `text` as a
/// `TextBytesAtom` — the smallest slide the reader extracts text from.
pub fn ppt_slide(text: &str) -> Vec<u8> {
    let mut textbox = ppt_rec(RT_TEXT_HEADER, 0, &0u32.to_le_bytes());
    textbox.extend(ppt_rec(RT_TEXT_BYTES, 0, text.as_bytes()));
    let textbox = ppt_rec(RT_CLIENT_TEXTBOX, 0x000F, &textbox);
    ppt_rec(RT_SLIDE, 0x000F, &textbox)
}

/// A `.ppt` whose `PowerPoint Document` stream holds one slide per entry
/// of `slides`, plus any `extra` streams (e.g. OLE property sets).
pub fn build_ppt(slides: &[&str], extra: &[(&str, &[u8])]) -> Vec<u8> {
    let stream: Vec<u8> = slides.iter().flat_map(|s| ppt_slide(s)).collect();
    let mut streams: Vec<(&str, &[u8])> = vec![("PowerPoint Document", &stream)];
    streams.extend_from_slice(extra);
    cfb_with_streams(&streams)
}
