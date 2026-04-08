//! Image extraction from XLS files.
//!
//! Images in Excel binary files are stored as OfficeArt BLIP records
//! embedded within MSODRAWINGGROUP records in the Workbook stream.
//! We use the shared BLIP parser from cfb.

pub use crate::cfb::blip::BlipFormat as ImageFormat;
pub use crate::cfb::blip::BlipImage as XlsImage;

/// Extract images from an XLS Workbook stream by scanning for BLIP signatures.
///
/// BLIPs in XLS are embedded inside MSODRAWINGGROUP BIFF records,
/// nested within OfficeArt containers. We scan byte-by-byte.
pub fn extract_images(data: &[u8]) -> Vec<XlsImage> {
    let mut images = Vec::new();
    let mut pos = 0;

    while pos + 8 <= data.len() {
        let rec_type = u16::from_le_bytes([data[pos + 2], data[pos + 3]]);

        if is_blip_type(rec_type) {
            let ver_inst = u16::from_le_bytes([data[pos], data[pos + 1]]);
            let rec_len = u32::from_le_bytes([
                data[pos + 4], data[pos + 5], data[pos + 6], data[pos + 7],
            ]) as usize;
            let inst = ver_inst >> 4;

            let data_start = pos + 8;
            let data_end = (data_start + rec_len).min(data.len());

            let skip = uid_size(rec_type, inst) + metafile_header_size(rec_type);
            let img_start = data_start + skip;

            if img_start < data_end {
                let img_data = &data[img_start..data_end];
                if has_valid_signature(rec_type, img_data) {
                    images.push(XlsImage {
                        format: to_format(rec_type),
                        data: img_data.to_vec(),
                        index: images.len(),
                    });
                }
            }
            pos = data_end;
        } else {
            pos += 1;
        }
    }

    images
}

fn is_blip_type(rt: u16) -> bool {
    matches!(rt, 0xF01A..=0xF01F | 0xF029 | 0xF02A)
}

fn uid_size(rec_type: u16, inst: u16) -> usize {
    let base = match rec_type {
        0xF01A..=0xF01C => 16,
        _ => 17,
    };
    if inst & 1 != 0 { base + 16 } else { base }
}

fn metafile_header_size(rec_type: u16) -> usize {
    match rec_type {
        0xF01A..=0xF01C => 34,
        _ => 0,
    }
}

fn has_valid_signature(rec_type: u16, data: &[u8]) -> bool {
    if data.is_empty() { return false; }
    match rec_type {
        0xF01D | 0xF02A => data.len() >= 2 && data[0] == 0xFF && data[1] == 0xD8,
        0xF01E => data.len() >= 4 && data.starts_with(b"\x89PNG"),
        0xF01A => data.len() >= 4 && data[..4] == [0x01, 0x00, 0x00, 0x00],
        _ => data.len() > 10,
    }
}

fn to_format(rt: u16) -> ImageFormat {
    match rt {
        0xF01A => ImageFormat::Emf,
        0xF01B => ImageFormat::Wmf,
        0xF01C => ImageFormat::Pict,
        0xF01D | 0xF02A => ImageFormat::Jpeg,
        0xF01E => ImageFormat::Png,
        0xF01F => ImageFormat::Dib,
        0xF029 => ImageFormat::Tiff,
        other => ImageFormat::Unknown(other),
    }
}
