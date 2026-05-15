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
            let rec_len =
                u32::from_le_bytes([data[pos + 4], data[pos + 5], data[pos + 6], data[pos + 7]])
                    as usize;
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
    if data.is_empty() {
        return false;
    }
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn blip_type_recognition() {
        assert!(is_blip_type(0xF01D));
        assert!(is_blip_type(0xF01E));
        assert!(is_blip_type(0xF02A));
        assert!(!is_blip_type(0xF000));
        assert!(!is_blip_type(0xF020));
    }

    #[test]
    fn uid_size_secondary_uid() {
        // Bit 0 of inst signals a secondary UID — adds 16 bytes.
        assert_eq!(uid_size(0xF01D, 0b00), 17);
        assert_eq!(uid_size(0xF01D, 0b01), 33);
        // 0xF01A..=0xF01C use the metafile-style header layout (base 16).
        assert_eq!(uid_size(0xF01A, 0b00), 16);
        assert_eq!(uid_size(0xF01A, 0b01), 32);
    }

    #[test]
    fn metafile_header_only_for_metafile_types() {
        assert_eq!(metafile_header_size(0xF01A), 34);
        assert_eq!(metafile_header_size(0xF01B), 34);
        assert_eq!(metafile_header_size(0xF01C), 34);
        assert_eq!(metafile_header_size(0xF01D), 0);
        assert_eq!(metafile_header_size(0xF01E), 0);
    }

    #[test]
    fn signature_validation() {
        // JPEG starts with FFD8.
        assert!(has_valid_signature(0xF01D, &[0xFF, 0xD8, 0x00]));
        assert!(!has_valid_signature(0xF01D, &[0x00, 0x00]));
        // PNG starts with 89 50 4E 47.
        assert!(has_valid_signature(0xF01E, b"\x89PNG\r\n"));
        assert!(!has_valid_signature(0xF01E, b"WRONG"));
        // EMF: 01 00 00 00 prefix.
        assert!(has_valid_signature(0xF01A, &[0x01, 0x00, 0x00, 0x00, 0xAA]));
        assert!(!has_valid_signature(0xF01A, &[0x00, 0x00, 0x00, 0x00]));
        // Empty payload always invalid.
        assert!(!has_valid_signature(0xF01D, &[]));
    }

    #[test]
    fn to_format_mapping() {
        assert!(matches!(to_format(0xF01A), ImageFormat::Emf));
        assert!(matches!(to_format(0xF01B), ImageFormat::Wmf));
        assert!(matches!(to_format(0xF01C), ImageFormat::Pict));
        assert!(matches!(to_format(0xF01D), ImageFormat::Jpeg));
        assert!(matches!(to_format(0xF02A), ImageFormat::Jpeg));
        assert!(matches!(to_format(0xF01E), ImageFormat::Png));
        assert!(matches!(to_format(0xF01F), ImageFormat::Dib));
        assert!(matches!(to_format(0xF029), ImageFormat::Tiff));
        assert!(matches!(to_format(0xABCD), ImageFormat::Unknown(0xABCD)));
    }

    #[test]
    fn extract_images_skips_non_blip_bytes() {
        // Random non-BLIP bytes produce no images and never crash.
        let data = vec![0u8; 64];
        assert!(extract_images(&data).is_empty());
    }

    #[test]
    fn extract_images_finds_embedded_png() {
        // Synthesize a record header followed by a PNG signature so the
        // scanner descends into a valid BLIP payload.
        let rec_type: u16 = 0xF01E; // PNG
        let inst: u16 = 0; // no secondary UID
        let ver_inst: u16 = inst << 4;
        let uid = 17usize; // base for non-metafile
        let png_body = b"\x89PNG\r\n\x1a\nIHDRfakebody";
        let payload_len = uid + png_body.len();

        let mut data = Vec::new();
        data.extend_from_slice(&ver_inst.to_le_bytes());
        data.extend_from_slice(&rec_type.to_le_bytes());
        data.extend_from_slice(&(payload_len as u32).to_le_bytes());
        data.extend_from_slice(&[0u8; 17]); // skipped UID bytes
        data.extend_from_slice(png_body);

        let images = extract_images(&data);
        assert_eq!(images.len(), 1);
        assert!(matches!(images[0].format, ImageFormat::Png));
        assert_eq!(images[0].data, png_body);
        assert_eq!(images[0].index, 0);
    }
}
