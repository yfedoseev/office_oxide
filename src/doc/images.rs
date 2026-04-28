//! Image extraction from DOC files.
//!
//! Images in Word binary files are stored as OfficeArt BLIP records
//! embedded at various offsets within the `Data` stream. Unlike PPT
//! where the Pictures stream is a flat sequence of BLIPs, DOC's Data
//! stream contains mixed data. We scan for BLIP record signatures.

pub use crate::cfb::blip::BlipFormat as ImageFormat;
pub use crate::cfb::blip::BlipImage as DocImage;

/// Extract images from a DOC Data stream by scanning for BLIP record signatures.
pub fn extract_images(data: &[u8]) -> Vec<DocImage> {
    let mut images = Vec::new();
    let mut pos = 0;

    while pos + 8 <= data.len() {
        let rec_type = u16::from_le_bytes([data[pos + 2], data[pos + 3]]);

        // Check if this looks like a BLIP record.
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

                // Validate: check for known image signatures to avoid false positives.
                if has_valid_signature(rec_type, img_data) {
                    images.push(DocImage {
                        format: ImageFormat::from_record_type(rec_type),
                        data: img_data.to_vec(),
                        index: images.len(),
                    });
                }
            }

            // Skip past this BLIP.
            pos = data_end;
        } else {
            pos += 1; // Scan byte-by-byte for next BLIP.
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

/// Check if the image data starts with a recognizable signature.
fn has_valid_signature(rec_type: u16, data: &[u8]) -> bool {
    if data.is_empty() {
        return false;
    }
    match rec_type {
        0xF01D | 0xF02A => data.len() >= 2 && data[0] == 0xFF && data[1] == 0xD8, // JPEG
        0xF01E => data.len() >= 4 && data.starts_with(b"\x89PNG"),                // PNG
        0xF01A => data.len() >= 4 && data[..4] == [0x01, 0x00, 0x00, 0x00],       // EMF
        0xF01B => data.len() > 10, // WMF (varied headers)
        _ => data.len() > 10,      // Others: trust if non-trivial
    }
}

// Re-export for use in ImageFormat construction
trait BlipFormatExt {
    fn from_record_type(rt: u16) -> Self;
}

impl BlipFormatExt for ImageFormat {
    fn from_record_type(rt: u16) -> Self {
        match rt {
            0xF01A => Self::Emf,
            0xF01B => Self::Wmf,
            0xF01C => Self::Pict,
            0xF01D | 0xF02A => Self::Jpeg,
            0xF01E => Self::Png,
            0xF01F => Self::Dib,
            0xF029 => Self::Tiff,
            other => Self::Unknown(other),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_blip_in_data(rec_type: u16, inst: u16, img_data: &[u8]) -> Vec<u8> {
        let ver_inst: u16 = inst << 4;
        let uid_sz = uid_size(rec_type, inst);
        let mf_sz = metafile_header_size(rec_type);
        let rec_len = uid_sz + mf_sz + img_data.len();

        // Prefix with some random non-BLIP data (simulating DOC Data stream).
        let mut buf = vec![0u8; 100]; // 100 bytes of junk before the BLIP
        buf.extend_from_slice(&ver_inst.to_le_bytes());
        buf.extend_from_slice(&rec_type.to_le_bytes());
        buf.extend_from_slice(&(rec_len as u32).to_le_bytes());
        buf.extend(vec![0u8; uid_sz]);
        buf.extend(vec![0u8; mf_sz]);
        buf.extend_from_slice(img_data);
        buf.extend(vec![0u8; 50]); // trailing junk
        buf
    }

    #[test]
    fn scan_finds_jpeg_in_data_stream() {
        let data = make_blip_in_data(0xF01D, 0x46A, b"\xff\xd8\xff\xe0JFIF");
        let images = extract_images(&data);
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].format, ImageFormat::Jpeg);
        assert!(images[0].data.starts_with(b"\xff\xd8"));
    }

    #[test]
    fn scan_finds_png_in_data_stream() {
        let data = make_blip_in_data(0xF01E, 0x6E0, b"\x89PNG\r\n\x1a\nIHDR");
        let images = extract_images(&data);
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].format, ImageFormat::Png);
        assert!(images[0].data.starts_with(b"\x89PNG"));
    }

    #[test]
    fn scan_finds_multiple_images() {
        let mut data = make_blip_in_data(0xF01D, 0x46A, b"\xff\xd8\xff\xe0JPEG1");
        data.extend(make_blip_in_data(0xF01E, 0x6E0, b"\x89PNG\r\n\x1a\nPNG2"));
        let images = extract_images(&data);
        assert_eq!(images.len(), 2);
        assert_eq!(images[0].format, ImageFormat::Jpeg);
        assert_eq!(images[1].format, ImageFormat::Png);
    }

    #[test]
    fn rejects_false_positive() {
        // Data that happens to have a BLIP type at the right offset but no valid image sig.
        let mut data = vec![0u8; 100];
        data[2] = 0x1D;
        data[3] = 0xF0; // looks like JPEG BLIP type
        data[4] = 30;
        data[5] = 0;
        data[6] = 0;
        data[7] = 0; // rec_len=30
        // But UID + tag (17 bytes) then data won't start with 0xFF 0xD8
        let images = extract_images(&data);
        assert!(images.is_empty());
    }

    #[test]
    fn empty_data_stream() {
        assert!(extract_images(&[]).is_empty());
    }
}
