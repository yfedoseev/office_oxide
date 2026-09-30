//! Image extraction from DOC files.
//!
//! Images in Word binary files are stored as OfficeArt BLIP records
//! embedded at various offsets within the `Data` stream. Unlike PPT
//! where the Pictures stream is a flat sequence of BLIPs, DOC's Data
//! stream contains mixed data. We scan for BLIP record signatures.

pub use crate::cfb::blip::BlipFormat as ImageFormat;
pub use crate::cfb::blip::BlipImage as DocImage;
use crate::cfb::blip::extract_one_blip;

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
            let data_end = data_start.saturating_add(rec_len).min(data.len());

            // The same record decoder the PPT `Pictures` walk uses —
            // UIDs skipped, compressed metafiles inflated — then the
            // signature check this heuristic scan needs to reject byte
            // runs that only look like a record header, applied to the
            // decoded image rather than to still-compressed bytes.
            if let Some(mut img) = extract_one_blip(data, rec_type, inst, data_start, data_end) {
                if has_valid_signature(rec_type, &img.data) {
                    img.index = images.len();
                    images.push(img);
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cfb::blip::{metafile_header_size, uid_size};

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
    fn test_scan_finds_jpeg_in_data_stream() {
        let data = make_blip_in_data(0xF01D, 0x46A, b"\xff\xd8\xff\xe0JFIF");
        let images = extract_images(&data);
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].format, ImageFormat::Jpeg);
        assert!(images[0].data.starts_with(b"\xff\xd8"));
    }

    #[test]
    fn test_scan_finds_png_in_data_stream() {
        let data = make_blip_in_data(0xF01E, 0x6E0, b"\x89PNG\r\n\x1a\nIHDR");
        let images = extract_images(&data);
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].format, ImageFormat::Png);
        assert!(images[0].data.starts_with(b"\x89PNG"));
    }

    #[test]
    fn test_scan_finds_multiple_images() {
        let mut data = make_blip_in_data(0xF01D, 0x46A, b"\xff\xd8\xff\xe0JPEG1");
        data.extend(make_blip_in_data(0xF01E, 0x6E0, b"\x89PNG\r\n\x1a\nPNG2"));
        let images = extract_images(&data);
        assert_eq!(images.len(), 2);
        assert_eq!(images[0].format, ImageFormat::Jpeg);
        assert_eq!(images[1].format, ImageFormat::Png);
    }

    #[test]
    fn test_rejects_false_positive() {
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

    /// A compressed EMF ([MS-ODRAW] §2.2.31, `compression` = 0x00) failed
    /// the EMF signature check on its still-compressed bytes and was
    /// silently dropped. It is inflated first, then validated.
    #[test]
    fn test_scan_finds_compressed_emf_in_data_stream() {
        use std::io::Write;
        let mut emf = vec![0x01, 0x00, 0x00, 0x00, 0x6C, 0x00, 0x00, 0x00];
        emf.extend(std::iter::repeat_n(0x41u8, 120));
        let mut enc = flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::default());
        enc.write_all(&emf).unwrap();
        let z = enc.finish().unwrap();
        let mut body = vec![0u8; 16]; // rgbUid1
        body.extend_from_slice(&(emf.len() as u32).to_le_bytes()); // cbSize
        body.extend_from_slice(&[0u8; 24]); // rcBounds + ptSize
        body.extend_from_slice(&(z.len() as u32).to_le_bytes()); // cbSave
        body.extend_from_slice(&[0x00, 0xFE]); // compression = deflate, filter
        body.extend_from_slice(&z);
        let mut data = vec![0u8; 40];
        data.extend_from_slice(&(0x3D4u16 << 4).to_le_bytes());
        data.extend_from_slice(&0xF01Au16.to_le_bytes());
        data.extend_from_slice(&(body.len() as u32).to_le_bytes());
        data.extend(body);
        let images = extract_images(&data);
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].format, ImageFormat::Emf);
        assert_eq!(images[0].data, emf);
    }

    #[test]
    fn test_empty_data_stream() {
        assert!(extract_images(&[]).is_empty());
    }
}
