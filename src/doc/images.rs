//! Image extraction from DOC files.
//!
//! Images in Word 97+ files are stored as OfficeArt BLIP records embedded
//! at various offsets within the `Data` stream. Unlike PPT where the
//! Pictures stream is a flat sequence of BLIPs, DOC's Data stream
//! contains mixed data. We scan for BLIP record signatures.
//!
//! A picture that is not an OfficeArt shape — every picture in a Word
//! 6.0/95 file (stored in the `WordDocument` stream itself), and old-style
//! pictures carried forward into Word 97 files — is a `PICF` header
//! ([MS-DOC] §2.9.192) followed directly by the picture's metafile or
//! bitmap. Those are found by the same scan: a `PICF`-shaped header whose
//! payload is a valid WMF or DIB.

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
        } else if let Some((img, len)) = picf_picture_at(data, pos) {
            images.push(DocImage {
                index: images.len(),
                ..img
            });
            pos += len;
        } else {
            pos += 1; // Scan byte-by-byte for next BLIP.
        }
    }

    images
}

/// `PICF.cbHeader`: [MS-DOC] §2.9.192 "MUST be 0x44".
const PICF_HEADER_SIZE: usize = 0x44;
/// `PICF.mfpf.mm` values whose payload is an OfficeArt shape, not a
/// picture ([MS-DOC] §2.9.155 `MFPF`: `MM_SHAPE`, `MM_SHAPEFILE`). The BLIP
/// scan handles those.
const MM_SHAPE: u16 = 0x0064;
const MM_SHAPEFILE: u16 = 0x0066;

/// A `PICF` picture starting at `pos`: its decoded image and the bytes the
/// whole structure spans. Requires a `PICF`-shaped header (`cbHeader` =
/// 0x44, an `lcb` that covers it and fits the buffer, a non-shape `mm`)
/// *and* a payload that is itself a valid WMF or DIB — the header alone
/// has no magic number to tell it from other bytes.
fn picf_picture_at(data: &[u8], pos: usize) -> Option<(DocImage, usize)> {
    let hdr = data.get(pos..pos + PICF_HEADER_SIZE)?;
    let lcb = i32::from_le_bytes([hdr[0], hdr[1], hdr[2], hdr[3]]);
    let cb_header = u16::from_le_bytes([hdr[4], hdr[5]]) as usize;
    let mm = u16::from_le_bytes([hdr[6], hdr[7]]);
    if cb_header != PICF_HEADER_SIZE || matches!(mm, MM_SHAPE | MM_SHAPEFILE) {
        return None;
    }
    let lcb = usize::try_from(lcb).ok()?;
    if lcb <= PICF_HEADER_SIZE || lcb > data.len() - pos {
        return None;
    }
    let payload = &data[pos + PICF_HEADER_SIZE..pos + lcb];
    let (format, len) = match wmf_len(payload) {
        Some(n) => (ImageFormat::Wmf, n),
        None => (ImageFormat::Dib, dib_len(payload)?),
    };
    Some((
        DocImage {
            format,
            data: payload[..len].to_vec(),
            index: 0,
        },
        lcb,
    ))
}

/// Length of a WMF at the start of `b`, from its `META_HEADER` ([MS-WMF]
/// §2.3.2.2: `Type` 1 or 2, `HeaderSize` 9 words, `Version` 0x0100 or
/// 0x0300, `Size` in 16-bit words), when the header is valid and the
/// metafile fits.
fn wmf_len(b: &[u8]) -> Option<usize> {
    if b.len() < 18 {
        return None;
    }
    let ty = u16::from_le_bytes([b[0], b[1]]);
    let header_size = u16::from_le_bytes([b[2], b[3]]);
    let version = u16::from_le_bytes([b[4], b[5]]);
    let words = u32::from_le_bytes([b[6], b[7], b[8], b[9]]) as usize;
    if !matches!(ty, 1 | 2) || header_size != 9 || !matches!(version, 0x0100 | 0x0300) {
        return None;
    }
    let len = words.checked_mul(2)?;
    (len >= 18 && len <= b.len()).then_some(len)
}

/// Length of a device-independent bitmap at the start of `b`: a
/// `BITMAPINFOHEADER` (40 bytes) with a plausible plane count and bit
/// depth. The DIB runs to the end of the `PICF` payload.
fn dib_len(b: &[u8]) -> Option<usize> {
    if b.len() < 40 {
        return None;
    }
    let size = u32::from_le_bytes([b[0], b[1], b[2], b[3]]);
    let planes = u16::from_le_bytes([b[12], b[13]]);
    let bits = u16::from_le_bytes([b[14], b[15]]);
    (size == 40 && planes == 1 && matches!(bits, 1 | 4 | 8 | 16 | 24 | 32)).then_some(b.len())
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

    /// A `PICF` + WMF picture (Word 6.0/95, or a non-shape picture in a
    /// Word 97 `Data` stream).
    fn picf_wmf(extra_after: usize) -> (Vec<u8>, Vec<u8>) {
        let mut wmf = Vec::new();
        wmf.extend_from_slice(&1u16.to_le_bytes()); // Type = memory
        wmf.extend_from_slice(&9u16.to_le_bytes()); // HeaderSize
        wmf.extend_from_slice(&0x0300u16.to_le_bytes()); // Version
        let words = (18 + 6) / 2;
        wmf.extend_from_slice(&(words as u32).to_le_bytes()); // Size
        wmf.extend_from_slice(&0u16.to_le_bytes()); // NumberOfObjects
        wmf.extend_from_slice(&3u32.to_le_bytes()); // MaxRecord
        wmf.extend_from_slice(&0u16.to_le_bytes()); // NumberOfMembers
        wmf.extend_from_slice(&[3, 0, 0, 0, 0, 0]); // META_EOF
        let mut picf = vec![0u8; 0x44];
        let lcb = (0x44 + wmf.len() + extra_after) as i32;
        picf[0..4].copy_from_slice(&lcb.to_le_bytes());
        picf[4..6].copy_from_slice(&0x44u16.to_le_bytes());
        picf[6..8].copy_from_slice(&8u16.to_le_bytes()); // MM_ANISOTROPIC
        picf.extend_from_slice(&wmf);
        picf.extend(std::iter::repeat_n(0u8, extra_after));
        (picf, wmf)
    }

    /// Non-shape pictures (`PICF` followed by the metafile itself) were
    /// never extracted: every Word 6.0/95 picture, and old-style pictures
    /// in Word 97 files.
    #[test]
    fn test_scan_finds_picf_metafile_picture() {
        let (picf, wmf) = picf_wmf(0);
        let mut data = vec![0x55u8; 37];
        data.extend_from_slice(&picf);
        data.extend(vec![0x55u8; 20]);
        let images = extract_images(&data);
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].format, ImageFormat::Wmf);
        assert_eq!(images[0].data, wmf);
    }

    /// A `PICF`-shaped header over a payload that is not a picture, and an
    /// OfficeArt-shape `PICF` (`MM_SHAPE`), are not picked up here.
    #[test]
    fn test_picf_scan_rejects_non_pictures() {
        let (mut picf, _) = picf_wmf(0);
        picf[0x44 + 2] = 7; // HeaderSize != 9
        assert!(extract_images(&picf).is_empty());
        let (mut shape, _) = picf_wmf(0);
        shape[6..8].copy_from_slice(&0x64u16.to_le_bytes());
        assert!(extract_images(&shape).is_empty());
        // lcb past the end of the buffer.
        let (mut long, _) = picf_wmf(0);
        long[0..4].copy_from_slice(&0x7FFF_FFFFi32.to_le_bytes());
        assert!(extract_images(&long).is_empty());
    }

    #[test]
    fn test_empty_data_stream() {
        assert!(extract_images(&[]).is_empty());
    }
}
