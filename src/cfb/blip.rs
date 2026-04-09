//! OfficeArt BLIP (image) record parsing.
//!
//! BLIP records are used in PPT Pictures streams and DOC Data streams
//! to store embedded images (JPEG, PNG, EMF, WMF, etc.).
//! Each BLIP has an 8-byte OfficeArt record header followed by a UID
//! and raw image data.

/// Image format stored in a BLIP record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BlipFormat {
    Emf,
    Wmf,
    Pict,
    Jpeg,
    Png,
    Dib,
    Tiff,
    Unknown(u16),
}

impl BlipFormat {
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

    /// Returns true if this is a recognized image BLIP type.
    pub fn is_image(&self) -> bool {
        !matches!(self, Self::Unknown(_))
    }

    /// File extension for this format.
    pub fn extension(&self) -> &'static str {
        match self {
            Self::Emf => "emf",
            Self::Wmf => "wmf",
            Self::Pict => "pict",
            Self::Jpeg => "jpg",
            Self::Png => "png",
            Self::Dib => "bmp",
            Self::Tiff => "tiff",
            Self::Unknown(_) => "bin",
        }
    }

    /// MIME type for this format.
    pub fn mime_type(&self) -> &'static str {
        match self {
            Self::Emf => "image/x-emf",
            Self::Wmf => "image/x-wmf",
            Self::Pict => "image/x-pict",
            Self::Jpeg => "image/jpeg",
            Self::Png => "image/png",
            Self::Dib => "image/bmp",
            Self::Tiff => "image/tiff",
            Self::Unknown(_) => "application/octet-stream",
        }
    }
}

/// An extracted image from a BLIP record.
#[derive(Debug, Clone)]
pub struct BlipImage {
    /// The image format.
    pub format: BlipFormat,
    /// Raw image data.
    pub data: Vec<u8>,
    /// Index of this image in the stream (0-based).
    pub index: usize,
}

/// UID header size for each BLIP type.
fn uid_size(rec_type: u16, inst: u16) -> usize {
    let base = match rec_type {
        0xF01A..=0xF01C => 16, // Metafiles: 16 bytes UID only
        _ => 17,               // Bitmaps: 16 bytes UID + 1 byte tag
    };
    // If inst bit 0 is set, there's a secondary UID (16 more bytes).
    if inst & 1 != 0 { base + 16 } else { base }
}

/// Extra header size for metafile BLIPs (EMF/WMF/PICT).
fn metafile_header_size(rec_type: u16) -> usize {
    match rec_type {
        0xF01A..=0xF01C => 34,
        _ => 0,
    }
}

/// Extract all BLIP images from an OfficeArt data stream.
///
/// Works for both PPT Pictures streams and DOC Data streams.
pub fn extract_blip_images(data: &[u8]) -> Vec<BlipImage> {
    let mut images = Vec::new();
    let mut pos = 0;

    while pos + 8 <= data.len() {
        let ver_inst = u16::from_le_bytes([data[pos], data[pos + 1]]);
        let rec_type = u16::from_le_bytes([data[pos + 2], data[pos + 3]]);
        let rec_len =
            u32::from_le_bytes([data[pos + 4], data[pos + 5], data[pos + 6], data[pos + 7]])
                as usize;

        let ver = ver_inst & 0x0F;
        let inst = ver_inst >> 4;
        let data_start = pos + 8;
        let data_end = (data_start + rec_len).min(data.len());

        let format = BlipFormat::from_record_type(rec_type);

        if format.is_image() {
            let skip = uid_size(rec_type, inst) + metafile_header_size(rec_type);
            let img_start = data_start + skip;
            if img_start < data_end {
                let img_data = &data[img_start..data_end];
                if !img_data.is_empty() {
                    images.push(BlipImage {
                        format,
                        data: img_data.to_vec(),
                        index: images.len(),
                    });
                }
            }
            pos = data_end;
        } else if ver == 0x0F {
            // Container record — descend into children.
            pos = data_start;
        } else {
            // Non-BLIP atom — skip over it.
            pos = data_end;
        }
    }

    images
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_blip(rec_type: u16, inst: u16, image_data: &[u8]) -> Vec<u8> {
        let ver_inst: u16 = inst << 4;
        let uid_sz = uid_size(rec_type, inst);
        let mf_sz = metafile_header_size(rec_type);
        let rec_len = uid_sz + mf_sz + image_data.len();

        let mut buf = Vec::new();
        buf.extend_from_slice(&ver_inst.to_le_bytes());
        buf.extend_from_slice(&rec_type.to_le_bytes());
        buf.extend_from_slice(&(rec_len as u32).to_le_bytes());
        buf.extend(vec![0u8; uid_sz]);
        buf.extend(vec![0u8; mf_sz]);
        buf.extend_from_slice(image_data);
        buf
    }

    #[test]
    fn extract_jpeg() {
        let jpeg_data = b"\xff\xd8\xff\xe0JFIF_DATA";
        let stream = make_blip(0xF01D, 0x46A, jpeg_data);
        let images = extract_blip_images(&stream);
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].format, BlipFormat::Jpeg);
        assert_eq!(images[0].data, jpeg_data);
    }

    #[test]
    fn extract_png() {
        let png_data = b"\x89PNG\r\n\x1a\nIHDR_DATA";
        let stream = make_blip(0xF01E, 0x6E0, png_data);
        let images = extract_blip_images(&stream);
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].format, BlipFormat::Png);
        assert_eq!(images[0].data, png_data);
    }

    #[test]
    fn extract_multiple() {
        let mut stream = make_blip(0xF01D, 0x46A, b"\xff\xd8\xff\xe0JPEG1");
        stream.extend(make_blip(0xF01E, 0x6E0, b"\x89PNGPNG2"));
        let images = extract_blip_images(&stream);
        assert_eq!(images.len(), 2);
        assert_eq!(images[0].format, BlipFormat::Jpeg);
        assert_eq!(images[1].format, BlipFormat::Png);
        assert_eq!(images[1].index, 1);
    }

    #[test]
    fn extract_with_secondary_uid() {
        let jpeg_data = b"\xff\xd8\xff\xe0TEST";
        let stream = make_blip(0xF01D, 0x46B, jpeg_data); // bit 0 set
        let images = extract_blip_images(&stream);
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].data, jpeg_data);
    }

    #[test]
    fn skips_container_records() {
        // Container (ver=0xF) wrapping a BLIP
        let blip = make_blip(0xF01D, 0x46A, b"\xff\xd8\xff\xe0test");
        let mut stream = Vec::new();
        // Container header
        let ver_inst: u16 = 0x0F;
        stream.extend_from_slice(&ver_inst.to_le_bytes());
        stream.extend_from_slice(&0xF000u16.to_le_bytes()); // some container type
        stream.extend_from_slice(&(blip.len() as u32).to_le_bytes());
        stream.extend(&blip);

        let images = extract_blip_images(&stream);
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].format, BlipFormat::Jpeg);
    }

    #[test]
    fn empty_stream() {
        assert!(extract_blip_images(&[]).is_empty());
    }

    #[test]
    fn format_metadata() {
        assert_eq!(BlipFormat::Jpeg.extension(), "jpg");
        assert_eq!(BlipFormat::Png.mime_type(), "image/png");
        assert!(BlipFormat::Jpeg.is_image());
        assert!(!BlipFormat::Unknown(0x1234).is_image());
    }
}
