//! OfficeArt BLIP (image) record parsing.
//!
//! BLIP records are used in PPT Pictures streams and DOC Data streams
//! to store embedded images (JPEG, PNG, EMF, WMF, etc.).
//! Each BLIP has an 8-byte OfficeArt record header followed by a UID
//! and raw image data.

/// Image format stored in a BLIP record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BlipFormat {
    /// Enhanced Metafile (EMF).
    Emf,
    /// Windows Metafile (WMF).
    Wmf,
    /// Apple PICT image.
    Pict,
    /// JPEG image.
    Jpeg,
    /// PNG image.
    Png,
    /// Device-Independent Bitmap (DIB/BMP).
    Dib,
    /// TIFF image.
    Tiff,
    /// Unrecognized record type.
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
    /// This image's 0-based position in the stream's own top-level
    /// array of `OfficeArtBStoreContainerFileBlock` entries — i.e. the
    /// same array position an `OfficeArtFOPT` shape's `pib` ("Blip to
    /// display") property references, per [MS-ODRAW]. Every top-level
    /// entry counts toward this position, including an `OfficeArtFBSE`
    /// wrapper entry that didn't yield an image — the
    /// index is *not* simply "how many images have been extracted so
    /// far", since that would drift out of alignment with `pib` as soon
    /// as any entry is skipped.
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

/// `OfficeArtFBSE` ("File BLIP Store Entry", [MS-ODRAW] §2.2.32):
/// `btWin32`(1) + `btMacOS`(1) + `rgbUid`(16) + `tag`(2) + `size`(4) +
/// `cRef`(4) + `foDelay`(4) + `unused1`(1) + `cbName`(1) + `unused2`(1) +
/// `unused3`(1) = 36 fixed bytes, then `nameData` (`cbName` bytes), then
/// an optional embedded `OfficeArtBlip`.
const FBSE_FIXED_SIZE: usize = 36;
/// Offset of `cbName` within an FBSE record's body (right after `rh`).
const FBSE_CBNAME_OFFSET: usize = 33;

/// Extract a single `OfficeArtBlip` record's image bytes, given its own
/// record header fields already decoded.
fn extract_one_blip(
    data: &[u8],
    rec_type: u16,
    inst: u16,
    data_start: usize,
    data_end: usize,
) -> Option<BlipImage> {
    let format = BlipFormat::from_record_type(rec_type);
    if !format.is_image() {
        return None;
    }
    let skip = uid_size(rec_type, inst) + metafile_header_size(rec_type);
    let img_start = data_start + skip;
    if img_start >= data_end {
        return None;
    }
    let img_data = &data[img_start..data_end];
    if img_data.is_empty() {
        return None;
    }
    Some(BlipImage {
        format,
        data: img_data.to_vec(),
        index: 0,
    })
}

/// Extract all BLIP images from an OfficeArt data stream.
///
/// Works for both PPT Pictures streams and DOC Data streams. Each
/// top-level entry — a raw `OfficeArtBlip`, an `OfficeArtFBSE` wrapper
///, or anything unrecognized — counts as one array slot
/// toward [`BlipImage::index`], whether or not it actually yielded an
/// image; only "descend into a container" steps into the array rather
/// than past a sibling of it, so those don't count.
pub fn extract_blip_images(data: &[u8]) -> Vec<BlipImage> {
    let mut images = Vec::new();
    let mut pos = 0;
    let mut slot_index = 0usize;

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

        if rec_type == 0xF007 {
            // OfficeArtFBSE: the actual pixel data (if present at all,
            // rather than referencing PowerPoint's separate delay
            // stream) is an OfficeArtBlip nested past the fixed header
            // and name field.
            if data_start + FBSE_FIXED_SIZE <= data_end {
                let cb_name = data[data_start + FBSE_CBNAME_OFFSET] as usize;
                let embedded_start = data_start + FBSE_FIXED_SIZE + cb_name;
                if embedded_start + 8 <= data_end {
                    let e_ver_inst =
                        u16::from_le_bytes([data[embedded_start], data[embedded_start + 1]]);
                    let e_rec_type =
                        u16::from_le_bytes([data[embedded_start + 2], data[embedded_start + 3]]);
                    let e_rec_len = u32::from_le_bytes([
                        data[embedded_start + 4],
                        data[embedded_start + 5],
                        data[embedded_start + 6],
                        data[embedded_start + 7],
                    ]) as usize;
                    let e_inst = e_ver_inst >> 4;
                    let e_data_start = embedded_start + 8;
                    let e_data_end = (e_data_start + e_rec_len).min(data_end);
                    if let Some(mut img) =
                        extract_one_blip(data, e_rec_type, e_inst, e_data_start, e_data_end)
                    {
                        img.index = slot_index;
                        images.push(img);
                    }
                }
            }
            slot_index += 1;
            pos = data_end;
        } else if let Some(mut img) = extract_one_blip(data, rec_type, inst, data_start, data_end) {
            img.index = slot_index;
            images.push(img);
            slot_index += 1;
            pos = data_end;
        } else if ver == 0x0F {
            // Container record — descend into children, not a sibling
            // array entry, so no slot is consumed here.
            pos = data_start;
        } else {
            // Non-BLIP atom — skip over it, but it still occupies a slot.
            slot_index += 1;
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

    /// Build an `OfficeArtFBSE` record wrapping the given
    /// embedded `OfficeArtBlip` bytes (from [`make_blip`]).
    fn make_fbse(name: &str, embedded_blip: &[u8]) -> Vec<u8> {
        let name_utf16: Vec<u8> = name
            .encode_utf16()
            .chain(std::iter::once(0u16))
            .flat_map(u16::to_le_bytes)
            .collect();
        let mut body = Vec::new();
        body.push(0); // btWin32
        body.push(0); // btMacOS
        body.extend(vec![0u8; 16]); // rgbUid
        body.extend_from_slice(&0xFFu16.to_le_bytes()); // tag
        body.extend_from_slice(&(embedded_blip.len() as u32).to_le_bytes()); // size
        body.extend_from_slice(&1u32.to_le_bytes()); // cRef
        body.extend_from_slice(&0xFFFFFFFFu32.to_le_bytes()); // foDelay
        body.push(0); // unused1
        body.push(name_utf16.len() as u8); // cbName
        body.push(0); // unused2
        body.push(0); // unused3
        body.extend_from_slice(&name_utf16);
        body.extend_from_slice(embedded_blip);

        let ver_inst: u16 = 0x2; // rh.recVer MUST be 0x2 for FBSE
        let mut buf = Vec::new();
        buf.extend_from_slice(&ver_inst.to_le_bytes());
        buf.extend_from_slice(&0xF007u16.to_le_bytes());
        buf.extend_from_slice(&(body.len() as u32).to_le_bytes());
        buf.extend(body);
        buf
    }

    #[test]
    fn test_fbse_wrapped_blip_is_extracted() {
        let jpeg = make_blip(0xF01D, 0x46A, b"\xff\xd8\xff\xe0FBSE_JPEG");
        let stream = make_fbse("pic.jpg", &jpeg);
        let images = extract_blip_images(&stream);
        assert_eq!(images.len(), 1, "the embedded blip inside an FBSE wrapper must be extracted");
        assert_eq!(images[0].format, BlipFormat::Jpeg);
        assert_eq!(images[0].data, b"\xff\xd8\xff\xe0FBSE_JPEG");
    }

    #[test]
    fn test_fbse_entry_before_a_raw_blip_does_not_shift_its_index() {
        // FBSE entry (slot 0) + a raw blip (slot 1) — the raw blip's
        // index must be 1, matching its real array position, not 0
        // (which is what "count of successfully extracted images so
        // far" would have produced before this was fixed).
        let jpeg = make_blip(0xF01D, 0x46A, b"\xff\xd8\xff\xe0FBSE_JPEG");
        let mut stream = make_fbse("pic.jpg", &jpeg);
        stream.extend(make_blip(0xF01E, 0x6E0, b"\x89PNGPNG_RAW"));

        let images = extract_blip_images(&stream);
        assert_eq!(images.len(), 2);
        assert_eq!(images[0].index, 0);
        assert_eq!(images[1].index, 1);
        assert_eq!(images[1].format, BlipFormat::Png);
    }

    #[test]
    fn test_empty_fbse_slot_still_advances_the_index() {
        // cRef=0 (an "empty slot" per spec) with no embedded blip at
        // all — it must still consume slot 0 so the following real
        // image correctly reports index 1.
        let mut body = Vec::new();
        body.push(0);
        body.push(0);
        body.extend(vec![0u8; 16]);
        body.extend_from_slice(&0xFFu16.to_le_bytes());
        body.extend_from_slice(&0u32.to_le_bytes()); // size
        body.extend_from_slice(&0u32.to_le_bytes()); // cRef = 0 (empty slot)
        body.extend_from_slice(&0xFFFFFFFFu32.to_le_bytes());
        body.push(0);
        body.push(0); // cbName = 0, no name
        body.push(0);
        body.push(0);
        let mut empty_fbse = Vec::new();
        empty_fbse.extend_from_slice(&0x2u16.to_le_bytes());
        empty_fbse.extend_from_slice(&0xF007u16.to_le_bytes());
        empty_fbse.extend_from_slice(&(body.len() as u32).to_le_bytes());
        empty_fbse.extend(body);

        let mut stream = empty_fbse;
        stream.extend(make_blip(0xF01D, 0x46A, b"\xff\xd8\xff\xe0AFTER_EMPTY"));

        let images = extract_blip_images(&stream);
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].index, 1);
    }

    #[test]
    fn test_extract_jpeg() {
        let jpeg_data = b"\xff\xd8\xff\xe0JFIF_DATA";
        let stream = make_blip(0xF01D, 0x46A, jpeg_data);
        let images = extract_blip_images(&stream);
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].format, BlipFormat::Jpeg);
        assert_eq!(images[0].data, jpeg_data);
    }

    #[test]
    fn test_extract_png() {
        let png_data = b"\x89PNG\r\n\x1a\nIHDR_DATA";
        let stream = make_blip(0xF01E, 0x6E0, png_data);
        let images = extract_blip_images(&stream);
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].format, BlipFormat::Png);
        assert_eq!(images[0].data, png_data);
    }

    #[test]
    fn test_extract_multiple() {
        let mut stream = make_blip(0xF01D, 0x46A, b"\xff\xd8\xff\xe0JPEG1");
        stream.extend(make_blip(0xF01E, 0x6E0, b"\x89PNGPNG2"));
        let images = extract_blip_images(&stream);
        assert_eq!(images.len(), 2);
        assert_eq!(images[0].format, BlipFormat::Jpeg);
        assert_eq!(images[1].format, BlipFormat::Png);
        assert_eq!(images[1].index, 1);
    }

    #[test]
    fn test_extract_with_secondary_uid() {
        let jpeg_data = b"\xff\xd8\xff\xe0TEST";
        let stream = make_blip(0xF01D, 0x46B, jpeg_data); // bit 0 set
        let images = extract_blip_images(&stream);
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].data, jpeg_data);
    }

    #[test]
    fn test_skips_container_records() {
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
    fn test_empty_stream() {
        assert!(extract_blip_images(&[]).is_empty());
    }

    #[test]
    fn test_format_metadata() {
        assert_eq!(BlipFormat::Jpeg.extension(), "jpg");
        assert_eq!(BlipFormat::Png.mime_type(), "image/png");
        assert!(BlipFormat::Jpeg.is_image());
        assert!(!BlipFormat::Unknown(0x1234).is_image());
    }
}
