//! Picture bytes that are shared, and optionally inflated only when used.
//!
//! A picture in an OOXML package is a zip entry, often deflated: a page image
//! written as an uncompressed PNG is stored at 3% of its size. Reading every
//! picture at open held them all inflated for the life of the document,
//! although text, Markdown and HTML without images never read them, and the
//! IR then copied each into every `Image` that drew it.
//!
//! A [`PictureBytes`] is a cheap, shared handle to one picture's bytes. It
//! holds them either inflated, or as the entry's stored deflate stream with
//! the length and CRC-32 the bytes had when the part was read at open, in
//! which case it inflates them each time they are asked for and keeps
//! nothing: the inflated copy belongs to the caller, so a conversion that
//! writes one picture at a time holds one picture at a time. The stream was
//! read and checked at open, through the same limits as every other part,
//! so inflating it again cannot be refused for a size the open allowed; a
//! result that differs from what the open read (the stream changed under the
//! handle) is reported as no bytes, never as different ones.

use std::io::Read;
use std::sync::Arc;

/// One picture's bytes, shared; see the [module documentation](self).
#[derive(Clone)]
pub struct PictureBytes(Arc<Source>);

enum Source {
    /// The bytes themselves.
    Inflated(Arc<[u8]>),
    /// A raw deflate stream ([RFC 1951]) that inflates to `len` bytes whose
    /// CRC-32 is `crc`.
    ///
    /// [RFC 1951]: https://www.rfc-editor.org/rfc/rfc1951
    Deflated {
        stored: Box<[u8]>,
        len: usize,
        crc: u32,
    },
}

impl PictureBytes {
    /// A handle that inflates `stored`, a raw deflate stream, on use; `bytes`
    /// are what it inflated to when the part was read, used here only for
    /// their length and checksum.
    pub(crate) fn deflated(stored: Vec<u8>, bytes: &[u8]) -> Self {
        let mut crc = flate2::Crc::new();
        crc.update(bytes);
        Self(Arc::new(Source::Deflated {
            stored: stored.into_boxed_slice(),
            len: bytes.len(),
            crc: crc.sum(),
        }))
    }

    /// The number of bytes, without inflating them.
    pub fn len(&self) -> usize {
        match &*self.0 {
            Source::Inflated(b) => b.len(),
            Source::Deflated { len, .. } => *len,
        }
    }

    /// Whether there are no bytes.
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }

    /// The bytes held in memory now: the stored stream of a picture that
    /// inflates on use, the bytes of one that does not.
    pub fn resident_len(&self) -> usize {
        match &*self.0 {
            Source::Inflated(b) => b.len(),
            Source::Deflated { stored, .. } => stored.len(),
        }
    }

    /// The bytes: shared when they are held inflated, inflated afresh (and
    /// not kept) when they are not. `None` only if the stored stream no
    /// longer inflates to what it did at open.
    pub fn bytes(&self) -> Option<Arc<[u8]>> {
        match &*self.0 {
            Source::Inflated(b) => Some(b.clone()),
            Source::Deflated { .. } => self.inflate().map(Into::into),
        }
    }

    /// The bytes as a `Vec` of their own: copied when held inflated, and
    /// when not, the inflated buffer itself, with no copy made.
    pub(crate) fn to_vec(&self) -> Option<Vec<u8>> {
        match &*self.0 {
            Source::Inflated(b) => Some(b.to_vec()),
            Source::Deflated { .. } => self.inflate(),
        }
    }

    /// Inflate the stored stream, checking it gives the bytes read at open.
    fn inflate(&self) -> Option<Vec<u8>> {
        let Source::Deflated { stored, len, crc } = &*self.0 else {
            return None;
        };
        let mut out = Vec::with_capacity(*len);
        let mut decoder = flate2::read::DeflateDecoder::new(&stored[..]);
        (&mut decoder)
            .take(*len as u64)
            .read_to_end(&mut out)
            .ok()?;
        let mut check = flate2::Crc::new();
        check.update(&out);
        if out.len() != *len || check.sum() != *crc {
            log::warn!("picture: the stored stream no longer inflates to the bytes read at open");
            return None;
        }
        Some(out)
    }

    /// Whether `self` and `other` are the same picture (one allocation), not
    /// merely equal bytes.
    pub fn ptr_eq(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
    }

    /// Whether the bytes are held as a stored stream and inflated on use.
    pub fn is_deflated(&self) -> bool {
        matches!(&*self.0, Source::Deflated { .. })
    }
}

impl From<Vec<u8>> for PictureBytes {
    fn from(bytes: Vec<u8>) -> Self {
        Self(Arc::new(Source::Inflated(bytes.into())))
    }
}

impl From<Arc<[u8]>> for PictureBytes {
    fn from(bytes: Arc<[u8]>) -> Self {
        Self(Arc::new(Source::Inflated(bytes)))
    }
}

impl From<&[u8]> for PictureBytes {
    fn from(bytes: &[u8]) -> Self {
        Self(Arc::new(Source::Inflated(bytes.into())))
    }
}

impl std::fmt::Debug for PictureBytes {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("PictureBytes")
            .field("len", &self.len())
            .field("resident", &self.resident_len())
            .finish()
    }
}

/// Equal bytes are equal pictures, however each is held.
impl PartialEq for PictureBytes {
    fn eq(&self, other: &Self) -> bool {
        self.ptr_eq(other)
            || (self.len() == other.len()
                && match (self.bytes(), other.bytes()) {
                    (Some(a), Some(b)) => a == b,
                    _ => false,
                })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    fn deflate(data: &[u8]) -> Vec<u8> {
        let mut e = flate2::write::DeflateEncoder::new(Vec::new(), flate2::Compression::default());
        e.write_all(data).unwrap();
        e.finish().unwrap()
    }

    #[test]
    fn test_a_deflated_picture_inflates_to_its_bytes_and_keeps_only_the_stream() {
        let data: Vec<u8> = (0..100_000u32).map(|i| (i % 7) as u8).collect();
        let stored = deflate(&data);
        let pic = PictureBytes::deflated(stored.clone(), &data);
        assert!(pic.is_deflated());
        assert_eq!(pic.len(), data.len());
        assert_eq!(pic.resident_len(), stored.len());
        assert!(pic.resident_len() * 20 < data.len());
        assert_eq!(&*pic.bytes().unwrap(), &data[..]);
        // Inflating again gives the same bytes; nothing was kept.
        assert_eq!(&*pic.bytes().unwrap(), &data[..]);
        assert_eq!(pic.resident_len(), stored.len());
        assert_eq!(pic, PictureBytes::from(data.clone()));
    }

    #[test]
    fn test_a_stream_that_does_not_match_its_checksum_gives_no_bytes() {
        let data = vec![1u8; 4096];
        let other = vec![2u8; 4096];
        let pic = PictureBytes::deflated(deflate(&other), &data);
        assert_eq!(pic.len(), 4096);
        assert!(pic.bytes().is_none(), "different bytes must not be handed out");
        let truncated = PictureBytes::deflated(deflate(&data)[..10].to_vec(), &data);
        assert!(truncated.bytes().is_none());
    }

    #[test]
    fn test_clones_share_one_allocation() {
        let a = PictureBytes::from(vec![9u8; 10]);
        let b = a.clone();
        assert!(a.ptr_eq(&b));
        assert!(!a.ptr_eq(&PictureBytes::from(vec![9u8; 10])));
        assert_eq!(a, PictureBytes::from(vec![9u8; 10]));
        assert!(PictureBytes::from(Vec::new()).is_empty());
    }
}
