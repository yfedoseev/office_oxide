//! Standard base64 (RFC 4648 §4, `+/` alphabet, `=` padding, no line breaks).
//!
//! Used for image bytes in every JSON surface of the IR and in the
//! `[image-base64:…]` / `data:` URI renderer options. Small enough that a
//! dependency is not worth its audit surface.

const ALPHABET: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

/// Encode `data` (whole 3-byte groups, then the tail) into `out`, which holds
/// `data.len().div_ceil(3) * 4` bytes; returns the number written.
fn encode_chunk(data: &[u8], out: &mut [u8]) -> usize {
    let mut at = 0;
    for chunk in data.chunks(3) {
        let b = [
            chunk[0],
            *chunk.get(1).unwrap_or(&0),
            *chunk.get(2).unwrap_or(&0),
        ];
        let n = ((b[0] as u32) << 16) | ((b[1] as u32) << 8) | b[2] as u32;
        out[at] = ALPHABET[(n >> 18) as usize & 63];
        out[at + 1] = ALPHABET[(n >> 12) as usize & 63];
        out[at + 2] = if chunk.len() > 1 {
            ALPHABET[(n >> 6) as usize & 63]
        } else {
            b'='
        };
        out[at + 3] = if chunk.len() > 2 {
            ALPHABET[n as usize & 63]
        } else {
            b'='
        };
        at += 4;
    }
    at
}

/// Encode `data` with padding and no line breaks (tests; the renderers write
/// through [`Display`]).
#[cfg(test)]
pub(crate) fn encode(data: &[u8]) -> String {
    format!("{}", Display(data))
}

/// `data` as base64 text, produced as it is written: formatting it into a
/// string, a file or a JSON serializer needs no intermediate copy of the
/// whole encoding (for a 30 MiB picture, 40 MiB).
pub(crate) struct Display<'a>(pub &'a [u8]);

impl std::fmt::Display for Display<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        // 768 input bytes make 1,024 output bytes and end on a group boundary.
        let mut buf = [0u8; 1024];
        for part in self.0.chunks(768) {
            let n = encode_chunk(part, &mut buf);
            // The alphabet and `=` are ASCII.
            f.write_str(std::str::from_utf8(&buf[..n]).map_err(|_| std::fmt::Error)?)?;
        }
        Ok(())
    }
}

/// Decode standard base64. Padding is optional, ASCII whitespace is
/// ignored, and any other character outside the alphabet is an error —
/// the IR JSON is a documented interchange format, so a corrupted image
/// field must fail loudly rather than yield plausible garbage.
pub(crate) fn decode(text: &str) -> Result<Vec<u8>, String> {
    fn value(c: u8) -> Option<u32> {
        Some(match c {
            b'A'..=b'Z' => c - b'A',
            b'a'..=b'z' => c - b'a' + 26,
            b'0'..=b'9' => c - b'0' + 52,
            b'+' => 62,
            b'/' => 63,
            _ => return None,
        } as u32)
    }
    let mut out = Vec::with_capacity(text.len() / 4 * 3);
    let mut acc = 0u32;
    let mut bits = 0u32;
    let mut seen_pad = false;
    for (i, &c) in text.as_bytes().iter().enumerate() {
        if c.is_ascii_whitespace() {
            continue;
        }
        if c == b'=' {
            seen_pad = true;
            continue;
        }
        if seen_pad {
            return Err(format!("base64 data after padding at byte {i}"));
        }
        let v = value(c)
            .ok_or_else(|| format!("invalid base64 character {:?} at byte {i}", c as char))?;
        acc = (acc << 6) | v;
        bits += 6;
        if bits >= 8 {
            bits -= 8;
            out.push((acc >> bits) as u8);
            acc &= (1 << bits) - 1;
        }
    }
    if bits >= 6 {
        return Err("base64 input has a dangling sextet".into());
    }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The streaming encoder agrees with the reference vectors of RFC 4648 §10
    /// for every length around a group and across the 768-byte chunk edge.
    #[test]
    fn test_display_matches_rfc_vectors_and_the_chunk_edges() {
        for (plain, coded) in [
            ("", ""),
            ("f", "Zg=="),
            ("fo", "Zm8="),
            ("foo", "Zm9v"),
            ("foob", "Zm9vYg=="),
            ("fooba", "Zm9vYmE="),
            ("foobar", "Zm9vYmFy"),
        ] {
            assert_eq!(encode(plain.as_bytes()), coded);
        }
        for len in [766usize, 767, 768, 769, 770, 1535, 1536, 1537, 5000] {
            let data: Vec<u8> = (0..len).map(|i| (i * 31 % 251) as u8).collect();
            let text = encode(&data);
            assert_eq!(text.len(), len.div_ceil(3) * 4, "length {len}");
            assert_eq!(decode(&text).unwrap(), data, "length {len}");
        }
    }

    #[test]
    fn test_round_trips_every_padding_length() {
        for len in 0..10usize {
            let data: Vec<u8> = (0..len as u8)
                .map(|i| i.wrapping_mul(37).wrapping_add(200))
                .collect();
            let text = encode(&data);
            assert_eq!(text.len() % 4, 0, "padded to a multiple of 4: {text}");
            assert_eq!(decode(&text).unwrap(), data, "{text}");
        }
    }

    #[test]
    fn test_decode_accepts_unpadded_and_whitespace_and_rejects_garbage() {
        assert_eq!(decode("aGk").unwrap(), b"hi");
        assert_eq!(decode("aGk=\n").unwrap(), b"hi");
        assert_eq!(decode("aG\r\nk=").unwrap(), b"hi");
        assert!(decode("aG!k").is_err());
        assert!(decode("aGk=a").is_err());
        assert!(decode("a").is_err());
    }

    #[test]
    fn test_known_vector() {
        assert_eq!(encode(b"Man"), "TWFu");
        assert_eq!(encode(b"Ma"), "TWE=");
        assert_eq!(encode(b"M"), "TQ==");
    }
}
