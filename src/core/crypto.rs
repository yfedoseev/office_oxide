//! The primitives that opening a document protected with Office's fixed
//! default password needs: MD5, SHA-1, SHA-2, AES decryption (ECB and CBC)
//! and RC4.
//!
//! Compiled only with the `default-password` feature, which is off by
//! default. The primitives themselves come from the RustCrypto crates
//! (`md-5`, `sha1`, `sha2`, `aes`, `rc4`), the implementations that
//! `pdf_oxide` and much of the ecosystem use; this module only adapts them
//! to the shapes the Office schemes need, and chains AES blocks in ECB and CBC
//! (a few lines, checked against NIST SP 800-38A). MD5 and RC4 are not
//! approved for use in a FIPS 140 validated module; they are here only because
//! the legacy Office file formats are defined in terms of them
//! ([MS-OFFCRYPTO] §2.3), and only to *decrypt* with the well-known default
//! password, never to protect anything.
//!
//! Each primitive is checked against its published test vectors.

use aes::cipher::{BlockDecrypt, KeyInit, generic_array::GenericArray};
use md5::Digest;

pub(crate) fn md5(data: &[u8]) -> [u8; 16] {
    md5::Md5::digest(data).into()
}

pub(crate) fn sha1(data: &[u8]) -> [u8; 20] {
    sha1::Sha1::digest(data).into()
}

pub(crate) fn sha256(data: &[u8]) -> [u8; 32] {
    sha2::Sha256::digest(data).into()
}

pub(crate) fn sha384(data: &[u8]) -> [u8; 48] {
    sha2::Sha384::digest(data).into()
}

pub(crate) fn sha512(data: &[u8]) -> [u8; 64] {
    sha2::Sha512::digest(data).into()
}

/// A hash function named by a file.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Hash {
    Sha1,
    Sha256,
    Sha384,
    Sha512,
}

impl Hash {
    /// The digest in a fixed buffer, and its length: no allocation, for the
    /// iterated hashing of a password.
    pub(crate) fn digest_array(self, data: &[u8]) -> ([u8; 64], usize) {
        let mut out = [0u8; 64];
        let len = match self {
            Self::Sha1 => {
                out[..20].copy_from_slice(&sha1(data));
                20
            },
            Self::Sha256 => {
                out[..32].copy_from_slice(&sha256(data));
                32
            },
            Self::Sha384 => {
                out[..48].copy_from_slice(&sha384(data));
                48
            },
            Self::Sha512 => {
                out = sha512(data);
                64
            },
        };
        (out, len)
    }

    pub(crate) fn digest(self, data: &[u8]) -> Vec<u8> {
        match self {
            Self::Sha1 => sha1(data).to_vec(),
            Self::Sha256 => sha256(data).to_vec(),
            Self::Sha384 => sha384(data).to_vec(),
            Self::Sha512 => sha512(data).to_vec(),
        }
    }
}

/// An AES key schedule of any of the three key sizes; it decrypts ([FIPS 197]),
/// and in tests encrypts, so a round trip can check the schemes built on it.
pub(crate) enum Aes {
    Aes128(aes::Aes128),
    Aes192(aes::Aes192),
    Aes256(aes::Aes256),
}

impl Aes {
    /// `None` unless `key` is 16, 24 or 32 bytes.
    pub(crate) fn new(key: &[u8]) -> Option<Self> {
        match key.len() {
            16 => aes::Aes128::new_from_slice(key).ok().map(Self::Aes128),
            24 => aes::Aes192::new_from_slice(key).ok().map(Self::Aes192),
            32 => aes::Aes256::new_from_slice(key).ok().map(Self::Aes256),
            _ => None,
        }
    }

    fn decrypt_block(&self, block: &mut [u8; 16]) {
        let block = GenericArray::from_mut_slice(block);
        match self {
            Self::Aes128(c) => c.decrypt_block(block),
            Self::Aes192(c) => c.decrypt_block(block),
            Self::Aes256(c) => c.decrypt_block(block),
        }
    }

    #[cfg(test)]
    pub(crate) fn encrypt_block(&self, block: &mut [u8; 16]) {
        use aes::cipher::BlockEncrypt;
        let block = GenericArray::from_mut_slice(block);
        match self {
            Self::Aes128(c) => c.encrypt_block(block),
            Self::Aes192(c) => c.encrypt_block(block),
            Self::Aes256(c) => c.encrypt_block(block),
        }
    }

    /// Encrypt whole blocks in ECB mode (tests only).
    #[cfg(test)]
    pub(crate) fn encrypt_ecb(&self, data: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        for chunk in data.as_chunks::<16>().0 {
            let mut block = [0u8; 16];
            block.copy_from_slice(chunk);
            self.encrypt_block(&mut block);
            out.extend_from_slice(&block);
        }
        out
    }

    /// Encrypt whole blocks in CBC mode from `iv` (tests only).
    #[cfg(test)]
    pub(crate) fn encrypt_cbc(&self, iv: &[u8; 16], data: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        let mut previous = *iv;
        for chunk in data.as_chunks::<16>().0 {
            let mut block = [0u8; 16];
            for (b, (c, p)) in block.iter_mut().zip(chunk.iter().zip(previous)) {
                *b = c ^ p;
            }
            self.encrypt_block(&mut block);
            out.extend_from_slice(&block);
            previous = block;
        }
        out
    }

    /// Decrypt `data` (whole 16-byte blocks; a trailing part block is dropped)
    /// in ECB mode.
    pub(crate) fn decrypt_ecb(&self, data: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(data.len() - data.len() % 16);
        for chunk in data.as_chunks::<16>().0 {
            let mut block = [0u8; 16];
            block.copy_from_slice(chunk);
            self.decrypt_block(&mut block);
            out.extend_from_slice(&block);
        }
        out
    }

    /// Decrypt `data` (whole blocks) in CBC mode from `iv`.
    pub(crate) fn decrypt_cbc(&self, iv: &[u8; 16], data: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(data.len() - data.len() % 16);
        let mut previous = *iv;
        for chunk in data.as_chunks::<16>().0 {
            let mut block = [0u8; 16];
            block.copy_from_slice(chunk);
            let cipher = block;
            self.decrypt_block(&mut block);
            for (b, p) in block.iter_mut().zip(previous) {
                *b ^= p;
            }
            out.extend_from_slice(&block);
            previous = cipher;
        }
        out
    }
}

/// The RC4 stream cipher, for a key of 1 to 16 bytes (Office uses 5 to 16).
pub(crate) struct Rc4(Rc4Key);

macro_rules! rc4_keys {
    ($($n:literal => $variant:ident / $size:ident),* $(,)?) => {
        enum Rc4Key { $($variant(rc4::Rc4<rc4::consts::$size>)),* }

        impl Rc4 {
            /// `None` for an empty key or one over 16 bytes.
            pub(crate) fn new(key: &[u8]) -> Option<Self> {
                match key.len() {
                    $($n => rc4::Rc4::<rc4::consts::$size>::new_from_slice(key)
                        .ok()
                        .map(|c| Self(Rc4Key::$variant(c))),)*
                    _ => None,
                }
            }

            /// XOR the next keystream bytes into `data`.
            pub(crate) fn apply(&mut self, data: &mut [u8]) {
                use rc4::StreamCipher;
                match &mut self.0 {
                    $(Rc4Key::$variant(c) => c.apply_keystream(data)),*
                }
            }
        }
    };
}

rc4_keys!(
    1 => K1 / U1, 2 => K2 / U2, 3 => K3 / U3, 4 => K4 / U4, 5 => K5 / U5, 6 => K6 / U6,
    7 => K7 / U7, 8 => K8 / U8, 9 => K9 / U9, 10 => K10 / U10, 11 => K11 / U11,
    12 => K12 / U12, 13 => K13 / U13, 14 => K14 / U14, 15 => K15 / U15, 16 => K16 / U16,
);

#[cfg(test)]
mod tests {
    use super::*;

    fn hex(bytes: &[u8]) -> String {
        bytes.iter().map(|b| format!("{b:02x}")).collect()
    }

    fn unhex(text: &str) -> Vec<u8> {
        (0..text.len())
            .step_by(2)
            .map(|i| u8::from_str_radix(&text[i..i + 2], 16).unwrap())
            .collect()
    }

    #[test]
    fn test_md5_known_answers() {
        assert_eq!(hex(&md5(b"")), "d41d8cd98f00b204e9800998ecf8427e");
        assert_eq!(hex(&md5(b"abc")), "900150983cd24fb0d6963f7d28e17f72");
        assert_eq!(
            hex(&md5(
                b"12345678901234567890123456789012345678901234567890123456789012345678901234567890"
            )),
            "57edf4a22be3c955ac49da2e2107b67a"
        );
    }

    #[test]
    fn test_sha_known_answers() {
        assert_eq!(hex(&sha1(b"abc")), "a9993e364706816aba3e25717850c26c9cd0d89d");
        assert_eq!(hex(&sha1(b"")), "da39a3ee5e6b4b0d3255bfef95601890afd80709");
        assert_eq!(
            hex(&sha256(b"abc")),
            "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
        );
        assert_eq!(
            hex(&sha384(b"abc")),
            "cb00753f45a35e8bb5a03d699ac65007272c32ab0eded1631a8b605a43ff5bed8086072ba1e7cc2358baeca134c825a7"
        );
        assert_eq!(
            hex(&sha512(b"abc")),
            "ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f"
        );
    }

    #[test]
    fn test_hashes_are_right_across_padding_boundaries() {
        // The digest of the digests of `n` bytes of 0xA5, for `n` around every
        // block edge of the 64- and 128-byte families, as an independent
        // implementation computes it.
        let lens = [
            0usize, 1, 55, 56, 63, 64, 65, 111, 112, 119, 120, 127, 128, 129, 255, 256,
        ];
        fn chain(hash: fn(&[u8]) -> Vec<u8>, lens: &[usize]) -> String {
            let all: Vec<u8> = lens.iter().flat_map(|&n| hash(&vec![0xa5u8; n])).collect();
            hex(&hash(&all))
        }
        assert_eq!(chain(|d| md5(d).to_vec(), &lens), "fc3cd4142fdd7090de6774db3d32c6d5");
        assert_eq!(chain(|d| sha1(d).to_vec(), &lens), "d51d0540bc6ddd75cdca66c5bb2968dcd839f8c5");
        assert_eq!(
            chain(|d| sha256(d).to_vec(), &lens),
            "4f707468ac426bc1505f6159093ebcb1e215f651350cb3b0fe3f43d0dd586fbd"
        );
        assert_eq!(
            chain(|d| sha384(d).to_vec(), &lens),
            "d1649d6c1ea2c81be8ee25ea09d3aebcafc6fa9eeb4d3b3a35b4bd16014d61e6bbe4f0b86374ae43b159a05306c4ee6d"
        );
        assert_eq!(
            chain(|d| sha512(d).to_vec(), &lens),
            "b9f24310347400894b8c5504b860ba31fb89febcb6d6452a8f7f50a7c078ad40ea6fcc1334555aa7aa1f87e2456da1b93ad43a529f93c052e63ec2f5ee98488e"
        );
    }

    #[test]
    fn test_aes_decrypts_the_published_examples() {
        let plain = unhex("00112233445566778899aabbccddeeff");
        for (key, cipher) in [
            ("000102030405060708090a0b0c0d0e0f", "69c4e0d86a7b0430d8cdb78070b4c55a"),
            (
                "000102030405060708090a0b0c0d0e0f1011121314151617",
                "dda97ca4864cdfe06eaf70a0ec0d7191",
            ),
            (
                "000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f",
                "8ea2b7ca516745bfeafc49904b496089",
            ),
        ] {
            let aes = Aes::new(&unhex(key)).unwrap();
            assert_eq!(aes.decrypt_ecb(&unhex(cipher)), plain, "key {key}");
            assert_eq!(aes.encrypt_ecb(&plain), unhex(cipher), "key {key}");
        }
        assert!(Aes::new(&[0; 15]).is_none());
    }

    #[test]
    fn test_aes_cbc_chains_the_blocks() {
        // NIST SP 800-38A F.2.2, CBC-AES128.Decrypt.
        let aes = Aes::new(&unhex("2b7e151628aed2a6abf7158809cf4f3c")).unwrap();
        let iv: [u8; 16] = unhex("000102030405060708090a0b0c0d0e0f")
            .try_into()
            .unwrap();
        let cipher = unhex("7649abac8119b246cee98e9b12e9197d5086cb9b507219ee95db113a917678b2");
        assert_eq!(
            hex(&aes.decrypt_cbc(&iv, &cipher)),
            "6bc1bee22e409f96e93d7e117393172aae2d8a571e03ac9c9eb76fac45af8e51"
        );
    }

    #[test]
    fn test_rc4_known_answers() {
        // The classic "Key" / "Plaintext" vector.
        let mut data = b"Plaintext".to_vec();
        Rc4::new(b"Key").unwrap().apply(&mut data);
        assert_eq!(hex(&data), "bbf316e8d940af0ad3");
        assert!(Rc4::new(&[]).is_none());
    }
}
