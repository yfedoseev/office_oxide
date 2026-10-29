//! Opening a document protected with Office's fixed default password.
//!
//! Excel (and Word) encrypt a file that is only structure- or write-protected
//! with the well-known password `VelvetSweatshop`, so it opens without a
//! prompt; some producers use an empty password. Those two passwords, and no
//! other, are tried ([MS-OFFCRYPTO] §2.3.7.4 notes the first for workbooks).
//! A file that needs a real password is refused exactly as before: nothing
//! is guessed.
//!
//! Two families are read, both named by the file:
//!
//! * the encrypted package of a `.docx` / `.xlsx` / `.pptx` — a compound
//!   file holding `EncryptionInfo` and `EncryptedPackage` — with the
//!   Standard (AES, SHA-1, §2.3.4.5) and Agile (AES-CBC, SHA-1/256/384/512,
//!   §2.3.4.10) encryption, giving back the zip;
//! * the RC4 workbook stream of an `.xls` (§2.3.6 RC4, §2.3.5 RC4 CryptoAPI),
//!   decrypted in place.
//!
//! Compiled only with the `default-password` feature.
//!
//! **Not for protecting anything, and not FIPS-validated.** The password tried
//! is public, so nothing here keeps a secret: the code only reads files whose
//! "protection" is a formality. Its MD5, SHA-1, SHA-2, AES and RC4 are written
//! for this one purpose (see `core::crypto`), are not constant-time and are not
//! a validated cryptographic module; MD5 and RC4 cannot be part of a FIPS 140
//! build, so a build that must be FIPS-compliant leaves the feature off.

use super::crypto::{Aes, Hash, Rc4, md5, sha1};

/// The passwords tried, in order.
const DEFAULT_PASSWORDS: [&str; 2] = ["VelvetSweatshop", ""];

/// The most hash iterations an Agile `spinCount` may ask for. Office writes
/// 100,000; the bound keeps a crafted file from holding the reader (at most
/// two derivations, one per default password, of this many hashes: well under
/// a second).
const MAX_SPIN_COUNT: u32 = 200_000;

/// Why a file was not decrypted.
#[derive(Debug, PartialEq, Eq)]
pub(crate) enum Failure {
    /// None of the default passwords opens it: it needs a real one.
    NeedsPassword,
    /// An encryption scheme or parameter this reader does not implement.
    Unsupported(&'static str),
    /// The encryption structures are damaged.
    Malformed(&'static str),
}

type Result<T> = std::result::Result<T, Failure>;

fn utf16le(password: &str) -> Vec<u8> {
    password.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

fn le16(data: &[u8], at: usize) -> Result<u16> {
    data.get(at..at + 2)
        .map(|b| u16::from_le_bytes([b[0], b[1]]))
        .ok_or(Failure::Malformed("the encryption header is cut short"))
}

fn le32(data: &[u8], at: usize) -> Result<u32> {
    data.get(at..at + 4)
        .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
        .ok_or(Failure::Malformed("the encryption header is cut short"))
}

fn take(data: &[u8], at: usize, len: usize) -> Result<&[u8]> {
    at.checked_add(len)
        .and_then(|end| data.get(at..end))
        .ok_or(Failure::Malformed("the encryption header is cut short"))
}

/// Pad `bytes` with `fill`, or cut it, to `len`.
fn resize(mut bytes: Vec<u8>, len: usize, fill: u8) -> Vec<u8> {
    bytes.resize(len, fill);
    bytes
}

fn iv16(bytes: &[u8]) -> [u8; 16] {
    let mut iv = [0x36u8; 16];
    let n = bytes.len().min(16);
    iv[..n].copy_from_slice(&bytes[..n]);
    iv
}

// ---------------------------------------------------------------------------
// The encrypted package of an OOXML file ([MS-OFFCRYPTO] §2.3.4)
// ---------------------------------------------------------------------------

/// Decrypt the package of an encrypted OOXML compound file from its
/// `EncryptionInfo` and `EncryptedPackage` streams, giving the zip.
pub(crate) fn decrypt_package(info: &[u8], package: &[u8]) -> Result<Vec<u8>> {
    let (major, minor) = (le16(info, 0)?, le16(info, 2)?);
    let size = package
        .get(..8)
        .map(|b| u64::from_le_bytes([b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7]]))
        .ok_or(Failure::Malformed("the encrypted package is cut short"))?;
    let body = &package[8..];
    let mut plain = match (major, minor) {
        (4, 4) => agile(info, body)?,
        (2..=4, 2) => standard(info, body)?,
        _ => return Err(Failure::Unsupported("this encryption version")),
    };
    // The stream size is the package's; the rest is padding. A size larger
    // than what decrypted is a damaged header, not a short package.
    let size = usize::try_from(size)
        .ok()
        .filter(|&size| size <= plain.len())
        .ok_or(Failure::Malformed("the encrypted package is shorter than its declared size"))?;
    plain.truncate(size);
    Ok(plain)
}

/// The value of attribute `name` of `e`.
fn attr(e: &quick_xml::events::BytesStart<'_>, name: &str) -> Option<String> {
    e.attributes()
        .with_checks(false)
        .flatten()
        .find(|a| a.key.as_ref() == name)
        .map(|a| a.value.into_owned())
}

fn attr_bytes(e: &quick_xml::events::BytesStart<'_>, name: &str) -> Result<Vec<u8>> {
    let text = attr(e, name).ok_or(Failure::Malformed("an encryption attribute is missing"))?;
    super::base64::decode(&text)
        .map_err(|_| Failure::Malformed("an encryption attribute is not base64"))
}

fn attr_number(e: &quick_xml::events::BytesStart<'_>, name: &str) -> Result<u32> {
    attr(e, name)
        .and_then(|v| v.parse().ok())
        .ok_or(Failure::Malformed("an encryption attribute is not a number"))
}

fn hash_named(e: &quick_xml::events::BytesStart<'_>) -> Result<Hash> {
    match attr(e, "hashAlgorithm").as_deref() {
        Some("SHA1") => Ok(Hash::Sha1),
        Some("SHA256") => Ok(Hash::Sha256),
        Some("SHA384") => Ok(Hash::Sha384),
        Some("SHA512") => Ok(Hash::Sha512),
        _ => Err(Failure::Unsupported("this hash algorithm")),
    }
}

fn require_aes_cbc(e: &quick_xml::events::BytesStart<'_>) -> Result<()> {
    if attr(e, "cipherAlgorithm").as_deref() == Some("AES")
        && attr(e, "cipherChaining").as_deref() == Some("ChainingModeCBC")
    {
        Ok(())
    } else {
        Err(Failure::Unsupported("this cipher"))
    }
}

/// The parameters of an Agile encryption descriptor.
struct Agile {
    data_salt: Vec<u8>,
    data_hash: Hash,
    key_salt: Vec<u8>,
    key_hash: Hash,
    key_bits: usize,
    spin: u32,
    verifier_input: Vec<u8>,
    verifier_value: Vec<u8>,
    key_value: Vec<u8>,
}

fn parse_agile(info: &[u8]) -> Result<Agile> {
    let xml = info
        .get(8..)
        .ok_or(Failure::Malformed("the encryption header is cut short"))?;
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut data: Option<(Vec<u8>, Hash)> = None;
    let mut key: Option<Agile> = None;
    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(e) | quick_xml::events::Event::Empty(e)) => {
                match e.local_name().as_ref() {
                    "keyData" if data.is_none() => {
                        require_aes_cbc(&e)?;
                        data = Some((attr_bytes(&e, "saltValue")?, hash_named(&e)?));
                    },
                    "encryptedKey" if key.is_none() => {
                        require_aes_cbc(&e)?;
                        key = Some(Agile {
                            data_salt: Vec::new(),
                            data_hash: Hash::Sha1,
                            key_salt: attr_bytes(&e, "saltValue")?,
                            key_hash: hash_named(&e)?,
                            key_bits: attr_number(&e, "keyBits")? as usize,
                            spin: attr_number(&e, "spinCount")?,
                            verifier_input: attr_bytes(&e, "encryptedVerifierHashInput")?,
                            verifier_value: attr_bytes(&e, "encryptedVerifierHashValue")?,
                            key_value: attr_bytes(&e, "encryptedKeyValue")?,
                        });
                    },
                    _ => {},
                }
            },
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => {
                return Err(Failure::Malformed("the encryption descriptor is not well-formed XML"));
            },
            Ok(_) => {},
        }
    }
    let (data_salt, data_hash) = data.ok_or(Failure::Malformed("the key data is missing"))?;
    let mut key = key.ok_or(Failure::Malformed("the password key is missing"))?;
    key.data_salt = data_salt;
    key.data_hash = data_hash;
    if key.spin > MAX_SPIN_COUNT {
        return Err(Failure::Unsupported("this password spin count"));
    }
    if !matches!(key.key_bits, 128 | 192 | 256) {
        return Err(Failure::Unsupported("this key length"));
    }
    Ok(key)
}

fn agile(info: &[u8], body: &[u8]) -> Result<Vec<u8>> {
    let a = parse_agile(info)?;
    let bytes = a.key_bits / 8;
    // [MS-OFFCRYPTO] §2.3.4.11: the verifier, and the intermediate key, are
    // encrypted under keys derived from the password and a block key.
    let derive = |h: &[u8], block: &[u8]| -> Vec<u8> {
        let mut input = h.to_vec();
        input.extend_from_slice(block);
        resize(a.key_hash.digest(&input), bytes, 0x36)
    };
    let cbc = |key: &[u8], data: &[u8]| -> Result<Vec<u8>> {
        let aes = Aes::new(key).ok_or(Failure::Unsupported("this key length"))?;
        Ok(aes.decrypt_cbc(&iv16(&a.key_salt), data))
    };
    let secret = DEFAULT_PASSWORDS.iter().find_map(|password| {
        let mut input = a.key_salt.clone();
        input.extend(utf16le(password));
        let (mut h, len) = a.key_hash.digest_array(&input);
        for i in 0..a.spin {
            let mut next = [0u8; 4 + 64];
            next[..4].copy_from_slice(&i.to_le_bytes());
            next[4..4 + len].copy_from_slice(&h[..len]);
            h = a.key_hash.digest_array(&next[..4 + len]).0;
        }
        let h = &h[..len];
        let hash_input =
            cbc(&derive(h, &[0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]), &a.verifier_input)
                .ok()?;
        let expected =
            cbc(&derive(h, &[0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]), &a.verifier_value)
                .ok()?;
        let actual = a.key_hash.digest(&hash_input);
        if expected.get(..actual.len()) != Some(&actual[..]) {
            return None;
        }
        let mut key =
            cbc(&derive(h, &[0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]), &a.key_value)
                .ok()?;
        key.truncate(bytes);
        Some(key)
    });
    let secret = secret.ok_or(Failure::NeedsPassword)?;
    let aes = Aes::new(&secret).ok_or(Failure::Unsupported("this key length"))?;
    // §2.3.4.15: 4096-byte segments, each with an IV from the data salt and
    // the segment number.
    let mut out = Vec::with_capacity(body.len());
    for (i, segment) in body.chunks(4096).enumerate() {
        let mut input = a.data_salt.clone();
        input.extend_from_slice(&(i as u32).to_le_bytes());
        let iv = iv16(&a.data_hash.digest(&input));
        out.extend_from_slice(&aes.decrypt_cbc(&iv, segment));
    }
    Ok(out)
}

fn standard(info: &[u8], body: &[u8]) -> Result<Vec<u8>> {
    // EncryptionHeader (§2.3.2): flags, sizeExtra, algID, algIDHash, keySize.
    let header_size = le32(info, 8)? as usize;
    let header = take(info, 12, header_size)?;
    let (flags, alg_id, hash_id, key_bits) =
        (le32(header, 0)?, le32(header, 8)?, le32(header, 12)?, le32(header, 16)?);
    // fCryptoAPI (0x04) and fAES (0x20), AES-128/192/256 with SHA-1.
    if flags & 0x24 != 0x24 || hash_id != 0x8004 {
        return Err(Failure::Unsupported("this encryption flavour"));
    }
    let expected_bits = match alg_id {
        0x660E => 128,
        0x660F => 192,
        0x6610 => 256,
        _ => return Err(Failure::Unsupported("this cipher")),
    };
    if key_bits != expected_bits {
        return Err(Failure::Malformed("the key length does not fit the cipher"));
    }
    // EncryptionVerifier (§2.3.3) follows the header.
    let verifier = info
        .get(12 + header_size..)
        .ok_or(Failure::Malformed("the encryption header is cut short"))?;
    if le32(verifier, 0)? != 16 {
        return Err(Failure::Malformed("the verifier salt is not 16 bytes"));
    }
    let salt = take(verifier, 4, 16)?;
    let enc_verifier = take(verifier, 20, 16)?;
    let enc_hash = take(verifier, 40, 32)?;
    let key = DEFAULT_PASSWORDS.iter().find_map(|password| {
        // §2.3.4.7: iterate SHA-1 50,000 times, then derive the key.
        let mut input = salt.to_vec();
        input.extend(utf16le(password));
        let mut h = sha1(&input);
        for i in 0..50_000u32 {
            let mut next = [0u8; 24];
            next[..4].copy_from_slice(&i.to_le_bytes());
            next[4..].copy_from_slice(&h);
            h = sha1(&next);
        }
        let mut last = [0u8; 24];
        last[..20].copy_from_slice(&h);
        let h_final = sha1(&last);
        let mut x = [[0u8; 64]; 2];
        for (buf, fill) in x.iter_mut().zip([0x36u8, 0x5c]) {
            buf.fill(fill);
            for (b, h) in buf.iter_mut().zip(h_final) {
                *b ^= h;
            }
        }
        let mut key = sha1(&x[0]).to_vec();
        key.extend_from_slice(&sha1(&x[1]));
        key.truncate((key_bits / 8) as usize);
        let aes = Aes::new(&key)?;
        let verifier = aes.decrypt_ecb(enc_verifier);
        let hash = aes.decrypt_ecb(enc_hash);
        (hash.get(..20) == Some(&sha1(&verifier)[..])).then_some(key)
    });
    let key = key.ok_or(Failure::NeedsPassword)?;
    let aes = Aes::new(&key).ok_or(Failure::Unsupported("this key length"))?;
    Ok(aes.decrypt_ecb(body))
}

// ---------------------------------------------------------------------------
// The RC4 workbook stream of an .xls ([MS-XLS] §2.2.10, [MS-OFFCRYPTO] §2.3.5-6)
// ---------------------------------------------------------------------------

const RT_BOF: u16 = 0x0809;
const RT_FILEPASS: u16 = 0x002F;
const RT_BOUNDSHEET: u16 = 0x0085;
/// Records that are never encrypted ([MS-XLS] §2.2.10): BOF, FilePass,
/// UsrExcl, FileLock, InterfaceHdr, RRDInfo and RRDHead.
const NEVER_ENCRYPTED: [u16; 7] = [RT_BOF, RT_FILEPASS, 0x0194, 0x0195, 0x00E1, 0x0196, 0x0138];
/// The record stream is encrypted in blocks of this many bytes, each with its own key.
const BLOCK: usize = 1024;

/// The per-block RC4 keys of one workbook.
enum BlockKeys {
    /// RC4 (§2.3.6.2): MD5 over the password and salt.
    Md5 { intermediate: [u8; 5] },
    /// RC4 CryptoAPI (§2.3.5.2): SHA-1 over salt and password.
    CryptoApi { h0: [u8; 20], key_bits: u32 },
}

impl BlockKeys {
    fn new_md5(password: &str, salt: &[u8]) -> Self {
        let h0 = md5(&utf16le(password));
        let mut input = Vec::with_capacity(16 * 21);
        for _ in 0..16 {
            input.extend_from_slice(&h0[..5]);
            input.extend_from_slice(salt);
        }
        let h1 = md5(&input);
        Self::Md5 {
            intermediate: [h1[0], h1[1], h1[2], h1[3], h1[4]],
        }
    }

    fn new_crypto_api(password: &str, salt: &[u8], key_bits: u32) -> Self {
        let mut input = salt.to_vec();
        input.extend(utf16le(password));
        Self::CryptoApi {
            h0: sha1(&input),
            key_bits,
        }
    }

    fn cipher(&self, block: u32) -> Option<Rc4> {
        match self {
            Self::Md5 { intermediate } => {
                let mut input = intermediate.to_vec();
                input.extend_from_slice(&block.to_le_bytes());
                Rc4::new(&md5(&input))
            },
            Self::CryptoApi { h0, key_bits } => {
                let mut input = h0.to_vec();
                input.extend_from_slice(&block.to_le_bytes());
                let h = sha1(&input);
                // A 40-bit key is padded with zeros to 128 bits.
                let key = if *key_bits == 40 {
                    resize(h[..5].to_vec(), 16, 0)
                } else {
                    h[..(*key_bits / 8) as usize].to_vec()
                };
                Rc4::new(&key)
            },
        }
    }
}

/// How a `FILEPASS` record's password is checked, and the keys it gives.
struct FilePass<'a> {
    salt: &'a [u8],
    enc_verifier: &'a [u8],
    enc_hash: &'a [u8],
    /// `None` for the MD5 flavour; the key length in bits for CryptoAPI.
    crypto_api: Option<u32>,
}

fn parse_filepass(data: &[u8]) -> Result<FilePass<'_>> {
    if le16(data, 0)? != 1 {
        return Err(Failure::Unsupported("XOR obfuscation"));
    }
    let (major, minor) = (le16(data, 2)?, le16(data, 4)?);
    if (major, minor) == (1, 1) {
        return Ok(FilePass {
            salt: take(data, 6, 16)?,
            enc_verifier: take(data, 22, 16)?,
            enc_hash: take(data, 38, 16)?,
            crypto_api: None,
        });
    }
    if !(2..=4).contains(&major) || minor != 2 {
        return Err(Failure::Unsupported("this RC4 version"));
    }
    // CryptoAPI: the version, flags (4), header size (4), the header, then the
    // verifier.
    let header_size = le32(data, 10)? as usize;
    let header = take(data, 14, header_size)?;
    if le32(header, 8)? != 0x6801 || le32(header, 12)? != 0x8004 {
        return Err(Failure::Unsupported("this CryptoAPI provider"));
    }
    let key_bits = match le32(header, 16)? {
        0 => 40,
        bits @ (40 | 56 | 64 | 80 | 96 | 112 | 128) => bits,
        _ => return Err(Failure::Malformed("the RC4 key length is not valid")),
    };
    let verifier = data
        .get(14 + header_size..)
        .ok_or(Failure::Malformed("the encryption header is cut short"))?;
    if le32(verifier, 0)? != 16 {
        return Err(Failure::Malformed("the verifier salt is not 16 bytes"));
    }
    Ok(FilePass {
        salt: take(verifier, 4, 16)?,
        enc_verifier: take(verifier, 20, 16)?,
        enc_hash: take(verifier, 40, 20)?,
        crypto_api: Some(key_bits),
    })
}

/// Decrypt the record stream of an `.xls` workbook in place, using a default
/// password. `Ok(false)` when the stream holds no `FILEPASS` record (it is
/// not encrypted); on `Err` the stream is unchanged.
pub(crate) fn decrypt_workbook(stream: &mut [u8]) -> Result<bool> {
    // The FILEPASS record is among the first, after the BOF.
    let mut at = 0usize;
    let mut filepass: Option<(usize, usize)> = None;
    while at + 4 <= stream.len() && at < 1 << 16 {
        let rt = u16::from_le_bytes([stream[at], stream[at + 1]]);
        let len = usize::from(u16::from_le_bytes([stream[at + 2], stream[at + 3]]));
        if rt == RT_FILEPASS {
            filepass = Some((at, len));
            break;
        }
        at += 4 + len;
    }
    let Some((fp_at, fp_len)) = filepass else {
        return Ok(false);
    };
    let data = stream
        .get(fp_at + 4..fp_at + 4 + fp_len)
        .ok_or(Failure::Malformed("the encryption record is cut short"))?;
    let fp = parse_filepass(data)?;
    let keys = DEFAULT_PASSWORDS
        .iter()
        .find_map(|password| {
            let keys = match fp.crypto_api {
                None => BlockKeys::new_md5(password, fp.salt),
                Some(bits) => BlockKeys::new_crypto_api(password, fp.salt, bits),
            };
            // The verifier and its hash are the first bytes of block 0.
            let mut rc4 = keys.cipher(0)?;
            let mut verifier = fp.enc_verifier.to_vec();
            rc4.apply(&mut verifier);
            let mut hash = fp.enc_hash.to_vec();
            rc4.apply(&mut hash);
            let expected: Vec<u8> = match fp.crypto_api {
                None => md5(&verifier).to_vec(),
                Some(_) => sha1(&verifier).to_vec(),
            };
            (expected == hash).then_some(keys)
        })
        .ok_or(Failure::NeedsPassword)?;
    // Every byte of the stream has a place in a keystream restarted at each
    // 1024-byte block; the headers and the records never encrypted count but
    // are left as they are.
    let mut keystream_block = u32::MAX;
    let mut keystream = [0u8; BLOCK];
    let mut at = fp_at + 4 + fp_len;
    while at + 4 <= stream.len() {
        let rt = u16::from_le_bytes([stream[at], stream[at + 1]]);
        let len = usize::from(u16::from_le_bytes([stream[at + 2], stream[at + 3]]));
        let start = at + 4;
        let end = (start + len).min(stream.len());
        let skip = if NEVER_ENCRYPTED.contains(&rt) {
            len
        } else if rt == RT_BOUNDSHEET {
            4
        } else {
            0
        };
        for pos in (start + skip).min(end)..end {
            let block = (pos / BLOCK) as u32;
            if block != keystream_block {
                keystream = [0u8; BLOCK];
                keys.cipher(block)
                    .ok_or(Failure::Malformed("the RC4 key is empty"))?
                    .apply(&mut keystream);
                keystream_block = block;
            }
            stream[pos] ^= keystream[pos % BLOCK];
        }
        at = end;
    }
    // The record that said the stream was encrypted no longer does.
    stream[fp_at..fp_at + 2].copy_from_slice(&0x0DFFu16.to_le_bytes());
    Ok(true)
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;
    use crate::core::base64;

    const AGILE_SPIN: u32 = 1000;

    pub(crate) fn pattern(len: usize, seed: u8) -> Vec<u8> {
        (0..len)
            .map(|i| (i as u8).wrapping_mul(7).wrapping_add(seed))
            .collect()
    }

    // -- encryptors, written from the specification text, one per scheme --

    /// An Agile `EncryptionInfo` and `EncryptedPackage` for `payload` under
    /// `password`, SHA-512 and AES-256.
    pub(crate) fn encrypt_agile(payload: &[u8], password: &str) -> (Vec<u8>, Vec<u8>) {
        let (key_salt, data_salt) = (pattern(16, 1), pattern(16, 2));
        let secret = pattern(32, 3);
        let hash = Hash::Sha512;
        let mut input = key_salt.clone();
        input.extend(utf16le(password));
        let mut h = hash.digest(&input);
        for i in 0..AGILE_SPIN {
            let mut next = i.to_le_bytes().to_vec();
            next.extend_from_slice(&h);
            h = hash.digest(&next);
        }
        let key_for = |block: &[u8]| {
            let mut input = h.clone();
            input.extend_from_slice(block);
            hash.digest(&input)[..32].to_vec()
        };
        let iv = iv16(&key_salt);
        let verifier = pattern(16, 4);
        let enc_input = Aes::new(&key_for(&[0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]))
            .unwrap()
            .encrypt_cbc(&iv, &verifier);
        let enc_value = Aes::new(&key_for(&[0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]))
            .unwrap()
            .encrypt_cbc(&iv, &hash.digest(&verifier));
        let enc_key = Aes::new(&key_for(&[0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]))
            .unwrap()
            .encrypt_cbc(&iv, &secret);
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="{}"/><keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><p:encryptedKey spinCount="{AGILE_SPIN}" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="{}" encryptedVerifierHashInput="{}" encryptedVerifierHashValue="{}" encryptedKeyValue="{}"/></keyEncryptor></keyEncryptors></encryption>"#,
            base64::encode(&data_salt),
            base64::encode(&key_salt),
            base64::encode(&enc_input),
            base64::encode(&enc_value),
            base64::encode(&enc_key),
        );
        let mut info = vec![4, 0, 4, 0];
        info.extend_from_slice(&0x40u32.to_le_bytes());
        info.extend_from_slice(xml.as_bytes());
        let aes = Aes::new(&secret).unwrap();
        let mut package = (payload.len() as u64).to_le_bytes().to_vec();
        for (i, segment) in payload.chunks(4096).enumerate() {
            let mut input = data_salt.clone();
            input.extend_from_slice(&(i as u32).to_le_bytes());
            let mut padded = segment.to_vec();
            padded.resize(segment.len().next_multiple_of(16), 0);
            package.extend(aes.encrypt_cbc(&iv16(&hash.digest(&input)), &padded));
        }
        (info, package)
    }

    /// A Standard `EncryptionInfo` and `EncryptedPackage` for `payload`
    /// under `password`, AES-128 and SHA-1.
    pub(crate) fn encrypt_standard(payload: &[u8], password: &str) -> (Vec<u8>, Vec<u8>) {
        let salt = pattern(16, 5);
        let mut input = salt.clone();
        input.extend(utf16le(password));
        let mut h = sha1(&input).to_vec();
        for i in 0..50_000u32 {
            let mut next = i.to_le_bytes().to_vec();
            next.extend_from_slice(&h);
            h = sha1(&next).to_vec();
        }
        h.extend_from_slice(&0u32.to_le_bytes());
        let h_final = sha1(&h);
        let mut key = Vec::new();
        for fill in [0x36u8, 0x5c] {
            let mut buf = [fill; 64];
            for (b, h) in buf.iter_mut().zip(h_final) {
                *b ^= h;
            }
            key.extend_from_slice(&sha1(&buf));
        }
        key.truncate(16);
        let aes = Aes::new(&key).unwrap();
        let verifier = pattern(16, 6);
        let mut hash = sha1(&verifier).to_vec();
        hash.resize(32, 0);
        let mut header = Vec::new();
        for field in [0x24u32, 0, 0x660E, 0x8004, 128, 0x18, 0, 0] {
            header.extend_from_slice(&field.to_le_bytes());
        }
        let mut info = vec![4, 0, 2, 0];
        info.extend_from_slice(&0x24u32.to_le_bytes());
        info.extend_from_slice(&(header.len() as u32).to_le_bytes());
        info.extend_from_slice(&header);
        info.extend_from_slice(&16u32.to_le_bytes());
        info.extend_from_slice(&salt);
        info.extend(aes.encrypt_ecb(&verifier));
        info.extend_from_slice(&20u32.to_le_bytes());
        info.extend(aes.encrypt_ecb(&hash));
        let mut padded = payload.to_vec();
        padded.resize(payload.len().next_multiple_of(16), 0);
        let mut package = (payload.len() as u64).to_le_bytes().to_vec();
        package.extend(aes.encrypt_ecb(&padded));
        (info, package)
    }

    pub(crate) fn record(rt: u16, data: &[u8]) -> Vec<u8> {
        let mut out = rt.to_le_bytes().to_vec();
        out.extend_from_slice(&(data.len() as u16).to_le_bytes());
        out.extend_from_slice(data);
        out
    }

    /// The plain records after the `FILEPASS`: enough to cross several
    /// 1024-byte blocks, with a BOUNDSHEET and a record that is never encrypted.
    pub(crate) fn plain_records() -> Vec<u8> {
        let mut out = record(0x00E1, &pattern(2, 9)); // InterfaceHdr
        out.extend(record(RT_BOUNDSHEET, &pattern(12, 10)));
        for i in 0..60u8 {
            out.extend(record(0x0204, &pattern(40, i)));
        }
        out.extend(record(0x000A, &[]));
        out
    }

    /// A BIFF8 stream: BOF, the `FILEPASS` data given, then the records
    /// encrypted the way the specification describes (per block, headers and
    /// never-encrypted records counted but left alone).
    pub(crate) fn encrypt_workbook(
        bof: &[u8],
        filepass: &[u8],
        cipher: impl Fn(u32) -> Rc4,
        tail: &[u8],
    ) -> (Vec<u8>, Vec<u8>) {
        let mut plain = bof.to_vec();
        plain.extend(record(RT_FILEPASS, filepass));
        let records_from = plain.len();
        plain.extend_from_slice(tail);
        let mut encrypted = plain.clone();
        let mut at = records_from;
        while at + 4 <= plain.len() {
            let rt = u16::from_le_bytes([plain[at], plain[at + 1]]);
            let len = usize::from(u16::from_le_bytes([plain[at + 2], plain[at + 3]]));
            let skip = if NEVER_ENCRYPTED.contains(&rt) {
                len
            } else if rt == RT_BOUNDSHEET {
                4
            } else {
                0
            };
            for pos in at + 4 + skip..at + 4 + len {
                let mut rc4 = cipher((pos / BLOCK) as u32);
                let mut skipped = vec![0u8; pos % BLOCK + 1];
                rc4.apply(&mut skipped);
                encrypted[pos] ^= skipped[pos % BLOCK];
            }
            at += 4 + len;
        }
        (plain, encrypted)
    }

    /// `plain`, a complete workbook stream, encrypted with the default RC4
    /// scheme under `password`: a `FILEPASS` record after its opening `BOF`,
    /// the `BOUNDSHEET` positions moved to match, and the records after it
    /// encrypted.
    pub(crate) fn encrypt_whole_workbook(plain: &[u8], password: &str) -> Vec<u8> {
        let bof_len = 4 + usize::from(u16::from_le_bytes([plain[2], plain[3]]));
        let mut tail = plain[bof_len..].to_vec();
        let filepass_len = 4 + 6 + 48;
        let mut at = 0;
        while at + 4 <= tail.len() {
            let rt = u16::from_le_bytes([tail[at], tail[at + 1]]);
            let len = usize::from(u16::from_le_bytes([tail[at + 2], tail[at + 3]]));
            if rt == RT_BOUNDSHEET {
                let pos =
                    u32::from_le_bytes([tail[at + 4], tail[at + 5], tail[at + 6], tail[at + 7]]);
                tail[at + 4..at + 8].copy_from_slice(&(pos + filepass_len as u32).to_le_bytes());
            }
            at += 4 + len;
        }
        let salt = pattern(16, 12);
        let keys = BlockKeys::new_md5(password, &salt);
        let verifier = pattern(16, 13);
        let mut hash = md5(&verifier).to_vec();
        let mut v = verifier.clone();
        let mut rc4 = keys.cipher(0).unwrap();
        rc4.apply(&mut v);
        rc4.apply(&mut hash);
        let mut data = vec![1, 0, 1, 0, 1, 0];
        data.extend_from_slice(&salt);
        data.extend(v);
        data.extend(hash);
        assert_eq!(4 + data.len(), filepass_len);
        encrypt_workbook(&plain[..bof_len], &data, |b| keys.cipher(b).unwrap(), &tail).1
    }

    pub(crate) fn rc4_md5_stream(password: &str, tail: &[u8]) -> (Vec<u8>, Vec<u8>) {
        let salt = pattern(16, 12);
        let keys = BlockKeys::new_md5(password, &salt);
        let verifier = pattern(16, 13);
        let mut hash = md5(&verifier).to_vec();
        let mut v = verifier.clone();
        let mut rc4 = keys.cipher(0).unwrap();
        rc4.apply(&mut v);
        rc4.apply(&mut hash);
        let mut data = vec![1, 0, 1, 0, 1, 0];
        data.extend_from_slice(&salt);
        data.extend(v);
        data.extend(hash);
        encrypt_workbook(
            &record(RT_BOF, &pattern(16, 11)),
            &data,
            |b| keys.cipher(b).unwrap(),
            tail,
        )
    }

    pub(crate) fn rc4_crypto_api_stream(
        password: &str,
        key_bits: u32,
        tail: &[u8],
    ) -> (Vec<u8>, Vec<u8>) {
        let salt = pattern(16, 14);
        let keys = BlockKeys::new_crypto_api(password, &salt, key_bits);
        let verifier = pattern(16, 15);
        let mut hash = sha1(&verifier).to_vec();
        let mut v = verifier.clone();
        let mut rc4 = keys.cipher(0).unwrap();
        rc4.apply(&mut v);
        rc4.apply(&mut hash);
        let mut header = Vec::new();
        for field in [0x04u32, 0, 0x6801, 0x8004, key_bits, 1, 0, 0] {
            header.extend_from_slice(&field.to_le_bytes());
        }
        let mut data = vec![1, 0, 2, 0, 2, 0];
        data.extend_from_slice(&0x04u32.to_le_bytes());
        data.extend_from_slice(&(header.len() as u32).to_le_bytes());
        data.extend_from_slice(&header);
        data.extend_from_slice(&16u32.to_le_bytes());
        data.extend_from_slice(&salt);
        data.extend(v);
        data.extend_from_slice(&20u32.to_le_bytes());
        data.extend(hash);
        encrypt_workbook(
            &record(RT_BOF, &pattern(16, 11)),
            &data,
            |b| keys.cipher(b).unwrap(),
            tail,
        )
    }

    /// A compound file holding `streams`, each in regular sectors.
    pub(crate) fn cfb_with_streams(streams: &[(&str, &[u8])]) -> Vec<u8> {
        const END: u32 = 0xFFFF_FFFE;
        const NONE: u32 = 0xFFFF_FFFF;
        let sectors = |len: usize| len.div_ceil(512);
        let mut fat = vec![NONE; 128];
        fat[0] = 0xFFFF_FFFD; // the FAT sector itself
        fat[1] = END; // the directory (up to four entries)
        let mut next = 2usize;
        let mut starts = Vec::new();
        for (_, data) in streams {
            let n = sectors(data.len());
            starts.push(next as u32);
            for k in 0..n {
                fat[next + k] = if k + 1 == n {
                    END
                } else {
                    (next + k + 1) as u32
                };
            }
            next += n;
        }
        assert!(next <= 128 && streams.len() <= 3);
        let mut file = vec![0u8; 512 * (1 + next)];
        file[..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
        file[0x18..0x1A].copy_from_slice(&0x003Eu16.to_le_bytes());
        file[0x1A..0x1C].copy_from_slice(&3u16.to_le_bytes());
        file[0x1C..0x1E].copy_from_slice(&0xFFFEu16.to_le_bytes());
        file[0x1E..0x20].copy_from_slice(&9u16.to_le_bytes());
        file[0x20..0x22].copy_from_slice(&6u16.to_le_bytes());
        file[0x2C..0x30].copy_from_slice(&1u32.to_le_bytes()); // FAT sectors
        file[0x30..0x34].copy_from_slice(&1u32.to_le_bytes()); // directory
        file[0x38..0x3C].copy_from_slice(&1u32.to_le_bytes()); // no mini stream
        file[0x3C..0x40].copy_from_slice(&END.to_le_bytes());
        file[0x44..0x48].copy_from_slice(&END.to_le_bytes());
        file[0x4C..0x50].copy_from_slice(&0u32.to_le_bytes()); // DIFAT[0]: sector 0
        for i in 1..109 {
            file[0x4C + 4 * i..0x50 + 4 * i].copy_from_slice(&NONE.to_le_bytes());
        }
        for (i, entry) in fat.iter().enumerate() {
            file[512 + 4 * i..516 + 4 * i].copy_from_slice(&entry.to_le_bytes());
        }
        let mut entries: Vec<(&str, u8, u32, u32, u32, u64)> =
            vec![("Root Entry", 5, NONE, NONE, 1, 0)];
        for (i, ((name, data), start)) in streams.iter().zip(&starts).enumerate() {
            let right = if i + 1 < streams.len() {
                (i + 2) as u32
            } else {
                NONE
            };
            entries.push((name, 2, NONE, right, NONE, 0));
            let last = entries.last_mut().unwrap();
            last.4 = *start;
            last.5 = data.len() as u64;
        }
        for (i, (name, kind, left, right, extra, size)) in entries.into_iter().enumerate() {
            let at = 1024 + 128 * i;
            for (k, unit) in name.encode_utf16().enumerate() {
                file[at + 2 * k..at + 2 * k + 2].copy_from_slice(&unit.to_le_bytes());
            }
            file[at + 64..at + 66].copy_from_slice(&((name.len() as u16 + 1) * 2).to_le_bytes());
            file[at + 66] = kind;
            file[at + 67] = 1;
            file[at + 68..at + 72].copy_from_slice(&left.to_le_bytes());
            file[at + 72..at + 76].copy_from_slice(&right.to_le_bytes());
            // The root's `extra` is its child; a stream's is its first sector.
            let (child, start) = if kind == 5 {
                (extra, END)
            } else {
                (NONE, extra)
            };
            file[at + 76..at + 80].copy_from_slice(&child.to_le_bytes());
            file[at + 116..at + 120].copy_from_slice(&start.to_le_bytes());
            file[at + 120..at + 128].copy_from_slice(&size.to_le_bytes());
        }
        for ((_, data), start) in streams.iter().zip(&starts) {
            let at = 512 * (1 + *start as usize);
            file[at..at + data.len()].copy_from_slice(data);
        }
        file
    }

    // -- the OOXML package --

    #[test]
    fn test_an_agile_package_under_a_default_password_decrypts() {
        let payload = pattern(10_000, 20);
        for password in DEFAULT_PASSWORDS {
            let (info, package) = encrypt_agile(&payload, password);
            assert_eq!(decrypt_package(&info, &package), Ok(payload.clone()), "a default password");
        }
    }

    #[test]
    fn test_a_standard_package_under_a_default_password_decrypts() {
        let payload = pattern(5000, 21);
        for password in DEFAULT_PASSWORDS {
            let (info, package) = encrypt_standard(&payload, password);
            assert_eq!(decrypt_package(&info, &package), Ok(payload.clone()), "a default password");
        }
    }

    #[test]
    fn test_a_package_under_any_other_password_needs_one() {
        let (info, package) = encrypt_agile(b"payload", "hunter2");
        assert_eq!(decrypt_package(&info, &package), Err(Failure::NeedsPassword));
        let (info, package) = encrypt_standard(b"payload", "hunter2");
        assert_eq!(decrypt_package(&info, &package), Err(Failure::NeedsPassword));
    }

    #[test]
    fn test_a_damaged_or_hostile_package_is_an_error_never_a_panic() {
        let (info, package) = encrypt_agile(&pattern(300, 22), "VelvetSweatshop");
        for cut in 0..info.len().min(400) {
            let _ = decrypt_package(&info[..cut], &package);
        }
        for cut in 0..package.len().min(40) {
            let _ = decrypt_package(&info, &package[..cut]);
        }
        let (info, package) = encrypt_standard(&pattern(300, 23), "VelvetSweatshop");
        for cut in 0..info.len() {
            let _ = decrypt_package(&info[..cut], &package);
        }
        // An absurd spin count is refused before it is spent.
        let (info, package) = encrypt_agile(b"x", "VelvetSweatshop");
        let text = String::from_utf8_lossy(&info[8..])
            .replace(&format!("spinCount=\"{AGILE_SPIN}\""), "spinCount=\"4000000000\"");
        let mut hostile = info[..8].to_vec();
        hostile.extend_from_slice(text.as_bytes());
        assert_eq!(
            decrypt_package(&hostile, &package),
            Err(Failure::Unsupported("this password spin count"))
        );
        // A stream size larger than the data is a damaged header, not a short package.
        let (info, mut package) = encrypt_agile(&pattern(100, 24), "VelvetSweatshop");
        package[..8].copy_from_slice(&u64::MAX.to_le_bytes());
        assert_eq!(
            decrypt_package(&info, &package),
            Err(Failure::Malformed("the encrypted package is shorter than its declared size"))
        );
        // A descriptor that is not well-formed XML says so, not "key data missing".
        let (mut info, package) = encrypt_agile(b"x", "VelvetSweatshop");
        info.truncate(8);
        info.extend_from_slice(b"<encryption><keyData <<");
        assert_eq!(
            decrypt_package(&info, &package),
            Err(Failure::Malformed("the encryption descriptor is not well-formed XML"))
        );
    }

    // -- the RC4 workbook --

    #[test]
    fn test_an_rc4_workbook_under_a_default_password_decrypts_in_place() {
        for password in DEFAULT_PASSWORDS {
            let (plain, mut stream) = rc4_md5_stream(password, &plain_records());
            assert_ne!(plain, stream);
            assert_eq!(decrypt_workbook(&mut stream), Ok(true), "a default password");
            // Everything but the record that said "encrypted" is the original.
            let fp = 4 + 16;
            assert_eq!(&stream[..fp], &plain[..fp]);
            let after = fp + 4 + usize::from(u16::from_le_bytes([plain[fp + 2], plain[fp + 3]]));
            assert_eq!(&stream[after..], &plain[after..]);
        }
    }

    #[test]
    fn test_an_rc4_crypto_api_workbook_decrypts_for_each_key_length() {
        for key_bits in [40, 128] {
            let (plain, mut stream) =
                rc4_crypto_api_stream("VelvetSweatshop", key_bits, &plain_records());
            assert_eq!(decrypt_workbook(&mut stream), Ok(true), "{key_bits}");
            let after = 4 + 16 + 4 + usize::from(u16::from_le_bytes([plain[22], plain[23]]));
            assert_eq!(&stream[after..], &plain[after..], "{key_bits}");
        }
    }

    #[test]
    fn test_a_workbook_under_another_password_is_left_untouched() {
        for (_, mut stream) in [
            rc4_md5_stream("hunter2", &plain_records()),
            rc4_crypto_api_stream("hunter2", 128, &plain_records()),
        ] {
            let before = stream.clone();
            assert_eq!(decrypt_workbook(&mut stream), Err(Failure::NeedsPassword));
            assert_eq!(stream, before);
        }
    }

    #[test]
    fn test_a_workbook_without_encryption_is_reported_as_plain() {
        let mut stream = record(RT_BOF, &pattern(16, 1));
        stream.extend(record(0x000A, &[]));
        let before = stream.clone();
        assert_eq!(decrypt_workbook(&mut stream), Ok(false));
        assert_eq!(stream, before);
    }

    #[test]
    fn test_xor_obfuscation_and_cut_headers_are_refused_not_guessed() {
        let mut xor = record(RT_BOF, &pattern(16, 1));
        xor.extend(record(RT_FILEPASS, &[0, 0, 0x12, 0x34, 0x56, 0x78]));
        assert_eq!(decrypt_workbook(&mut xor), Err(Failure::Unsupported("XOR obfuscation")));
        let (_, stream) = rc4_crypto_api_stream("VelvetSweatshop", 128, &plain_records());
        for cut in 0..80 {
            let mut part = stream[..cut.min(stream.len())].to_vec();
            let _ = decrypt_workbook(&mut part);
        }
    }

    // -- through the public readers --

    fn docx_bytes(marker: &str) -> Vec<u8> {
        let mut w = crate::docx::write::DocxWriter::new();
        w.add_paragraph(marker);
        let mut out = std::io::Cursor::new(Vec::new());
        w.write_to(&mut out).unwrap();
        out.into_inner()
    }

    fn encrypted_file(package: (Vec<u8>, Vec<u8>)) -> Vec<u8> {
        // Both streams live in regular sectors, however short they are.
        cfb_with_streams(&[
            ("EncryptionInfo", &package.0),
            ("EncryptedPackage", &package.1),
        ])
    }

    #[test]
    fn test_an_encrypted_docx_opens_with_the_default_password() {
        for package in [
            encrypt_agile(&docx_bytes("agile marker"), "VelvetSweatshop"),
            encrypt_standard(&docx_bytes("standard marker"), ""),
        ] {
            let file = encrypted_file(package);
            let doc = crate::Document::from_reader(
                std::io::Cursor::new(file.clone()),
                crate::DocumentFormat::Docx,
            )
            .unwrap();
            assert!(doc.plain_text().contains("marker"), "{}", doc.plain_text());
            // And by path, where the extension names the format.
            let path = std::env::temp_dir()
                .join(format!("office_oxide_default_pw_{}.docx", std::process::id()));
            std::fs::write(&path, &file).unwrap();
            let opened = crate::Document::open(&path);
            std::fs::remove_file(&path).ok();
            assert!(opened.unwrap().plain_text().contains("marker"));
        }
    }

    #[test]
    fn test_an_encrypted_docx_that_needs_a_real_password_is_still_refused() {
        let file = encrypted_file(encrypt_agile(&docx_bytes("secret"), "hunter2"));
        let error =
            crate::Document::from_reader(std::io::Cursor::new(file), crate::DocumentFormat::Docx)
                .err()
                .expect("refused");
        assert!(error.to_string().contains("password-protected"), "{error}");
    }

    #[test]
    fn test_a_decrypted_payload_that_is_another_container_is_not_decrypted_again() {
        let inner = encrypted_file(encrypt_agile(&docx_bytes("deep"), "VelvetSweatshop"));
        let outer = encrypted_file(encrypt_agile(&inner, "VelvetSweatshop"));
        let error =
            crate::Document::from_reader(std::io::Cursor::new(outer), crate::DocumentFormat::Docx)
                .err()
                .expect("refused");
        assert!(error.to_string().contains("not a zip package"), "{error}");
    }

    #[test]
    fn test_a_spin_count_past_the_bound_is_refused_before_it_is_spent() {
        let (info, package) = encrypt_agile(b"x", "VelvetSweatshop");
        let text = String::from_utf8_lossy(&info[8..]).replace(
            &format!("spinCount=\"{AGILE_SPIN}\""),
            &format!("spinCount=\"{}\"", MAX_SPIN_COUNT + 1),
        );
        let mut hostile = info[..8].to_vec();
        hostile.extend_from_slice(text.as_bytes());
        assert_eq!(
            decrypt_package(&hostile, &package),
            Err(Failure::Unsupported("this password spin count"))
        );
    }

    #[test]
    fn test_a_file_that_cannot_be_opened_says_why() {
        let open = |file: Vec<u8>| {
            crate::Document::from_reader(std::io::Cursor::new(file), crate::DocumentFormat::Docx)
                .err()
                .expect("refused")
                .to_string()
        };
        // A real password is needed.
        let (info, package) = encrypt_agile(b"x", "hunter2");
        assert!(open(encrypted_file((info, package))).contains("not Office's default one"));
        // A scheme the reader does not implement (an encryption version it does not know).
        let (mut info, package) = encrypt_agile(b"x", "VelvetSweatshop");
        info[..4].copy_from_slice(&[9, 0, 9, 0]);
        assert!(open(encrypted_file((info, package))).contains("does not support"));
        // A damaged header.
        let (info, package) = encrypt_standard(b"x", "VelvetSweatshop");
        assert!(open(encrypted_file((info[..20].to_vec(), package))).contains("damaged"));
    }
}
