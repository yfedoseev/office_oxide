//! The cryptographic primitives that opening a document protected with
//! Office's fixed default password needs: MD5, SHA-1, SHA-2, AES decryption
//! (ECB and CBC) and RC4.
//!
//! Compiled only with the `default-password` feature, which is off by
//! default: a build without it contains none of this code, and no dependency
//! is added either way. MD5 and RC4 are not approved for use in a FIPS
//! 140 validated module; they are here only because the legacy Office file
//! formats are defined in terms of them ([MS-OFFCRYPTO] §2.3), and only to
//! *decrypt* with the well-known default password, never to protect anything.
//!
//! Each primitive is checked against its published test vectors.

const MD5_K: [u32; 64] = [
    0xd76aa478, 0xe8c7b756, 0x242070db, 0xc1bdceee, 0xf57c0faf, 0x4787c62a, 0xa8304613, 0xfd469501,
    0x698098d8, 0x8b44f7af, 0xffff5bb1, 0x895cd7be, 0x6b901122, 0xfd987193, 0xa679438e, 0x49b40821,
    0xf61e2562, 0xc040b340, 0x265e5a51, 0xe9b6c7aa, 0xd62f105d, 0x02441453, 0xd8a1e681, 0xe7d3fbc8,
    0x21e1cde6, 0xc33707d6, 0xf4d50d87, 0x455a14ed, 0xa9e3e905, 0xfcefa3f8, 0x676f02d9, 0x8d2a4c8a,
    0xfffa3942, 0x8771f681, 0x6d9d6122, 0xfde5380c, 0xa4beea44, 0x4bdecfa9, 0xf6bb4b60, 0xbebfbc70,
    0x289b7ec6, 0xeaa127fa, 0xd4ef3085, 0x04881d05, 0xd9d4d039, 0xe6db99e5, 0x1fa27cf8, 0xc4ac5665,
    0xf4292244, 0x432aff97, 0xab9423a7, 0xfc93a039, 0x655b59c3, 0x8f0ccc92, 0xffeff47d, 0x85845dd1,
    0x6fa87e4f, 0xfe2ce6e0, 0xa3014314, 0x4e0811a1, 0xf7537e82, 0xbd3af235, 0x2ad7d2bb, 0xeb86d391,
];
const SHA256_K: [u32; 64] = [
    0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
    0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
    0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
    0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
    0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
    0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
    0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
    0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
];
const SHA512_K: [u64; 80] = [
    0x428a2f98d728ae22,
    0x7137449123ef65cd,
    0xb5c0fbcfec4d3b2f,
    0xe9b5dba58189dbbc,
    0x3956c25bf348b538,
    0x59f111f1b605d019,
    0x923f82a4af194f9b,
    0xab1c5ed5da6d8118,
    0xd807aa98a3030242,
    0x12835b0145706fbe,
    0x243185be4ee4b28c,
    0x550c7dc3d5ffb4e2,
    0x72be5d74f27b896f,
    0x80deb1fe3b1696b1,
    0x9bdc06a725c71235,
    0xc19bf174cf692694,
    0xe49b69c19ef14ad2,
    0xefbe4786384f25e3,
    0x0fc19dc68b8cd5b5,
    0x240ca1cc77ac9c65,
    0x2de92c6f592b0275,
    0x4a7484aa6ea6e483,
    0x5cb0a9dcbd41fbd4,
    0x76f988da831153b5,
    0x983e5152ee66dfab,
    0xa831c66d2db43210,
    0xb00327c898fb213f,
    0xbf597fc7beef0ee4,
    0xc6e00bf33da88fc2,
    0xd5a79147930aa725,
    0x06ca6351e003826f,
    0x142929670a0e6e70,
    0x27b70a8546d22ffc,
    0x2e1b21385c26c926,
    0x4d2c6dfc5ac42aed,
    0x53380d139d95b3df,
    0x650a73548baf63de,
    0x766a0abb3c77b2a8,
    0x81c2c92e47edaee6,
    0x92722c851482353b,
    0xa2bfe8a14cf10364,
    0xa81a664bbc423001,
    0xc24b8b70d0f89791,
    0xc76c51a30654be30,
    0xd192e819d6ef5218,
    0xd69906245565a910,
    0xf40e35855771202a,
    0x106aa07032bbd1b8,
    0x19a4c116b8d2d0c8,
    0x1e376c085141ab53,
    0x2748774cdf8eeb99,
    0x34b0bcb5e19b48a8,
    0x391c0cb3c5c95a63,
    0x4ed8aa4ae3418acb,
    0x5b9cca4f7763e373,
    0x682e6ff3d6b2b8a3,
    0x748f82ee5defb2fc,
    0x78a5636f43172f60,
    0x84c87814a1f0ab72,
    0x8cc702081a6439ec,
    0x90befffa23631e28,
    0xa4506cebde82bde9,
    0xbef9a3f7b2c67915,
    0xc67178f2e372532b,
    0xca273eceea26619c,
    0xd186b8c721c0c207,
    0xeada7dd6cde0eb1e,
    0xf57d4f7fee6ed178,
    0x06f067aa72176fba,
    0x0a637dc5a2c898a6,
    0x113f9804bef90dae,
    0x1b710b35131c471b,
    0x28db77f523047d84,
    0x32caab7b40c72493,
    0x3c9ebe0a15c9bebc,
    0x431d67c49c100d4c,
    0x4cc5d4becb3e42b6,
    0x597f299cfc657e2a,
    0x5fcb6fab3ad6faec,
    0x6c44198c4a475817,
];
const SHA256_IV: [u32; 8] = [
    0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19,
];
const SHA512_IV: [u64; 8] = [
    0x6a09e667f3bcc908,
    0xbb67ae8584caa73b,
    0x3c6ef372fe94f82b,
    0xa54ff53a5f1d36f1,
    0x510e527fade682d1,
    0x9b05688c2b3e6c1f,
    0x1f83d9abfb41bd6b,
    0x5be0cd19137e2179,
];
const SHA384_IV: [u64; 8] = [
    0xcbbb9d5dc1059ed8,
    0x629a292a367cd507,
    0x9159015a3070dd17,
    0x152fecd8f70e5939,
    0x67332667ffc00b31,
    0x8eb44a8768581511,
    0xdb0c2e0d64f98fa7,
    0x47b5481dbefa4fa4,
];
const SBOX: [u8; 256] = [
    0x63, 0x7c, 0x77, 0x7b, 0xf2, 0x6b, 0x6f, 0xc5, 0x30, 0x01, 0x67, 0x2b, 0xfe, 0xd7, 0xab, 0x76,
    0xca, 0x82, 0xc9, 0x7d, 0xfa, 0x59, 0x47, 0xf0, 0xad, 0xd4, 0xa2, 0xaf, 0x9c, 0xa4, 0x72, 0xc0,
    0xb7, 0xfd, 0x93, 0x26, 0x36, 0x3f, 0xf7, 0xcc, 0x34, 0xa5, 0xe5, 0xf1, 0x71, 0xd8, 0x31, 0x15,
    0x04, 0xc7, 0x23, 0xc3, 0x18, 0x96, 0x05, 0x9a, 0x07, 0x12, 0x80, 0xe2, 0xeb, 0x27, 0xb2, 0x75,
    0x09, 0x83, 0x2c, 0x1a, 0x1b, 0x6e, 0x5a, 0xa0, 0x52, 0x3b, 0xd6, 0xb3, 0x29, 0xe3, 0x2f, 0x84,
    0x53, 0xd1, 0x00, 0xed, 0x20, 0xfc, 0xb1, 0x5b, 0x6a, 0xcb, 0xbe, 0x39, 0x4a, 0x4c, 0x58, 0xcf,
    0xd0, 0xef, 0xaa, 0xfb, 0x43, 0x4d, 0x33, 0x85, 0x45, 0xf9, 0x02, 0x7f, 0x50, 0x3c, 0x9f, 0xa8,
    0x51, 0xa3, 0x40, 0x8f, 0x92, 0x9d, 0x38, 0xf5, 0xbc, 0xb6, 0xda, 0x21, 0x10, 0xff, 0xf3, 0xd2,
    0xcd, 0x0c, 0x13, 0xec, 0x5f, 0x97, 0x44, 0x17, 0xc4, 0xa7, 0x7e, 0x3d, 0x64, 0x5d, 0x19, 0x73,
    0x60, 0x81, 0x4f, 0xdc, 0x22, 0x2a, 0x90, 0x88, 0x46, 0xee, 0xb8, 0x14, 0xde, 0x5e, 0x0b, 0xdb,
    0xe0, 0x32, 0x3a, 0x0a, 0x49, 0x06, 0x24, 0x5c, 0xc2, 0xd3, 0xac, 0x62, 0x91, 0x95, 0xe4, 0x79,
    0xe7, 0xc8, 0x37, 0x6d, 0x8d, 0xd5, 0x4e, 0xa9, 0x6c, 0x56, 0xf4, 0xea, 0x65, 0x7a, 0xae, 0x08,
    0xba, 0x78, 0x25, 0x2e, 0x1c, 0xa6, 0xb4, 0xc6, 0xe8, 0xdd, 0x74, 0x1f, 0x4b, 0xbd, 0x8b, 0x8a,
    0x70, 0x3e, 0xb5, 0x66, 0x48, 0x03, 0xf6, 0x0e, 0x61, 0x35, 0x57, 0xb9, 0x86, 0xc1, 0x1d, 0x9e,
    0xe1, 0xf8, 0x98, 0x11, 0x69, 0xd9, 0x8e, 0x94, 0x9b, 0x1e, 0x87, 0xe9, 0xce, 0x55, 0x28, 0xdf,
    0x8c, 0xa1, 0x89, 0x0d, 0xbf, 0xe6, 0x42, 0x68, 0x41, 0x99, 0x2d, 0x0f, 0xb0, 0x54, 0xbb, 0x16,
];
const INV_SBOX: [u8; 256] = [
    0x52, 0x09, 0x6a, 0xd5, 0x30, 0x36, 0xa5, 0x38, 0xbf, 0x40, 0xa3, 0x9e, 0x81, 0xf3, 0xd7, 0xfb,
    0x7c, 0xe3, 0x39, 0x82, 0x9b, 0x2f, 0xff, 0x87, 0x34, 0x8e, 0x43, 0x44, 0xc4, 0xde, 0xe9, 0xcb,
    0x54, 0x7b, 0x94, 0x32, 0xa6, 0xc2, 0x23, 0x3d, 0xee, 0x4c, 0x95, 0x0b, 0x42, 0xfa, 0xc3, 0x4e,
    0x08, 0x2e, 0xa1, 0x66, 0x28, 0xd9, 0x24, 0xb2, 0x76, 0x5b, 0xa2, 0x49, 0x6d, 0x8b, 0xd1, 0x25,
    0x72, 0xf8, 0xf6, 0x64, 0x86, 0x68, 0x98, 0x16, 0xd4, 0xa4, 0x5c, 0xcc, 0x5d, 0x65, 0xb6, 0x92,
    0x6c, 0x70, 0x48, 0x50, 0xfd, 0xed, 0xb9, 0xda, 0x5e, 0x15, 0x46, 0x57, 0xa7, 0x8d, 0x9d, 0x84,
    0x90, 0xd8, 0xab, 0x00, 0x8c, 0xbc, 0xd3, 0x0a, 0xf7, 0xe4, 0x58, 0x05, 0xb8, 0xb3, 0x45, 0x06,
    0xd0, 0x2c, 0x1e, 0x8f, 0xca, 0x3f, 0x0f, 0x02, 0xc1, 0xaf, 0xbd, 0x03, 0x01, 0x13, 0x8a, 0x6b,
    0x3a, 0x91, 0x11, 0x41, 0x4f, 0x67, 0xdc, 0xea, 0x97, 0xf2, 0xcf, 0xce, 0xf0, 0xb4, 0xe6, 0x73,
    0x96, 0xac, 0x74, 0x22, 0xe7, 0xad, 0x35, 0x85, 0xe2, 0xf9, 0x37, 0xe8, 0x1c, 0x75, 0xdf, 0x6e,
    0x47, 0xf1, 0x1a, 0x71, 0x1d, 0x29, 0xc5, 0x89, 0x6f, 0xb7, 0x62, 0x0e, 0xaa, 0x18, 0xbe, 0x1b,
    0xfc, 0x56, 0x3e, 0x4b, 0xc6, 0xd2, 0x79, 0x20, 0x9a, 0xdb, 0xc0, 0xfe, 0x78, 0xcd, 0x5a, 0xf4,
    0x1f, 0xdd, 0xa8, 0x33, 0x88, 0x07, 0xc7, 0x31, 0xb1, 0x12, 0x10, 0x59, 0x27, 0x80, 0xec, 0x5f,
    0x60, 0x51, 0x7f, 0xa9, 0x19, 0xb5, 0x4a, 0x0d, 0x2d, 0xe5, 0x7a, 0x9f, 0x93, 0xc9, 0x9c, 0xef,
    0xa0, 0xe0, 0x3b, 0x4d, 0xae, 0x2a, 0xf5, 0xb0, 0xc8, 0xeb, 0xbb, 0x3c, 0x83, 0x53, 0x99, 0x61,
    0x17, 0x2b, 0x04, 0x7e, 0xba, 0x77, 0xd6, 0x26, 0xe1, 0x69, 0x14, 0x63, 0x55, 0x21, 0x0c, 0x7d,
];

/// Feed `data`, padded as MD5 and the SHA family pad (a `0x80` byte, zeros,
/// and the bit length in `len_bytes` bytes), to `f` in blocks of `block`.
fn for_each_block(
    data: &[u8],
    block: usize,
    len_bytes: usize,
    big_endian: bool,
    mut f: impl FnMut(&[u8]),
) {
    let mut chunks = data.chunks_exact(block);
    for chunk in &mut chunks {
        f(chunk);
    }
    let rest = chunks.remainder();
    let mut tail = [0u8; 256];
    tail[..rest.len()].copy_from_slice(rest);
    tail[rest.len()] = 0x80;
    let total = if rest.len() + 1 + len_bytes <= block {
        block
    } else {
        2 * block
    };
    let bits = (data.len() as u128) * 8;
    if big_endian {
        tail[total - len_bytes..total].copy_from_slice(&bits.to_be_bytes()[16 - len_bytes..]);
    } else {
        tail[total - len_bytes..total].copy_from_slice(&bits.to_le_bytes()[..len_bytes]);
    }
    f(&tail[..block]);
    if total > block {
        f(&tail[block..total]);
    }
}

const MD5_S: [u32; 64] = [
    7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9,
    14, 20, 5, 9, 14, 20, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 6, 10, 15,
    21, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21,
];

/// MD5 ([RFC 1321]).
pub(crate) fn md5(data: &[u8]) -> [u8; 16] {
    let mut s = [0x6745_2301u32, 0xefcd_ab89, 0x98ba_dcfe, 0x1032_5476];
    for_each_block(data, 64, 8, false, |block| {
        let m: Vec<u32> = block
            .as_chunks::<4>()
            .0
            .iter()
            .map(|c| u32::from_le_bytes([c[0], c[1], c[2], c[3]]))
            .collect();
        let [mut a, mut b, mut c, mut d] = s;
        for i in 0..64 {
            let (f, g) = match i / 16 {
                0 => ((b & c) | (!b & d), i),
                1 => ((d & b) | (!d & c), (5 * i + 1) % 16),
                2 => (b ^ c ^ d, (3 * i + 5) % 16),
                _ => (c ^ (b | !d), (7 * i) % 16),
            };
            let f = f.wrapping_add(a).wrapping_add(MD5_K[i]).wrapping_add(m[g]);
            a = d;
            d = c;
            c = b;
            b = b.wrapping_add(f.rotate_left(MD5_S[i]));
        }
        for (word, add) in s.iter_mut().zip([a, b, c, d]) {
            *word = word.wrapping_add(add);
        }
    });
    let mut out = [0u8; 16];
    for (chunk, word) in out.as_chunks_mut::<4>().0.iter_mut().zip(s) {
        chunk.copy_from_slice(&word.to_le_bytes());
    }
    out
}

/// SHA-1 ([FIPS 180-4]).
pub(crate) fn sha1(data: &[u8]) -> [u8; 20] {
    let mut h = [
        0x6745_2301u32,
        0xefcd_ab89,
        0x98ba_dcfe,
        0x1032_5476,
        0xc3d2_e1f0,
    ];
    for_each_block(data, 64, 8, true, |block| {
        let mut w = [0u32; 80];
        for (word, c) in w.iter_mut().zip(block.as_chunks::<4>().0.iter()) {
            *word = u32::from_be_bytes([c[0], c[1], c[2], c[3]]);
        }
        for i in 16..80 {
            w[i] = (w[i - 3] ^ w[i - 8] ^ w[i - 14] ^ w[i - 16]).rotate_left(1);
        }
        let [mut a, mut b, mut c, mut d, mut e] = h;
        for (i, &wi) in w.iter().enumerate() {
            let (f, k) = match i / 20 {
                0 => ((b & c) | (!b & d), 0x5a82_7999u32),
                1 => (b ^ c ^ d, 0x6ed9_eba1),
                2 => ((b & c) | (b & d) | (c & d), 0x8f1b_bcdc),
                _ => (b ^ c ^ d, 0xca62_c1d6),
            };
            let t = a
                .rotate_left(5)
                .wrapping_add(f)
                .wrapping_add(e)
                .wrapping_add(k)
                .wrapping_add(wi);
            e = d;
            d = c;
            c = b.rotate_left(30);
            b = a;
            a = t;
        }
        for (word, add) in h.iter_mut().zip([a, b, c, d, e]) {
            *word = word.wrapping_add(add);
        }
    });
    let mut out = [0u8; 20];
    for (chunk, word) in out.as_chunks_mut::<4>().0.iter_mut().zip(h) {
        chunk.copy_from_slice(&word.to_be_bytes());
    }
    out
}

/// SHA-256 ([FIPS 180-4]).
pub(crate) fn sha256(data: &[u8]) -> [u8; 32] {
    let mut h = SHA256_IV;
    for_each_block(data, 64, 8, true, |block| {
        let mut w = [0u32; 64];
        for (word, c) in w.iter_mut().zip(block.as_chunks::<4>().0.iter()) {
            *word = u32::from_be_bytes([c[0], c[1], c[2], c[3]]);
        }
        for i in 16..64 {
            let s0 = w[i - 15].rotate_right(7) ^ w[i - 15].rotate_right(18) ^ (w[i - 15] >> 3);
            let s1 = w[i - 2].rotate_right(17) ^ w[i - 2].rotate_right(19) ^ (w[i - 2] >> 10);
            w[i] = w[i - 16]
                .wrapping_add(s0)
                .wrapping_add(w[i - 7])
                .wrapping_add(s1);
        }
        let mut v = h;
        for i in 0..64 {
            let s1 = v[4].rotate_right(6) ^ v[4].rotate_right(11) ^ v[4].rotate_right(25);
            let ch = (v[4] & v[5]) ^ (!v[4] & v[6]);
            let t1 = v[7]
                .wrapping_add(s1)
                .wrapping_add(ch)
                .wrapping_add(SHA256_K[i])
                .wrapping_add(w[i]);
            let s0 = v[0].rotate_right(2) ^ v[0].rotate_right(13) ^ v[0].rotate_right(22);
            let maj = (v[0] & v[1]) ^ (v[0] & v[2]) ^ (v[1] & v[2]);
            let t2 = s0.wrapping_add(maj);
            v = [
                t1.wrapping_add(t2),
                v[0],
                v[1],
                v[2],
                v[3].wrapping_add(t1),
                v[4],
                v[5],
                v[6],
            ];
        }
        for (word, add) in h.iter_mut().zip(v) {
            *word = word.wrapping_add(add);
        }
    });
    let mut out = [0u8; 32];
    for (chunk, word) in out.as_chunks_mut::<4>().0.iter_mut().zip(h) {
        chunk.copy_from_slice(&word.to_be_bytes());
    }
    out
}

/// SHA-512 over `iv` ([FIPS 180-4]); SHA-384 is the same with another start
/// value, cut short.
fn sha512_with(iv: [u64; 8], data: &[u8]) -> [u8; 64] {
    let mut h = iv;
    for_each_block(data, 128, 16, true, |block| {
        let mut w = [0u64; 80];
        for (word, c) in w.iter_mut().zip(block.as_chunks::<8>().0.iter()) {
            *word = u64::from_be_bytes([c[0], c[1], c[2], c[3], c[4], c[5], c[6], c[7]]);
        }
        for i in 16..80 {
            let s0 = w[i - 15].rotate_right(1) ^ w[i - 15].rotate_right(8) ^ (w[i - 15] >> 7);
            let s1 = w[i - 2].rotate_right(19) ^ w[i - 2].rotate_right(61) ^ (w[i - 2] >> 6);
            w[i] = w[i - 16]
                .wrapping_add(s0)
                .wrapping_add(w[i - 7])
                .wrapping_add(s1);
        }
        let mut v = h;
        for i in 0..80 {
            let s1 = v[4].rotate_right(14) ^ v[4].rotate_right(18) ^ v[4].rotate_right(41);
            let ch = (v[4] & v[5]) ^ (!v[4] & v[6]);
            let t1 = v[7]
                .wrapping_add(s1)
                .wrapping_add(ch)
                .wrapping_add(SHA512_K[i])
                .wrapping_add(w[i]);
            let s0 = v[0].rotate_right(28) ^ v[0].rotate_right(34) ^ v[0].rotate_right(39);
            let maj = (v[0] & v[1]) ^ (v[0] & v[2]) ^ (v[1] & v[2]);
            let t2 = s0.wrapping_add(maj);
            v = [
                t1.wrapping_add(t2),
                v[0],
                v[1],
                v[2],
                v[3].wrapping_add(t1),
                v[4],
                v[5],
                v[6],
            ];
        }
        for (word, add) in h.iter_mut().zip(v) {
            *word = word.wrapping_add(add);
        }
    });
    let mut out = [0u8; 64];
    for (chunk, word) in out.as_chunks_mut::<8>().0.iter_mut().zip(h) {
        chunk.copy_from_slice(&word.to_be_bytes());
    }
    out
}

/// SHA-512 ([FIPS 180-4]).
pub(crate) fn sha512(data: &[u8]) -> [u8; 64] {
    sha512_with(SHA512_IV, data)
}

/// SHA-384 ([FIPS 180-4]).
pub(crate) fn sha384(data: &[u8]) -> [u8; 48] {
    let mut out = [0u8; 48];
    out.copy_from_slice(&sha512_with(SHA384_IV, data)[..48]);
    out
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
    pub(crate) fn digest(self, data: &[u8]) -> Vec<u8> {
        match self {
            Self::Sha1 => sha1(data).to_vec(),
            Self::Sha256 => sha256(data).to_vec(),
            Self::Sha384 => sha384(data).to_vec(),
            Self::Sha512 => sha512(data).to_vec(),
        }
    }
}

const fn xtime(x: u8) -> u8 {
    (x << 1) ^ if x & 0x80 != 0 { 0x1b } else { 0 }
}

const fn gmul(mut a: u8, mut b: u8) -> u8 {
    let mut r = 0;
    while b != 0 {
        if b & 1 != 0 {
            r ^= a;
        }
        a = xtime(a);
        b >>= 1;
    }
    r
}

const fn mul_table(by: u8) -> [u8; 256] {
    let mut t = [0u8; 256];
    let mut i = 0;
    while i < 256 {
        t[i] = gmul(i as u8, by);
        i += 1;
    }
    t
}

const MUL9: [u8; 256] = mul_table(9);
const MUL11: [u8; 256] = mul_table(11);
const MUL13: [u8; 256] = mul_table(13);
const MUL14: [u8; 256] = mul_table(14);

/// An AES key schedule; it decrypts ([FIPS 197]), and in tests encrypts, so a
/// round trip can check the schemes built on it.
pub(crate) struct Aes {
    round_keys: Vec<[u8; 16]>,
}

impl Aes {
    /// `None` unless `key` is 16, 24 or 32 bytes.
    pub(crate) fn new(key: &[u8]) -> Option<Self> {
        let nk = match key.len() {
            16 | 24 | 32 => key.len() / 4,
            _ => return None,
        };
        let rounds = nk + 6;
        let mut words: Vec<[u8; 4]> = key
            .as_chunks::<4>()
            .0
            .iter()
            .map(|c| [c[0], c[1], c[2], c[3]])
            .collect();
        let mut rcon = 1u8;
        for i in nk..4 * (rounds + 1) {
            let mut t = words[i - 1];
            if i % nk == 0 {
                t.rotate_left(1);
                for b in &mut t {
                    *b = SBOX[usize::from(*b)];
                }
                t[0] ^= rcon;
                rcon = xtime(rcon);
            } else if nk > 6 && i % nk == 4 {
                for b in &mut t {
                    *b = SBOX[usize::from(*b)];
                }
            }
            let prev = words[i - nk];
            words.push([
                t[0] ^ prev[0],
                t[1] ^ prev[1],
                t[2] ^ prev[2],
                t[3] ^ prev[3],
            ]);
        }
        let round_keys = words
            .as_chunks::<4>()
            .0
            .iter()
            .map(|w| {
                let mut k = [0u8; 16];
                for (dst, src) in k.as_chunks_mut::<4>().0.iter_mut().zip(w) {
                    dst.copy_from_slice(src);
                }
                k
            })
            .collect();
        Some(Self { round_keys })
    }

    fn decrypt_block(&self, block: &mut [u8; 16]) {
        let xor = |block: &mut [u8; 16], key: &[u8; 16]| {
            for (b, k) in block.iter_mut().zip(key) {
                *b ^= k;
            }
        };
        let inv_shift_sub = |block: &mut [u8; 16]| {
            let old = *block;
            for col in 0..4 {
                for row in 0..4 {
                    // Row `row` moves right by `row` columns.
                    block[4 * ((col + row) % 4) + row] = INV_SBOX[usize::from(old[4 * col + row])];
                }
            }
        };
        let last = self.round_keys.len() - 1;
        xor(block, &self.round_keys[last]);
        for round in (1..last).rev() {
            inv_shift_sub(block);
            xor(block, &self.round_keys[round]);
            for col in block.as_chunks_mut::<4>().0.iter_mut() {
                let [a, b, c, d] = [
                    usize::from(col[0]),
                    usize::from(col[1]),
                    usize::from(col[2]),
                    usize::from(col[3]),
                ];
                col[0] = MUL14[a] ^ MUL11[b] ^ MUL13[c] ^ MUL9[d];
                col[1] = MUL9[a] ^ MUL14[b] ^ MUL11[c] ^ MUL13[d];
                col[2] = MUL13[a] ^ MUL9[b] ^ MUL14[c] ^ MUL11[d];
                col[3] = MUL11[a] ^ MUL13[b] ^ MUL9[c] ^ MUL14[d];
            }
        }
        inv_shift_sub(block);
        xor(block, &self.round_keys[0]);
    }

    /// Encrypt one block (tests only).
    #[cfg(test)]
    pub(crate) fn encrypt_block(&self, block: &mut [u8; 16]) {
        let xor = |block: &mut [u8; 16], key: &[u8; 16]| {
            for (b, k) in block.iter_mut().zip(key) {
                *b ^= k;
            }
        };
        let last = self.round_keys.len() - 1;
        xor(block, &self.round_keys[0]);
        for round in 1..=last {
            let old = *block;
            for col in 0..4 {
                for row in 0..4 {
                    block[4 * col + row] = SBOX[usize::from(old[4 * ((col + row) % 4) + row])];
                }
            }
            if round != last {
                for col in block.as_chunks_mut::<4>().0.iter_mut() {
                    let [a, b, c, d] = [col[0], col[1], col[2], col[3]];
                    col[0] = xtime(a) ^ (xtime(b) ^ b) ^ c ^ d;
                    col[1] = a ^ xtime(b) ^ (xtime(c) ^ c) ^ d;
                    col[2] = a ^ b ^ xtime(c) ^ (xtime(d) ^ d);
                    col[3] = (xtime(a) ^ a) ^ b ^ c ^ xtime(d);
                }
            }
            xor(block, &self.round_keys[round]);
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

/// The RC4 stream cipher.
pub(crate) struct Rc4 {
    s: [u8; 256],
    i: u8,
    j: u8,
}

impl Rc4 {
    /// `None` for an empty key.
    pub(crate) fn new(key: &[u8]) -> Option<Self> {
        if key.is_empty() {
            return None;
        }
        let mut s = [0u8; 256];
        for (i, b) in s.iter_mut().enumerate() {
            *b = i as u8;
        }
        let mut j = 0u8;
        for i in 0..256 {
            j = j.wrapping_add(s[i]).wrapping_add(key[i % key.len()]);
            s.swap(i, usize::from(j));
        }
        Some(Self { s, i: 0, j: 0 })
    }

    fn next_byte(&mut self) -> u8 {
        self.i = self.i.wrapping_add(1);
        self.j = self.j.wrapping_add(self.s[usize::from(self.i)]);
        self.s.swap(usize::from(self.i), usize::from(self.j));
        let t = self.s[usize::from(self.i)].wrapping_add(self.s[usize::from(self.j)]);
        self.s[usize::from(t)]
    }

    /// XOR the next keystream bytes into `data`.
    pub(crate) fn apply(&mut self, data: &mut [u8]) {
        for b in data {
            *b ^= self.next_byte();
        }
    }
}

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
