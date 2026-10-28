//! The Adobe Symbol font's built-in encoding, for `w:sym` characters and the
//! like that name the Symbol font by a private-use code (`F0xx`).
//!
//! The table is generated, not typed: each code's glyph name comes from the
//! Adobe core-14 `Symbol.afm` (the `C <code> ; ... ; N <name> ;` lines) and
//! each name's character from the Adobe Glyph List (BSD-3-Clause,
//! github.com/adobe-type-tools/agl-aglfn). Three names are normalised to
//! the Greek letter the Symbol font draws (the list maps `Delta` to the
//! increment sign, `Omega` to the ohm sign and `mu` to the micro sign); the
//! copyright, registered and trademark glyphs, whose list entries are
//! private-use, map to the plain signs; and the bracket, brace and
//! parenthesis pieces, the integral extension and the two arrow extensions
//! map to their Unicode counterparts in Miscellaneous Technical. The one
//! glyph with no counterpart (`radicalex`, code 0x60) has no entry.

/// The character the Symbol font draws for `code`, `None` where the font has
/// no glyph (or none with a Unicode counterpart).
pub(crate) fn symbol_char(code: u8) -> Option<char> {
    match code {
        0x20 => Some('\u{0020}'), // space
        0x21 => Some('\u{0021}'), // exclam
        0x22 => Some('\u{2200}'), // universal
        0x23 => Some('\u{0023}'), // numbersign
        0x24 => Some('\u{2203}'), // existential
        0x25 => Some('\u{0025}'), // percent
        0x26 => Some('\u{0026}'), // ampersand
        0x27 => Some('\u{220B}'), // suchthat
        0x28 => Some('\u{0028}'), // parenleft
        0x29 => Some('\u{0029}'), // parenright
        0x2A => Some('\u{2217}'), // asteriskmath
        0x2B => Some('\u{002B}'), // plus
        0x2C => Some('\u{002C}'), // comma
        0x2D => Some('\u{2212}'), // minus
        0x2E => Some('\u{002E}'), // period
        0x2F => Some('\u{002F}'), // slash
        0x30 => Some('\u{0030}'), // zero
        0x31 => Some('\u{0031}'), // one
        0x32 => Some('\u{0032}'), // two
        0x33 => Some('\u{0033}'), // three
        0x34 => Some('\u{0034}'), // four
        0x35 => Some('\u{0035}'), // five
        0x36 => Some('\u{0036}'), // six
        0x37 => Some('\u{0037}'), // seven
        0x38 => Some('\u{0038}'), // eight
        0x39 => Some('\u{0039}'), // nine
        0x3A => Some('\u{003A}'), // colon
        0x3B => Some('\u{003B}'), // semicolon
        0x3C => Some('\u{003C}'), // less
        0x3D => Some('\u{003D}'), // equal
        0x3E => Some('\u{003E}'), // greater
        0x3F => Some('\u{003F}'), // question
        0x40 => Some('\u{2245}'), // congruent
        0x41 => Some('\u{0391}'), // Alpha
        0x42 => Some('\u{0392}'), // Beta
        0x43 => Some('\u{03A7}'), // Chi
        0x44 => Some('\u{0394}'), // Delta
        0x45 => Some('\u{0395}'), // Epsilon
        0x46 => Some('\u{03A6}'), // Phi
        0x47 => Some('\u{0393}'), // Gamma
        0x48 => Some('\u{0397}'), // Eta
        0x49 => Some('\u{0399}'), // Iota
        0x4A => Some('\u{03D1}'), // theta1
        0x4B => Some('\u{039A}'), // Kappa
        0x4C => Some('\u{039B}'), // Lambda
        0x4D => Some('\u{039C}'), // Mu
        0x4E => Some('\u{039D}'), // Nu
        0x4F => Some('\u{039F}'), // Omicron
        0x50 => Some('\u{03A0}'), // Pi
        0x51 => Some('\u{0398}'), // Theta
        0x52 => Some('\u{03A1}'), // Rho
        0x53 => Some('\u{03A3}'), // Sigma
        0x54 => Some('\u{03A4}'), // Tau
        0x55 => Some('\u{03A5}'), // Upsilon
        0x56 => Some('\u{03C2}'), // sigma1
        0x57 => Some('\u{03A9}'), // Omega
        0x58 => Some('\u{039E}'), // Xi
        0x59 => Some('\u{03A8}'), // Psi
        0x5A => Some('\u{0396}'), // Zeta
        0x5B => Some('\u{005B}'), // bracketleft
        0x5C => Some('\u{2234}'), // therefore
        0x5D => Some('\u{005D}'), // bracketright
        0x5E => Some('\u{22A5}'), // perpendicular
        0x5F => Some('\u{005F}'), // underscore
        0x61 => Some('\u{03B1}'), // alpha
        0x62 => Some('\u{03B2}'), // beta
        0x63 => Some('\u{03C7}'), // chi
        0x64 => Some('\u{03B4}'), // delta
        0x65 => Some('\u{03B5}'), // epsilon
        0x66 => Some('\u{03C6}'), // phi
        0x67 => Some('\u{03B3}'), // gamma
        0x68 => Some('\u{03B7}'), // eta
        0x69 => Some('\u{03B9}'), // iota
        0x6A => Some('\u{03D5}'), // phi1
        0x6B => Some('\u{03BA}'), // kappa
        0x6C => Some('\u{03BB}'), // lambda
        0x6D => Some('\u{03BC}'), // mu
        0x6E => Some('\u{03BD}'), // nu
        0x6F => Some('\u{03BF}'), // omicron
        0x70 => Some('\u{03C0}'), // pi
        0x71 => Some('\u{03B8}'), // theta
        0x72 => Some('\u{03C1}'), // rho
        0x73 => Some('\u{03C3}'), // sigma
        0x74 => Some('\u{03C4}'), // tau
        0x75 => Some('\u{03C5}'), // upsilon
        0x76 => Some('\u{03D6}'), // omega1
        0x77 => Some('\u{03C9}'), // omega
        0x78 => Some('\u{03BE}'), // xi
        0x79 => Some('\u{03C8}'), // psi
        0x7A => Some('\u{03B6}'), // zeta
        0x7B => Some('\u{007B}'), // braceleft
        0x7C => Some('\u{007C}'), // bar
        0x7D => Some('\u{007D}'), // braceright
        0x7E => Some('\u{223C}'), // similar
        0xA0 => Some('\u{20AC}'), // Euro
        0xA1 => Some('\u{03D2}'), // Upsilon1
        0xA2 => Some('\u{2032}'), // minute
        0xA3 => Some('\u{2264}'), // lessequal
        0xA4 => Some('\u{2044}'), // fraction
        0xA5 => Some('\u{221E}'), // infinity
        0xA6 => Some('\u{0192}'), // florin
        0xA7 => Some('\u{2663}'), // club
        0xA8 => Some('\u{2666}'), // diamond
        0xA9 => Some('\u{2665}'), // heart
        0xAA => Some('\u{2660}'), // spade
        0xAB => Some('\u{2194}'), // arrowboth
        0xAC => Some('\u{2190}'), // arrowleft
        0xAD => Some('\u{2191}'), // arrowup
        0xAE => Some('\u{2192}'), // arrowright
        0xAF => Some('\u{2193}'), // arrowdown
        0xB0 => Some('\u{00B0}'), // degree
        0xB1 => Some('\u{00B1}'), // plusminus
        0xB2 => Some('\u{2033}'), // second
        0xB3 => Some('\u{2265}'), // greaterequal
        0xB4 => Some('\u{00D7}'), // multiply
        0xB5 => Some('\u{221D}'), // proportional
        0xB6 => Some('\u{2202}'), // partialdiff
        0xB7 => Some('\u{2022}'), // bullet
        0xB8 => Some('\u{00F7}'), // divide
        0xB9 => Some('\u{2260}'), // notequal
        0xBA => Some('\u{2261}'), // equivalence
        0xBB => Some('\u{2248}'), // approxequal
        0xBC => Some('\u{2026}'), // ellipsis
        0xBD => Some('\u{23D0}'), // arrowvertex
        0xBE => Some('\u{23AF}'), // arrowhorizex
        0xBF => Some('\u{21B5}'), // carriagereturn
        0xC0 => Some('\u{2135}'), // aleph
        0xC1 => Some('\u{2111}'), // Ifraktur
        0xC2 => Some('\u{211C}'), // Rfraktur
        0xC3 => Some('\u{2118}'), // weierstrass
        0xC4 => Some('\u{2297}'), // circlemultiply
        0xC5 => Some('\u{2295}'), // circleplus
        0xC6 => Some('\u{2205}'), // emptyset
        0xC7 => Some('\u{2229}'), // intersection
        0xC8 => Some('\u{222A}'), // union
        0xC9 => Some('\u{2283}'), // propersuperset
        0xCA => Some('\u{2287}'), // reflexsuperset
        0xCB => Some('\u{2284}'), // notsubset
        0xCC => Some('\u{2282}'), // propersubset
        0xCD => Some('\u{2286}'), // reflexsubset
        0xCE => Some('\u{2208}'), // element
        0xCF => Some('\u{2209}'), // notelement
        0xD0 => Some('\u{2220}'), // angle
        0xD1 => Some('\u{2207}'), // gradient
        0xD2 => Some('\u{00AE}'), // registerserif
        0xD3 => Some('\u{00A9}'), // copyrightserif
        0xD4 => Some('\u{2122}'), // trademarkserif
        0xD5 => Some('\u{220F}'), // product
        0xD6 => Some('\u{221A}'), // radical
        0xD7 => Some('\u{22C5}'), // dotmath
        0xD8 => Some('\u{00AC}'), // logicalnot
        0xD9 => Some('\u{2227}'), // logicaland
        0xDA => Some('\u{2228}'), // logicalor
        0xDB => Some('\u{21D4}'), // arrowdblboth
        0xDC => Some('\u{21D0}'), // arrowdblleft
        0xDD => Some('\u{21D1}'), // arrowdblup
        0xDE => Some('\u{21D2}'), // arrowdblright
        0xDF => Some('\u{21D3}'), // arrowdbldown
        0xE0 => Some('\u{25CA}'), // lozenge
        0xE1 => Some('\u{2329}'), // angleleft
        0xE2 => Some('\u{00AE}'), // registersans
        0xE3 => Some('\u{00A9}'), // copyrightsans
        0xE4 => Some('\u{2122}'), // trademarksans
        0xE5 => Some('\u{2211}'), // summation
        0xE6 => Some('\u{239B}'), // parenlefttp
        0xE7 => Some('\u{239C}'), // parenleftex
        0xE8 => Some('\u{239D}'), // parenleftbt
        0xE9 => Some('\u{23A1}'), // bracketlefttp
        0xEA => Some('\u{23A2}'), // bracketleftex
        0xEB => Some('\u{23A3}'), // bracketleftbt
        0xEC => Some('\u{23A7}'), // bracelefttp
        0xED => Some('\u{23A8}'), // braceleftmid
        0xEE => Some('\u{23A9}'), // braceleftbt
        0xEF => Some('\u{23AA}'), // braceex
        0xF1 => Some('\u{232A}'), // angleright
        0xF2 => Some('\u{222B}'), // integral
        0xF3 => Some('\u{2320}'), // integraltp
        0xF4 => Some('\u{23AE}'), // integralex
        0xF5 => Some('\u{2321}'), // integralbt
        0xF6 => Some('\u{239E}'), // parenrighttp
        0xF7 => Some('\u{239F}'), // parenrightex
        0xF8 => Some('\u{23A0}'), // parenrightbt
        0xF9 => Some('\u{23A4}'), // bracketrighttp
        0xFA => Some('\u{23A5}'), // bracketrightex
        0xFB => Some('\u{23A6}'), // bracketrightbt
        0xFC => Some('\u{23AB}'), // bracerighttp
        0xFD => Some('\u{23AC}'), // bracerightmid
        0xFE => Some('\u{23AD}'), // bracerightbt
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::symbol_char;

    /// Codes whose glyph is easy to confuse with a Latin-1 or Windows-1252
    /// reading of the same byte: the cases the reader used to get wrong.
    #[test]
    fn test_symbol_codes_are_not_read_as_latin_1() {
        assert_eq!(symbol_char(0xD1), Some('\u{2207}')); // nabla, not N-tilde
        assert_eq!(symbol_char(0xA7), Some('\u{2663}')); // club, not section sign
        assert_eq!(symbol_char(0xB4), Some('\u{00D7}')); // multiply
        assert_eq!(symbol_char(0xB8), Some('\u{00F7}')); // divide
        assert_eq!(symbol_char(0xB5), Some('\u{221D}')); // proportional
        assert_eq!(symbol_char(0xD0), Some('\u{2220}')); // angle
    }

    #[test]
    fn test_greek_letters_and_the_names_the_glyph_list_maps_elsewhere() {
        assert_eq!(symbol_char(0x61), Some('\u{03B1}')); // alpha
        assert_eq!(symbol_char(0x70), Some('\u{03C0}')); // pi
        assert_eq!(symbol_char(0x44), Some('\u{0394}')); // Delta, not U+2206
        assert_eq!(symbol_char(0x57), Some('\u{03A9}')); // Omega, not U+2126
        assert_eq!(symbol_char(0x6D), Some('\u{03BC}')); // mu, not U+00B5
    }

    #[test]
    fn test_bullet_digits_and_gaps() {
        assert_eq!(symbol_char(0xB7), Some('\u{2022}')); // bullet
        assert_eq!(symbol_char(0x31), Some('1'));
        assert_eq!(symbol_char(0x20), Some(' '));
        assert_eq!(symbol_char(0xA0), Some('\u{20AC}')); // Euro
        assert_eq!(symbol_char(0x60), None); // radicalex has no counterpart
        assert_eq!(symbol_char(0x00), None);
        assert_eq!(symbol_char(0xFF), None);
    }

    /// Every defined code maps to a distinct, non-private-use character.
    #[test]
    fn test_defined_codes_map_to_distinct_public_characters() {
        let mut seen = std::collections::HashSet::new();
        let mut defined = 0;
        for code in 0..=255u8 {
            if let Some(c) = symbol_char(code) {
                defined += 1;
                assert!(!('\u{E000}'..='\u{F8FF}').contains(&c), "{code:#04x}");
                // Registered, copyright and trademark exist twice in the font (serif and sans).
                if !matches!(code, 0xD2 | 0xD3 | 0xD4 | 0xE2 | 0xE3 | 0xE4) {
                    assert!(seen.insert(c), "{code:#04x} repeats {c:?}");
                }
            }
        }
        assert_eq!(defined, 188);
    }
}
