//! The encodings of the symbol fonts Word names in `<w:sym>` (`F0xx` private-use
//! codes): Symbol, Wingdings, Wingdings 2, Wingdings 3 and Webdings.
//!
//! **Symbol** is generated, not typed: each code's glyph name comes from the
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
//!
//! **Wingdings, Wingdings 2, Wingdings 3 and Webdings** are the normative
//! source references of the Unicode Consortium: the `UCS;ID` list that closes
//! "Status of encoding of Wingdings and Webdings Symbols" (UTC L2/12-368, ISO/IEC
//! JTC1/SC2/WG2 N4384), which gives the Unicode character for each of the 866
//! encoded glyphs of the four fonts. An ID is `1000 * font + code` (Webdings 0,
//! Wingdings 1, Wingdings 2 is 2, Wingdings 3 is 3). A code with no entry in
//! that list has no character here, and `0x20` is a space in every font.

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

/// A symbol font with a table here.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SymbolFont {
    Symbol,
    Wingdings,
    Wingdings2,
    Wingdings3,
    Webdings,
}

/// The symbol font a `w:font` name denotes, matched without regard to case or
/// spacing (`Wingdings 2`, `WINGDINGS2`, ` Wingdings `).
pub(crate) fn symbol_font(name: &str) -> Option<SymbolFont> {
    let compact: String = name
        .chars()
        .filter(|c| !c.is_whitespace())
        .flat_map(char::to_lowercase)
        .collect();
    match compact.as_str() {
        "symbol" => Some(SymbolFont::Symbol),
        "wingdings" => Some(SymbolFont::Wingdings),
        "wingdings2" => Some(SymbolFont::Wingdings2),
        "wingdings3" => Some(SymbolFont::Wingdings3),
        "webdings" => Some(SymbolFont::Webdings),
        _ => None,
    }
}

/// The character `font` draws for `code`, `None` where it has none.
pub(crate) fn char_for(font: SymbolFont, code: u8) -> Option<char> {
    let table: &[u32; 224] = match font {
        SymbolFont::Symbol => return symbol_char(code),
        SymbolFont::Wingdings => &WINGDINGS,
        SymbolFont::Wingdings2 => &WINGDINGS_2,
        SymbolFont::Wingdings3 => &WINGDINGS_3,
        SymbolFont::Webdings => &WEBDINGS,
    };
    let index = usize::from(code).checked_sub(0x20)?;
    table
        .get(index)
        .copied()
        .filter(|&u| u != 0)
        .and_then(char::from_u32)
}

/// Wingdings, codes 0x20..=0xFF; 0 where the font has no character.
const WINGDINGS: [u32; 224] = [
    0x00020, 0x1F589, 0x02702, 0x02701, 0x1F453, 0x1F56D, 0x1F56E, 0x1F56F, 0x1F57F, 0x02706,
    0x1F582, 0x1F583, 0x1F4EA, 0x1F4EB, 0x1F4EC, 0x1F4ED, 0x1F5C0, 0x1F5C1, 0x1F5CE, 0x1F5CF,
    0x1F5D0, 0x1F5C4, 0x023F3, 0x1F5AE, 0x1F5B0, 0x1F5B2, 0x1F5B3, 0x1F5B4, 0x1F5AB, 0x1F5AC,
    0x02707, 0x0270D, 0x1F58E, 0x0270C, 0x1F58F, 0x1F44D, 0x1F44E, 0x0261C, 0x0261E, 0x0261D,
    0x1F597, 0x1F590, 0x0263A, 0x1F610, 0x02639, 0x1F4A3, 0x1F571, 0x1F3F3, 0x1F3F1, 0x02708,
    0x0263C, 0x1F322, 0x02744, 0x1F546, 0x0271E, 0x1F548, 0x02720, 0x02721, 0x0262A, 0x0262F,
    0x1F549, 0x02638, 0x02648, 0x02649, 0x0264A, 0x0264B, 0x0264C, 0x0264D, 0x0264E, 0x0264F,
    0x02650, 0x02651, 0x02652, 0x02653, 0x1F670, 0x1F675, 0x026AB, 0x1F53E, 0x025FC, 0x1F78F,
    0x1F790, 0x02751, 0x02752, 0x1F79F, 0x029EB, 0x025C6, 0x02756, 0x1F799, 0x02327, 0x02BB9,
    0x02318, 0x1F3F5, 0x1F3F6, 0x1F676, 0x1F677, 0x00000, 0x1F10B, 0x02780, 0x02781, 0x02782,
    0x02783, 0x02784, 0x02785, 0x02786, 0x02787, 0x02788, 0x02789, 0x1F10C, 0x0278A, 0x0278B,
    0x0278C, 0x0278D, 0x0278E, 0x0278F, 0x02790, 0x02791, 0x02792, 0x02793, 0x1F662, 0x1F660,
    0x1F661, 0x1F663, 0x1F65E, 0x1F65C, 0x1F65D, 0x1F65F, 0x02219, 0x02022, 0x02B1D, 0x02B58,
    0x1F786, 0x1F788, 0x1F78A, 0x1F78B, 0x1F53F, 0x025AA, 0x1F78E, 0x1F7C1, 0x1F7C5, 0x02605,
    0x1F7CB, 0x1F7CF, 0x1F7D3, 0x1F7D1, 0x02BD0, 0x02316, 0x02BCE, 0x02BCF, 0x02BD1, 0x0272A,
    0x02730, 0x1F550, 0x1F551, 0x1F552, 0x1F553, 0x1F554, 0x1F555, 0x1F556, 0x1F557, 0x1F558,
    0x1F559, 0x1F55A, 0x1F55B, 0x02BB0, 0x02BB1, 0x02BB2, 0x02BB3, 0x02BB4, 0x02BB5, 0x02BB6,
    0x02BB7, 0x1F66A, 0x1F66B, 0x1F655, 0x1F654, 0x1F657, 0x1F656, 0x1F650, 0x1F651, 0x1F652,
    0x1F653, 0x0232B, 0x02326, 0x02B98, 0x02B9A, 0x02B99, 0x02B9B, 0x02B88, 0x02B8A, 0x02B89,
    0x02B8B, 0x1F868, 0x1F86A, 0x1F869, 0x1F86B, 0x1F86C, 0x1F86D, 0x1F86F, 0x1F86E, 0x1F878,
    0x1F87A, 0x1F879, 0x1F87B, 0x1F87C, 0x1F87D, 0x1F87F, 0x1F87E, 0x021E6, 0x021E8, 0x021E7,
    0x021E9, 0x02B04, 0x021F3, 0x02B01, 0x02B00, 0x02B03, 0x02B02, 0x1F8AC, 0x1F8AD, 0x1F5F6,
    0x02713, 0x1F5F7, 0x1F5F9, 0x00000,
];

/// Wingdings 2, codes 0x20..=0xFF; 0 where the font has no character.
const WINGDINGS_2: [u32; 224] = [
    0x00020, 0x1F58A, 0x1F58B, 0x1F58C, 0x1F58D, 0x02704, 0x02700, 0x1F57E, 0x1F57D, 0x1F5C5,
    0x1F5C6, 0x1F5C7, 0x1F5C8, 0x1F5C9, 0x1F5CA, 0x1F5CB, 0x1F5CC, 0x1F5CD, 0x1F4CB, 0x1F5D1,
    0x1F5D4, 0x1F5B5, 0x1F5B6, 0x1F5B7, 0x1F5B8, 0x1F5AD, 0x1F5AF, 0x1F5B1, 0x1F592, 0x1F593,
    0x1F598, 0x1F599, 0x1F59A, 0x1F59B, 0x1F448, 0x1F449, 0x1F59C, 0x1F59D, 0x1F59E, 0x1F59F,
    0x1F5A0, 0x1F5A1, 0x1F446, 0x1F447, 0x1F5A2, 0x1F5A3, 0x1F591, 0x1F5F4, 0x1F5F8, 0x1F5F5,
    0x02611, 0x02BBD, 0x02612, 0x02BBE, 0x02BBF, 0x1F6C7, 0x029B8, 0x1F671, 0x1F674, 0x1F672,
    0x1F673, 0x0203D, 0x1F679, 0x1F67A, 0x1F67B, 0x1F666, 0x1F664, 0x1F665, 0x1F667, 0x1F65A,
    0x1F658, 0x1F659, 0x1F65B, 0x024EA, 0x02460, 0x02461, 0x02462, 0x02463, 0x02464, 0x02465,
    0x02466, 0x02467, 0x02468, 0x02469, 0x024FF, 0x02776, 0x02777, 0x02778, 0x02779, 0x0277A,
    0x0277B, 0x0277C, 0x0277D, 0x0277E, 0x0277F, 0x00000, 0x02609, 0x1F315, 0x0263D, 0x0263E,
    0x02E3F, 0x0271D, 0x1F547, 0x1F55C, 0x1F55D, 0x1F55E, 0x1F55F, 0x1F560, 0x1F561, 0x1F562,
    0x1F563, 0x1F564, 0x1F565, 0x1F566, 0x1F567, 0x1F668, 0x1F669, 0x022C5, 0x1F784, 0x02981,
    0x025CF, 0x025CB, 0x1F785, 0x1F787, 0x1F789, 0x02299, 0x029BF, 0x1F78C, 0x1F78D, 0x025FE,
    0x025A0, 0x025A1, 0x1F791, 0x1F792, 0x1F793, 0x1F794, 0x025A3, 0x1F795, 0x1F796, 0x1F797,
    0x1F798, 0x02B29, 0x02B25, 0x025C7, 0x1F79A, 0x025C8, 0x1F79B, 0x1F79C, 0x1F79D, 0x1F79E,
    0x02B2A, 0x02B27, 0x025CA, 0x1F7A0, 0x025D6, 0x025D7, 0x02BCA, 0x02BCB, 0x02BC0, 0x02BC1,
    0x02B1F, 0x02BC2, 0x02B23, 0x02B22, 0x02BC3, 0x02BC4, 0x1F7A1, 0x1F7A2, 0x1F7A3, 0x1F7A4,
    0x1F7A5, 0x1F7A6, 0x1F7A7, 0x1F7A8, 0x1F7A9, 0x1F7AA, 0x1F7AB, 0x1F7AC, 0x1F7AD, 0x1F7AE,
    0x1F7AF, 0x1F7B0, 0x1F7B1, 0x1F7B2, 0x1F7B3, 0x1F7B4, 0x1F7B5, 0x1F7B6, 0x1F7B7, 0x1F7B8,
    0x1F7B9, 0x1F7BA, 0x1F7BB, 0x1F7BC, 0x1F7BD, 0x1F7BE, 0x1F7BF, 0x1F7C0, 0x1F7C2, 0x1F7C4,
    0x1F7C6, 0x1F7C9, 0x1F7CA, 0x02736, 0x1F7CC, 0x1F7CE, 0x1F7D0, 0x1F7D2, 0x02739, 0x1F7C3,
    0x1F7C7, 0x0272F, 0x1F7CD, 0x1F7D4, 0x02BCC, 0x02BCD, 0x0203B, 0x02042, 0x00000, 0x00000,
    0x00000, 0x00000, 0x00000, 0x00000,
];

/// Wingdings 3, codes 0x20..=0xFF; 0 where the font has no character.
const WINGDINGS_3: [u32; 224] = [
    0x00020, 0x02B60, 0x02B62, 0x02B61, 0x02B63, 0x02B66, 0x02B67, 0x02B69, 0x02B68, 0x02B70,
    0x02B72, 0x02B71, 0x02B73, 0x02B76, 0x02B78, 0x02B7B, 0x02B7D, 0x02B64, 0x02B65, 0x02B6A,
    0x02B6C, 0x02B6B, 0x02B6D, 0x02B4D, 0x02BA0, 0x02BA1, 0x02BA2, 0x02BA3, 0x02BA4, 0x02BA5,
    0x02BA6, 0x02BA7, 0x02B90, 0x02B91, 0x02B92, 0x02B93, 0x02B80, 0x02B83, 0x02B7E, 0x02B7F,
    0x02B84, 0x02B86, 0x02B85, 0x02B87, 0x02B8F, 0x02B8D, 0x02B8E, 0x02B8C, 0x02B6E, 0x02B6F,
    0x0238B, 0x02324, 0x02303, 0x02325, 0x02423, 0x0237D, 0x021EA, 0x02BB8, 0x1F8A0, 0x1F8A1,
    0x1F8A2, 0x1F8A3, 0x1F8A4, 0x1F8A5, 0x1F8A6, 0x1F8A7, 0x1F8A8, 0x1F8A9, 0x1F8AA, 0x1F8AB,
    0x1F850, 0x1F852, 0x1F851, 0x1F853, 0x1F854, 0x1F855, 0x1F857, 0x1F856, 0x1F858, 0x1F859,
    0x025B2, 0x025BC, 0x025B3, 0x025BD, 0x025C0, 0x025B6, 0x025C1, 0x025B7, 0x025E3, 0x025E2,
    0x025E4, 0x025E5, 0x1F780, 0x1F782, 0x1F781, 0x00000, 0x1F783, 0x02BC5, 0x02BC6, 0x02BC7,
    0x02BC8, 0x02B9C, 0x02B9E, 0x02B9D, 0x02B9F, 0x1F810, 0x1F812, 0x1F811, 0x1F813, 0x1F814,
    0x1F816, 0x1F815, 0x1F817, 0x1F818, 0x1F81A, 0x1F819, 0x1F81B, 0x1F81C, 0x1F81E, 0x1F81D,
    0x1F81F, 0x1F800, 0x1F802, 0x1F801, 0x1F803, 0x1F804, 0x1F806, 0x1F805, 0x1F807, 0x1F808,
    0x1F80A, 0x1F809, 0x1F80B, 0x1F820, 0x1F822, 0x1F824, 0x1F826, 0x1F828, 0x1F82A, 0x1F82C,
    0x1F89C, 0x1F89D, 0x1F89E, 0x1F89F, 0x1F82E, 0x1F830, 0x1F832, 0x1F834, 0x1F836, 0x1F838,
    0x1F83A, 0x1F839, 0x1F83B, 0x1F898, 0x1F89A, 0x1F899, 0x1F89B, 0x1F83C, 0x1F83E, 0x1F83D,
    0x1F83F, 0x1F840, 0x1F842, 0x1F841, 0x1F843, 0x1F844, 0x1F846, 0x1F845, 0x1F847, 0x02BA8,
    0x02BA9, 0x02BAA, 0x02BAB, 0x02BAC, 0x02BAD, 0x02BAE, 0x02BAF, 0x1F860, 0x1F862, 0x1F861,
    0x1F863, 0x1F864, 0x1F865, 0x1F867, 0x1F866, 0x1F870, 0x1F872, 0x1F871, 0x1F873, 0x1F874,
    0x1F875, 0x1F877, 0x1F876, 0x1F880, 0x1F882, 0x1F881, 0x1F883, 0x1F884, 0x1F885, 0x1F887,
    0x1F886, 0x1F890, 0x1F892, 0x1F891, 0x1F893, 0x1F894, 0x1F896, 0x1F895, 0x1F897, 0x00000,
    0x00000, 0x00000, 0x00000, 0x00000, 0x00000, 0x00000, 0x00000, 0x00000, 0x00000, 0x00000,
    0x00000, 0x00000, 0x00000, 0x00000,
];

/// Webdings, codes 0x20..=0xFF; 0 where the font has no character.
const WEBDINGS: [u32; 224] = [
    0x00020, 0x1F577, 0x1F578, 0x1F572, 0x1F576, 0x1F3C6, 0x1F396, 0x1F587, 0x1F5E8, 0x1F5E9,
    0x1F5F0, 0x1F5F1, 0x1F336, 0x1F397, 0x1F67E, 0x1F67C, 0x1F5D5, 0x1F5D6, 0x1F5D7, 0x023F4,
    0x023F5, 0x023F6, 0x023F7, 0x023EA, 0x023E9, 0x023EE, 0x023ED, 0x023F8, 0x023F9, 0x023FA,
    0x1F5DA, 0x1F5F3, 0x1F6E0, 0x1F3D7, 0x1F3D8, 0x1F3D9, 0x1F3DA, 0x1F3DC, 0x1F3ED, 0x1F3DB,
    0x1F3E0, 0x1F3D6, 0x1F3DD, 0x1F6E3, 0x1F50D, 0x1F3D4, 0x1F441, 0x1F442, 0x1F3DE, 0x1F3D5,
    0x1F6E4, 0x1F3DF, 0x1F6F3, 0x1F56C, 0x1F56B, 0x1F568, 0x1F508, 0x1F394, 0x1F395, 0x1F5EC,
    0x1F67D, 0x1F5ED, 0x1F5EA, 0x1F5EB, 0x02B94, 0x02714, 0x1F6B2, 0x02B1C, 0x1F6E1, 0x1F4E6,
    0x1F6F1, 0x02B1B, 0x1F691, 0x1F6C8, 0x1F6E9, 0x1F6F0, 0x1F7C8, 0x1F574, 0x02B24, 0x1F6E5,
    0x1F694, 0x1F5D8, 0x1F5D9, 0x02753, 0x1F6F2, 0x1F687, 0x1F68D, 0x026F3, 0x029B8, 0x02296,
    0x1F6AD, 0x1F5EE, 0x023D0, 0x1F5EF, 0x1F5F2, 0x00000, 0x1F6B9, 0x1F6BA, 0x1F6C9, 0x1F6CA,
    0x1F6BC, 0x1F47D, 0x1F3CB, 0x026F7, 0x1F3C2, 0x1F3CC, 0x1F3CA, 0x1F3C4, 0x1F3CD, 0x1F3CE,
    0x1F698, 0x1F5E0, 0x1F6E2, 0x1F4B0, 0x1F3F7, 0x1F4B3, 0x1F46A, 0x1F5E1, 0x1F5E2, 0x1F5E3,
    0x0272F, 0x1F584, 0x1F585, 0x1F583, 0x1F586, 0x1F5B9, 0x1F5BA, 0x1F5BB, 0x1F575, 0x1F570,
    0x1F5BD, 0x1F5BE, 0x1F4CB, 0x1F5D2, 0x1F5D3, 0x1F56E, 0x1F4DA, 0x1F5DE, 0x1F5DF, 0x1F5C3,
    0x1F5C2, 0x1F5BC, 0x1F3AD, 0x1F39C, 0x1F398, 0x1F399, 0x1F3A7, 0x1F4BF, 0x1F39E, 0x1F4F7,
    0x1F39F, 0x1F3AC, 0x1F4FD, 0x1F4F9, 0x1F4FE, 0x1F4FB, 0x1F39A, 0x1F39B, 0x1F4FA, 0x1F4BB,
    0x1F5A5, 0x1F5A6, 0x1F5A7, 0x1F579, 0x1F3AE, 0x1F57B, 0x1F57C, 0x1F4DF, 0x1F581, 0x1F580,
    0x1F5A8, 0x1F5A9, 0x1F5BF, 0x1F5AA, 0x1F5DC, 0x1F512, 0x1F513, 0x1F5DD, 0x1F4E5, 0x1F4E4,
    0x1F573, 0x1F323, 0x1F324, 0x1F325, 0x1F326, 0x02601, 0x1F328, 0x1F327, 0x1F329, 0x1F32A,
    0x1F32C, 0x1F32B, 0x1F31C, 0x1F321, 0x1F6CB, 0x1F6CF, 0x1F37D, 0x1F378, 0x1F6CE, 0x1F6CD,
    0x024C5, 0x0267F, 0x1F6C6, 0x1F588, 0x1F393, 0x1F5E4, 0x1F5E5, 0x1F5E6, 0x1F5E7, 0x1F6EA,
    0x1F43F, 0x1F426, 0x1F41F, 0x1F415, 0x1F408, 0x1F66C, 0x1F66E, 0x1F66D, 0x1F66F, 0x1F5FA,
    0x1F30D, 0x1F30F, 0x1F30E, 0x1F54A,
];

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

    #[test]
    fn test_font_names_are_matched_without_regard_to_case_or_spacing() {
        use super::{SymbolFont, symbol_font};
        assert_eq!(symbol_font("Symbol"), Some(SymbolFont::Symbol));
        assert_eq!(symbol_font(" SYMBOL "), Some(SymbolFont::Symbol));
        assert_eq!(symbol_font("Wingdings"), Some(SymbolFont::Wingdings));
        assert_eq!(symbol_font("Wingdings 2"), Some(SymbolFont::Wingdings2));
        assert_eq!(symbol_font("wingdings2"), Some(SymbolFont::Wingdings2));
        assert_eq!(symbol_font("Wingdings 3"), Some(SymbolFont::Wingdings3));
        assert_eq!(symbol_font("Webdings"), Some(SymbolFont::Webdings));
        assert_eq!(symbol_font("Arial"), None);
        assert_eq!(symbol_font("Symbol MT"), None);
        assert_eq!(symbol_font(""), None);
    }

    /// Every table maps its defined codes to distinct, public characters, and
    /// the counts are those of the source mappings.
    #[test]
    fn test_symbol_font_tables_are_complete_and_public() {
        use super::{SymbolFont::*, char_for};
        for (font, defined) in [
            (Wingdings, 222usize),
            (Wingdings2, 217),
            (Wingdings3, 208),
            (Webdings, 223),
        ] {
            let mut seen = std::collections::HashSet::new();
            let mut count = 0;
            for code in 0..=255u8 {
                if let Some(c) = char_for(font, code) {
                    count += 1;
                    assert!(!('\u{E000}'..='\u{F8FF}').contains(&c), "{font:?} {code:#04x}");
                    if code != 0x20 {
                        assert!(c as u32 > 0x7F, "{font:?} {code:#04x}");
                        let _ = seen.insert(c);
                    }
                }
            }
            assert_eq!(count, defined, "{font:?}");
            assert!(seen.len() + 8 >= defined, "{font:?}: {} distinct", seen.len());
            assert_eq!(char_for(font, 0x20), Some(' '));
            assert_eq!(char_for(font, 0x1F), None);
        }
    }

    #[test]
    fn test_known_wingdings_codes() {
        use super::{SymbolFont::*, char_for};
        assert_eq!(char_for(Wingdings, 0xFC), Some('\u{2713}')); // check mark
        assert_eq!(char_for(Wingdings, 0x6E), Some('\u{25FC}')); // black medium square
        assert_eq!(char_for(Wingdings, 0xA7), Some('\u{25AA}')); // black small square
        assert_eq!(char_for(Wingdings2, 0x52), Some('\u{2611}')); // ballot box with check
        assert_eq!(char_for(Webdings, 0x61), Some('\u{2714}')); // heavy check mark
        // Where the Unicode list differs from the nearest older character.
        assert_eq!(char_for(Wingdings, 0x30), Some('\u{1F5C0}')); // folder
        assert_eq!(char_for(Wingdings2, 0x50), Some('\u{1F5F8}')); // light check mark
        assert_eq!(char_for(Wingdings, 0x4E), Some('\u{1F571}')); // black skull and crossbones
        // Codes the list does not name.
        assert_eq!(char_for(Wingdings, 0xFF), None);
        assert_eq!(char_for(Wingdings3, 0xF1), None);
        assert_eq!(char_for(Webdings, 0x7F), None);
        assert_eq!(char_for(Symbol, 0xD1), Some('\u{2207}')); // dispatches to the Symbol table
    }
}
