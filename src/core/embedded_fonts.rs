//! Embedded font helpers shared by DOCX, PPTX, and XLSX writers.
//!
//! All three Office formats can carry TrueType / OpenType font programs
//! inside their package so a downstream consumer can render with the
//! original typeface even when the system doesn't have it installed.
//! The OPC layout is identical across formats:
//!
//! - DOCX:  `/word/fonts/font_<n>_<safe_name>.ttf`
//! - PPTX:  `/ppt/fonts/font_<n>_<safe_name>.ttf`
//! - XLSX:  `/xl/fonts/font_<n>_<safe_name>.ttf`
//!
//! Other apps (Word, PowerPoint, Excel) require additional manifest
//! plumbing (`<w:embeddedFontLst>`, `<p:embeddedFontLst>`, etc.) to
//! actually pick up the embed; until that lands the in-process reader
//! is the only consumer. It scans the `*/fonts/` directory directly,
//! which is why the layout is uniform across formats.
//!
//! `sanitize_font_filename` strips characters that aren't legal in OPC
//! part names so font names can be embedded into the path safely.

use super::Result;
use super::opc::{OpcWriter, PartName};
use std::io::{Seek, Write};

/// Generic content type for embedded font payloads. The package
/// remains valid OPC even though Word/PowerPoint/Excel won't
/// auto-discover the font without the per-format manifest entries.
const FONT_CONTENT_TYPE: &str = "application/x-font-ttf";

/// Strip path-unsafe characters from a font name so it can live
/// inside an OPC part name (`/word/fonts/font_<n>_<safe_name>.ttf`).
/// Keeps ASCII alphanumeric, `-`, and `_`; replaces everything else
/// with `_` and clamps to 40 characters.
pub fn sanitize_font_filename(name: &str) -> String {
    name.chars()
        .map(|c| {
            if c.is_ascii_alphanumeric() || c == '-' || c == '_' {
                c
            } else {
                '_'
            }
        })
        .take(40)
        .collect()
}

/// Write the supplied font programs into the OPC package under the
/// given path prefix (e.g. `/word/fonts/`, `/ppt/fonts/`, or
/// `/xl/fonts/`). Each entry becomes
/// `<prefix>font_<n>_<safe_name>.ttf` with `n` starting at 1.
///
/// `prefix` must end with `/` and start with `/`.
pub fn write_embedded_fonts<W: Write + Seek>(
    opc: &mut OpcWriter<W>,
    prefix: &str,
    fonts: &[(String, Vec<u8>)],
) -> Result<()> {
    debug_assert!(prefix.starts_with('/') && prefix.ends_with('/'));
    if !fonts.is_empty() {
        // Register `ttf` once as a Default content-type entry. The
        // per-part Overrides we emit alongside still take precedence
        // at lookup; the Default just keeps OOXML SDK validators
        // happy ("missing Default for extension ttf").
        opc.register_default_content_type("ttf", FONT_CONTENT_TYPE);
    }
    for (idx, (name, data)) in fonts.iter().enumerate() {
        let n = idx + 1;
        let safe_name = sanitize_font_filename(name);
        let target = format!("{prefix}font_{n}_{safe_name}.ttf");
        let part = PartName::new(&target)?;
        opc.add_part(&part, FONT_CONTENT_TYPE, data)?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sanitize_keeps_alphanumeric() {
        assert_eq!(sanitize_font_filename("Calibri"), "Calibri");
        assert_eq!(sanitize_font_filename("Arial123"), "Arial123");
    }

    #[test]
    fn sanitize_keeps_dash_and_underscore() {
        assert_eq!(sanitize_font_filename("Times-Roman"), "Times-Roman");
        assert_eq!(sanitize_font_filename("TeXGyreTermesX-Regular"), "TeXGyreTermesX-Regular");
        assert_eq!(sanitize_font_filename("my_font"), "my_font");
    }

    #[test]
    fn sanitize_replaces_path_unsafe_chars() {
        assert_eq!(sanitize_font_filename("Arial/Bold"), "Arial_Bold");
        assert_eq!(sanitize_font_filename("a*b?c"), "a_b_c");
        assert_eq!(sanitize_font_filename("Noto Sans"), "Noto_Sans");
        assert_eq!(sanitize_font_filename("a.b"), "a_b");
    }

    #[test]
    fn sanitize_replaces_non_ascii() {
        // Non-ASCII alphanumeric is replaced with '_'.
        assert_eq!(sanitize_font_filename("Café"), "Caf_");
    }

    #[test]
    fn sanitize_clamps_to_40_chars() {
        let long = "A".repeat(100);
        let s = sanitize_font_filename(&long);
        assert_eq!(s.len(), 40);
        assert!(s.chars().all(|c| c == 'A'));
    }

    #[test]
    fn sanitize_empty_input() {
        assert_eq!(sanitize_font_filename(""), "");
    }
}
