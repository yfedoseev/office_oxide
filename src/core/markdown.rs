//! Markdown escaping shared by every renderer, so that document text
//! cannot inject structure into the rendered output.
//!
//! A cell containing `|` split its table row, a cell holding a newline
//! ended the table, `<Company Name>` became an (invisible) HTML tag and
//! `*not bold*` became emphasis — in the four direct table renderers,
//! which escaped nothing, and for the newline in the IR renderer too.

/// Escape the markdown metacharacters in a run of document text.
///
/// `_` is deliberately absent: CommonMark does not treat intra-word `_`
/// as emphasis, and escaping it turns ordinary identifiers like
/// `HEADER_TEXT` into unreadable `HEADER\_TEXT`.
pub fn escape_text(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        if matches!(c, '\\' | '`' | '*' | '[' | ']' | '<' | '>' | '|' | '~') {
            out.push('\\');
        }
        out.push(c);
    }
    out
}

/// Escape text for a GFM table cell: [`escape_text`] plus line breaks
/// as `<br>`, since a raw newline ends the row (and the table).
pub fn escape_cell(s: &str) -> String {
    let escaped = escape_text(s);
    let mut out = String::with_capacity(escaped.len());
    let mut parts = escaped.split(['\r', '\n']).filter(|p| !p.is_empty());
    if let Some(first) = parts.next() {
        out.push_str(first.trim());
        for p in parts {
            out.push_str("<br>");
            out.push_str(p.trim());
        }
    }
    out
}

/// A picture's alt text for `![alt](…)`: one line, brackets escaped —
/// a multi-line `descr` (`fig:\n\nlalune.jpg`) split the image syntax
/// across paragraphs.
pub fn image_alt(s: &str) -> String {
    let joined = s.split_whitespace().collect::<Vec<_>>().join(" ");
    joined.replace('[', "\\[").replace(']', "\\]")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_image_alt_is_one_line() {
        assert_eq!(image_alt("fig:\n\nlalune.jpg"), "fig: lalune.jpg");
        assert_eq!(image_alt("a [b]"), "a \\[b\\]");
    }

    #[test]
    fn test_cell_text_cannot_break_a_table() {
        assert_eq!(escape_cell("a|b"), "a\\|b");
        assert_eq!(escape_cell("line1\nline2"), "line1<br>line2");
        assert_eq!(escape_cell("line1\r\nline2\n"), "line1<br>line2");
        assert_eq!(escape_cell("<Company Name>"), "\\<Company Name\\>");
        assert_eq!(escape_cell("*not bold*"), "\\*not bold\\*");
        assert_eq!(escape_text("HEADER_TEXT 100%"), "HEADER_TEXT 100%");
    }
}
