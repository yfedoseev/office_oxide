//! Excel number format rendering.
//!
//! Applies a numeric format string (or built-in format ID) to an f64 value
//! and returns the display string. Covers the cases that matter in practice:
//! integers, fixed decimals, thousands separators, percentages, currency,
//! and scientific notation. Complex conditions/colors are stripped gracefully.

/// Return the canonical format code for a built-in number-format ID
/// (OOXML §18.8.30 reserved IDs 0–49). Custom formats (ID ≥ 164) are not
/// built-in and are looked up from the workbook's `<numFmts>` table instead,
/// so this returns `None` for them. Used to surface a human-readable format
/// string in the IR even when a cell uses a built-in date/number format that
/// never appears in `<numFmts>`.
pub fn builtin_format_code(fmt_id: u32) -> Option<&'static str> {
    let code = match fmt_id {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        5 => "$#,##0",
        6 => "$#,##0;[Red]$#,##0",
        7 => "$#,##0.00",
        8 => "$#,##0.00;[Red]$#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "m/d/yyyy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yyyy h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mm:ss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => return None,
    };
    Some(code)
}

/// Apply an Excel number format to a numeric value.
pub fn apply_format(n: f64, fmt_id: u32, fmt_str: Option<&str>) -> String {
    if n.is_nan() {
        return "NaN".to_string();
    }
    if n.is_infinite() {
        return if n < 0.0 {
            "-Infinity".to_string()
        } else {
            "Infinity".to_string()
        };
    }

    // Built-in format IDs per OOXML spec §18.8.30.
    match fmt_id {
        0 | 49 => return format_general(n),         // General / @
        1 => return format_integer(n),              // 0
        2 => return format_fixed(n, 2),             // 0.00
        3 => return format_commas(n, 0),            // #,##0
        4 => return format_commas(n, 2),            // #,##0.00
        5 | 6 => return format_currency(n, "$", 0), // $#,##0
        7 | 8 => return format_currency(n, "$", 2), // $#,##0.00
        9 => return format_percent(n, 0),           // 0%
        10 => return format_percent(n, 2),          // 0.00%
        11 => return format_scientific(n),          // 0.00E+00
        12 => return format_general(n),             // # ?/? (fractions — approx)
        13 => return format_general(n),             // # ??/??
        37 | 38 => return format_commas(n, 0),      // #,##0 accounting variants
        39 | 40 => return format_commas(n, 2),      // #,##0.00 accounting variants
        41..=44 => return format_commas(n, 2),      // _(* ...) accounting
        _ => {},
    }

    // Custom format string (IDs 164+).
    if let Some(fmt) = fmt_str {
        let fmt = fmt.trim();
        // The General/text sentinels are matched case-insensitively, and
        // after leading `[...]` directives are stripped: real workbooks
        // declare `numFmt formatCode="GENERAL"` and
        // `"[DBNum1][$-804]General"`, and treating either as a literal
        // format code dropped the cell's value and printed the code.
        let bare = strip_leading_directives(fmt);
        if !fmt.is_empty() && !bare.eq_ignore_ascii_case("General") && fmt != "@" {
            return apply_custom(n, fmt);
        }
    }

    format_general(n)
}

// ── Simple format primitives ───────────────────────────────────────────────

/// Format a number using Excel's General format (integer if whole, float otherwise).
pub fn format_general(n: f64) -> String {
    if n == n.trunc() && n.abs() < 1e15 {
        format!("{}", n as i64)
    } else {
        // Trim unnecessary trailing zeros from float repr.
        let s = format!("{}", n);
        s
    }
}

fn format_integer(n: f64) -> String {
    format!("{}", n.round() as i64)
}

fn format_fixed(n: f64, decimals: u8) -> String {
    format!("{:.prec$}", n, prec = decimals as usize)
}

/// Format a number with thousands-separator commas and the given decimal places.
pub fn format_commas(n: f64, decimals: u8) -> String {
    let negative = n < 0.0;
    let abs = n.abs();
    let sign = if negative { "-" } else { "" };

    let factor = 10f64.powi(decimals as i32);
    let scaled = (abs * factor).round();

    // Fall back to the locale-free Rust formatter for magnitudes that
    // overflow u64 — better to lose the thousands separators than to
    // emit a silently-wrapped integer.
    if !scaled.is_finite() || scaled >= u64::MAX as f64 {
        return format!("{}{:.prec$}", sign, abs, prec = decimals as usize);
    }

    let scaled_int = scaled as u64;

    if decimals == 0 {
        format!("{}{}", sign, insert_commas(scaled_int))
    } else {
        let divisor = factor as u64;
        let int_part = scaled_int / divisor;
        let frac = scaled_int % divisor;
        format!(
            "{}{}.{:0>width$}",
            sign,
            insert_commas(int_part),
            frac,
            width = decimals as usize
        )
    }
}

fn format_currency(n: f64, symbol: &str, decimals: u8) -> String {
    // Put any minus sign before the currency symbol so callers see
    // "-$99.50" rather than "$-99.50".
    if n < 0.0 {
        format!("-{}{}", symbol, format_commas(n.abs(), decimals))
    } else {
        format!("{}{}", symbol, format_commas(n, decimals))
    }
}

/// Format a number as a percentage (multiplied by 100, with optional decimal places).
pub fn format_percent(n: f64, decimals: u8) -> String {
    let pct = n * 100.0;
    if decimals == 0 {
        format!("{}%", pct.round() as i64)
    } else {
        format!("{:.prec$}%", pct, prec = decimals as usize)
    }
}

fn format_scientific(n: f64) -> String {
    // Excel uses E+XX notation (no leading zero in exponent on some locales, but
    // two-digit exponent is safest for matching).
    format!("{:.2E}", n)
}

fn insert_commas(n: u64) -> String {
    let s = n.to_string();
    let bytes = s.as_bytes();
    let len = bytes.len();
    let mut out = String::with_capacity(len + len / 3);
    for (i, &b) in bytes.iter().enumerate() {
        if i > 0 && (len - i).is_multiple_of(3) {
            out.push(',');
        }
        out.push(b as char);
    }
    out
}

// ── Custom format string interpreter ──────────────────────────────────────

/// Simplified parser for Excel format strings. Handles the common cases:
/// thousands separators, decimal places, percentages, currency symbols,
/// and scientific notation. Strips color/condition brackets and literals.
fn apply_custom(n: f64, fmt: &str) -> String {
    // Sections are positive;negative;zero;text. Pick the one that applies
    // to this value: a format like `#,##0;[Red](#,##0)` renders -1234 with
    // the *second* section, and a two-section `"yes";"no"` is how a boolean
    // flag column is written.
    let sections = split_format_sections(fmt);

    // A format whose sections carry `[<op><number>]` conditions is a
    // *conditional* format, not positive;negative;zero: sections are tested
    // in order and the first match wins, with an unconditioned section as
    // the fallback (ECMA-376 §18.8.31). `[>999999]#,,"M";[>999]#,"K";#` is
    // how a sheet renders 1.02 as `1` and 102102 as `102K`; reading it
    // positionally scaled every value by 1e6 and produced `0M`.
    let conditional = sections.iter().any(|s| section_condition(s).is_some());
    let (section, use_magnitude) = if conditional {
        let chosen = sections
            .iter()
            .find(|s| match section_condition(s) {
                Some((op, v)) => condition_holds(n, op, v),
                None => true,
            })
            .copied()
            .unwrap_or(fmt);
        (chosen, false)
    } else {
        let chosen = match sections.len() {
            0 => fmt,
            1 => sections[0],
            2 => {
                if n < 0.0 {
                    sections[1]
                } else {
                    sections[0]
                }
            },
            _ => {
                if n < 0.0 {
                    sections[1]
                } else if n == 0.0 {
                    sections[2]
                } else {
                    sections[0]
                }
            },
        };
        // The negative section supplies its own sign (usually parentheses or
        // a literal '-'), so format its magnitude.
        (chosen, sections.len() >= 2 && n < 0.0)
    };
    let n = if use_magnitude { n.abs() } else { n };

    // ── Parse the section ────────────────────────────────────────────────
    let mut currency_prefix = String::new();
    let mut suffix = String::new(); // literal text after the number
    let mut has_percent = false;
    let mut has_comma_in_num = false;
    let mut decimal_zeros = 0u8; // '0' chars after '.'
    let mut _decimal_hashes = 0u8; // '#' chars after '.'  (optional digits)
    let mut has_scientific = false;
    let mut in_decimal = false;
    let mut in_num_part = false;
    // Divisor accumulated from commas trailing the digit placeholders.
    let mut scale_divisor = 1.0f64;
    // Literal text collected before any digit placeholder appears.
    let mut prefix_literal = String::new();

    let mut chars = section.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            // Bracketed: colour like [Red] or locale/currency like [$€-407]
            '[' => {
                let mut inner = String::new();
                for ch in chars.by_ref() {
                    if ch == ']' {
                        break;
                    }
                    inner.push(ch);
                }
                if let Some(rest) = inner.strip_prefix('$') {
                    // [$symbol-locale] — extract symbol
                    let sym: String = rest.chars().take_while(|&ch| ch != '-').collect();
                    if !sym.is_empty() {
                        currency_prefix = sym;
                    }
                }
                // Colour directives ignored.
            },
            // Quoted literal text. Before any digit placeholder it is a
            // prefix; after one it is a suffix.
            '"' => {
                for ch in chars.by_ref() {
                    if ch == '"' {
                        break;
                    }
                    if in_num_part {
                        suffix.push(ch);
                    } else {
                        prefix_literal.push(ch);
                    }
                }
            },
            // Escape: next char is literal
            '\\' => {
                if let Some(ch) = chars.next() {
                    if in_num_part {
                        suffix.push(ch);
                    } else {
                        prefix_literal.push(ch);
                    }
                }
            },
            // _X = pad with X (alignment) — skip X
            '_' => {
                chars.next();
            },
            // *X = repeat X (fill) — skip X
            '*' => {
                chars.next();
            },

            '%' => {
                has_percent = true;
                in_num_part = true;
            },
            '.' => {
                in_decimal = true;
                in_num_part = true;
            },
            '0' => {
                in_num_part = true;
                if in_decimal {
                    decimal_zeros += 1;
                }
            },
            '#' => {
                in_num_part = true;
                if in_decimal {
                    _decimal_hashes += 1;
                }
            },
            ',' => {
                // A comma *between* digit placeholders is the thousands
                // separator; a comma *after* the last placeholder scales the
                // value down by 1000 each. `#,##0,," M"` is how a
                // finance sheet renders 12,500,000 as "12 M" — treating the
                // trailing commas as separators printed the full number.
                if in_num_part {
                    if chars
                        .peek()
                        .is_some_and(|c| matches!(c, '0' | '#' | '?' | '.'))
                    {
                        has_comma_in_num = true;
                    } else {
                        scale_divisor *= 1000.0;
                    }
                }
            },
            'E' | 'e' => {
                // Only treat this as scientific notation when followed by
                // `+` or `-` (per ECMA-376 §18.8.31). Bare `E` is just
                // a literal in formats like "000E" and must not consume
                // the next character.
                if matches!(chars.peek(), Some('+') | Some('-')) {
                    has_scientific = true;
                    chars.next(); // consume the sign
                    while chars.peek().is_some_and(|c| c.is_ascii_digit()) {
                        chars.next();
                    }
                } else if !in_num_part {
                    currency_prefix.push(c);
                } else {
                    suffix.push(c);
                }
            },
            '$' => {
                currency_prefix = "$".to_string();
                in_num_part = true;
            },
            // Other literal characters before the number part = currency prefix
            c if !in_num_part && !c.is_ascii_whitespace() => {
                currency_prefix.push(c);
            },
            _ => {},
        }
    }

    let decimals = decimal_zeros; // treat '0' decimals as the required precision

    // ── Format the value ─────────────────────────────────────────────────
    // A section with no digit placeholder at all is pure literal text —
    // `"yes";"no"` names two strings, not two numbers. Emitting the number
    // alongside them produced "1yes".
    //
    // Unquoted literal characters accumulate in `currency_prefix`, so they
    // must be included: dropping them turned an unrecognised format code
    // into an empty cell, losing the value entirely.
    //
    // A date/time code (`h"时"mm"分"ss"秒"`) also has no digit placeholder,
    // but it denotes a *value*, not a literal — echoing its letters back
    // prints the format code where the data should be. Such a cell should
    // have been rendered by the date path; if it reaches here, fall back to
    // the plain number rather than inventing text.
    if !in_num_part {
        if super::date::is_date_format_string(section) {
            return format_general(n);
        }
        return format!("{currency_prefix}{prefix_literal}{suffix}");
    }

    let value = if has_percent { n * 100.0 } else { n } / scale_divisor;

    let body = if has_scientific {
        format_scientific(value)
    } else if has_comma_in_num {
        format_commas(value, decimals)
    } else if in_decimal && decimals > 0 {
        format_fixed(value, decimals)
    } else if in_num_part {
        format_integer(value)
    } else {
        format_general(value)
    };

    let pct_suffix = if has_percent { "%" } else { "" };

    format!("{currency_prefix}{prefix_literal}{body}{suffix}{pct_suffix}")
}

/// Strip leading `[...]` directives — `[DBNum1]`, `[$-804]`, `[Red]` — from
/// a format code, leaving the code proper.
fn strip_leading_directives(fmt: &str) -> &str {
    let mut rest = fmt.trim_start();
    while let Some(inner) = rest.strip_prefix('[') {
        match inner.find(']') {
            Some(i) => rest = inner[i + 1..].trim_start(),
            None => break,
        }
    }
    rest
}

/// Read a leading `[<op><number>]` comparison condition off a section.
///
/// Only comparisons count: `[Red]` is a colour and `[$-409]` a locale, and
/// neither selects a section.
fn section_condition(section: &str) -> Option<(&'static str, f64)> {
    let rest = section.trim_start().strip_prefix('[')?;
    let end = rest.find(']')?;
    let inner = &rest[..end];
    for op in ["<=", ">=", "<>", "<", ">", "="] {
        if let Some(num) = inner.strip_prefix(op) {
            if let Ok(v) = num.trim().parse::<f64>() {
                // `op` is one of the literals above, so it outlives the call.
                let op: &'static str = match op {
                    "<=" => "<=",
                    ">=" => ">=",
                    "<>" => "<>",
                    "<" => "<",
                    ">" => ">",
                    _ => "=",
                };
                return Some((op, v));
            }
        }
    }
    None
}

fn condition_holds(n: f64, op: &str, v: f64) -> bool {
    match op {
        "<" => n < v,
        ">" => n > v,
        "<=" => n <= v,
        ">=" => n >= v,
        "<>" => n != v,
        _ => n == v,
    }
}

/// Split a number-format code into its `;`-separated sections, ignoring
/// semicolons inside quoted literals, `\`-escapes and `[...]` directives.
fn split_format_sections(fmt: &str) -> Vec<&str> {
    let mut out = Vec::new();
    let mut start = 0usize;
    let mut in_quotes = false;
    let mut in_bracket = false;
    let mut escaped = false;
    for (i, c) in fmt.char_indices() {
        if escaped {
            escaped = false;
            continue;
        }
        match c {
            '\\' => escaped = true,
            '"' if !in_bracket => in_quotes = !in_quotes,
            '[' if !in_quotes => in_bracket = true,
            ']' if in_bracket => in_bracket = false,
            ';' if !in_quotes && !in_bracket => {
                out.push(&fmt[start..i]);
                start = i + 1;
            },
            _ => {},
        }
    }
    out.push(&fmt[start..]);
    out
}

// ── Tests ──────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    /// `[>999999]#,,"M";[>999]#,"K";#` is how a sheet renders large numbers
    /// compactly. The sections carry *conditions*, not positive/negative
    /// roles: reading them positionally scaled every value by 1e6 and
    /// rendered 1.02 as `0M`.
    #[test]
    fn conditional_sections_select_by_comparison_not_by_position() {
        let fmt = Some(r#"[>999999]#,,"M";[>999]#,"K";#"#);
        assert_eq!(apply_format(1.02, 164, fmt), "1");
        assert_eq!(apply_format(102.0, 164, fmt), "102");
        assert_eq!(apply_format(1021.02, 164, fmt), "1K");
        assert_eq!(apply_format(102102.102, 164, fmt), "102K");
        assert_eq!(apply_format(1_500_000.0, 164, fmt), "2M");
    }

    #[test]
    fn a_colour_or_locale_directive_is_not_a_condition() {
        // `[Red]` and `[$-409]` must leave the positional
        // positive;negative;zero reading intact.
        let fmt = Some(r#"#,##0;[Red]-#,##0"#);
        assert_eq!(apply_format(1234.0, 164, fmt), "1,234");
        assert_eq!(apply_format(-1234.0, 164, fmt), "-1,234");
    }

    /// A custom `numFmt` whose code is `GENERAL` is the General format, not
    /// a literal. Matching it case-sensitively sent it down the custom path,
    /// where it produced no digit placeholder and rendered every numeric
    /// cell in the sheet as an empty string.
    #[test]
    fn a_general_format_code_is_not_a_literal() {
        for code in ["GENERAL", "General", "general"] {
            assert_eq!(apply_format(70.0, 164, Some(code)), "70");
            assert_eq!(apply_format(3.5, 164, Some(code)), "3.5");
        }
    }

    /// `[DBNum1][$-804]General` is the General format behind two directives.
    /// Treating the whole code as a literal printed `General` and dropped
    /// the cell's value.
    #[test]
    fn general_behind_bracket_directives_is_still_general() {
        assert_eq!(apply_format(12323.0, 180, Some("[DBNum1][$-804]General")), "12323");
        assert_eq!(apply_format(1.5, 180, Some("[$-409]General")), "1.5");
    }

    /// A date/time code has no digit placeholder either, but it denotes a
    /// value rather than a literal: echoing its letters printed the format
    /// code where the data should be.
    #[test]
    fn a_time_format_does_not_echo_its_own_code() {
        let out = apply_format(0.5555671296296296, 179, Some(r#"h"时"mm"分"ss"秒";@"#));
        assert!(
            !out.contains('h') && !out.contains('时'),
            "the format code leaked into the value: {out}"
        );
        assert!(out.starts_with("0.55"), "expected the numeric value, got {out}");
    }

    #[test]
    fn an_unrecognised_literal_format_keeps_its_text() {
        // Unquoted literal characters accumulate separately from quoted
        // ones; dropping them turned the cell into an empty string.
        assert_eq!(apply_format(1.0, 164, Some("ABC")), "ABC");
    }

    #[test]
    fn builtin_general() {
        assert_eq!(apply_format(42.0, 0, None), "42");
        assert_eq!(apply_format(4.25, 0, None), "4.25");
    }

    #[test]
    fn builtin_format_code_lookup() {
        assert_eq!(builtin_format_code(0), Some("General"));
        assert_eq!(builtin_format_code(4), Some("#,##0.00"));
        assert_eq!(builtin_format_code(14), Some("m/d/yyyy"));
        assert_eq!(builtin_format_code(49), Some("@"));
        // Custom-format ID range has no built-in code.
        assert_eq!(builtin_format_code(164), None);
        assert_eq!(builtin_format_code(200), None);
    }

    #[test]
    fn builtin_integer() {
        assert_eq!(apply_format(42.7, 1, None), "43");
    }

    #[test]
    fn builtin_fixed_two() {
        assert_eq!(apply_format(4.25678, 2, None), "4.26");
    }

    #[test]
    fn builtin_commas_zero() {
        assert_eq!(apply_format(1234567.0, 3, None), "1,234,567");
    }

    #[test]
    fn builtin_commas_two() {
        assert_eq!(apply_format(1234567.891, 4, None), "1,234,567.89");
    }

    #[test]
    fn builtin_percent_zero() {
        assert_eq!(apply_format(0.75, 9, None), "75%");
    }

    #[test]
    fn builtin_percent_two() {
        assert_eq!(apply_format(0.1234, 10, None), "12.34%");
    }

    #[test]
    fn builtin_currency_usd() {
        assert_eq!(apply_format(1234.5, 7, None), "$1,234.50");
    }

    #[test]
    fn custom_thousands() {
        assert_eq!(apply_format(1234567.0, 164, Some("#,##0")), "1,234,567");
    }

    #[test]
    fn custom_thousands_two_decimals() {
        assert_eq!(apply_format(1234.5, 164, Some("#,##0.00")), "1,234.50");
    }

    #[test]
    fn custom_percent() {
        assert_eq!(apply_format(0.5, 164, Some("0%")), "50%");
    }

    #[test]
    fn custom_percent_decimals() {
        assert_eq!(apply_format(0.1256, 164, Some("0.00%")), "12.56%");
    }

    #[test]
    fn custom_euro() {
        let result = apply_format(1234.5, 164, Some("[$€-407]#,##0.00"));
        assert!(result.contains("€"), "expected euro symbol, got: {result}");
        assert!(result.contains("1,234.50"), "expected formatted number, got: {result}");
    }

    #[test]
    fn custom_dollar_prefix() {
        assert_eq!(apply_format(99.9, 164, Some("$#,##0.00")), "$99.90");
    }

    #[test]
    fn negative_commas() {
        assert_eq!(apply_format(-1234.5, 4, None), "-1,234.50");
    }

    #[test]
    fn zero_percent() {
        assert_eq!(apply_format(0.0, 9, None), "0%");
    }

    #[test]
    fn large_commas() {
        assert_eq!(apply_format(1_000_000_000.0, 3, None), "1,000,000,000");
    }

    // ── Edge cases ──────────────────────────────────────────────────────

    #[test]
    fn nan_renders_as_label() {
        // Returning the literal "NaN" rather than an empty string keeps
        // anomalous cells visible in extracted text so they're not
        // mistaken for empty data.
        assert_eq!(apply_format(f64::NAN, 0, None), "NaN");
    }

    #[test]
    fn infinity_renders_as_label() {
        assert_eq!(apply_format(f64::INFINITY, 0, None), "Infinity");
        assert_eq!(apply_format(f64::NEG_INFINITY, 0, None), "-Infinity");
    }

    #[test]
    fn zero_renders_uniformly() {
        assert_eq!(apply_format(0.0, 0, None), "0");
        assert_eq!(apply_format(0.0, 2, None), "0.00");
        assert_eq!(apply_format(0.0, 4, None), "0.00");
    }

    #[test]
    fn negative_percent() {
        assert_eq!(apply_format(-0.25, 9, None), "-25%");
        assert_eq!(apply_format(-0.1234, 10, None), "-12.34%");
    }

    #[test]
    fn negative_currency() {
        assert_eq!(apply_format(-99.5, 7, None), "-$99.50");
    }

    #[test]
    fn scientific_builtin() {
        // Format id 11 = 0.00E+00 → uses Rust's "{:.2E}" wrapper.
        let s = apply_format(12345.6789, 11, None);
        assert!(s.contains('E'), "scientific got: {s}");
    }

    #[test]
    fn accounting_alias() {
        // 37–40 map to comma formats matching #,##0 family.
        assert_eq!(apply_format(1234.0, 37, None), "1,234");
        assert_eq!(apply_format(1234.5, 39, None), "1,234.50");
    }

    #[test]
    fn accounting_paren_range() {
        // 41..=44 are accounting variants → commas with 2 decimals.
        for id in 41u32..=44 {
            assert_eq!(apply_format(1234.5, id, None), "1,234.50", "fmt id {id}");
        }
    }

    #[test]
    fn fraction_falls_back_to_general() {
        // Fraction formats (12,13) currently render as general.
        assert_eq!(apply_format(1.5, 12, None), "1.5");
        assert_eq!(apply_format(2.0, 13, None), "2");
    }

    #[test]
    fn custom_general_falls_through_to_default() {
        // "General" and "@" should fall back to General formatting.
        assert_eq!(apply_format(42.5, 164, Some("General")), "42.5");
        assert_eq!(apply_format(42.0, 164, Some("@")), "42");
    }

    #[test]
    fn custom_blank_falls_back_to_general() {
        assert_eq!(apply_format(4.25, 164, Some("")), "4.25");
        assert_eq!(apply_format(4.25, 164, Some("   ")), "4.25");
    }

    #[test]
    fn custom_multi_section_uses_first() {
        // Multi-section format: positives use first section only.
        assert_eq!(apply_format(1234.5, 164, Some("#,##0.00;-#,##0.00")), "1,234.50");
    }

    #[test]
    fn custom_with_quoted_literal_suffix() {
        let result = apply_format(42.0, 164, Some(r#"0" units""#));
        assert!(result.contains("42"), "got: {result}");
        assert!(result.contains("units"), "got: {result}");
    }

    #[test]
    fn custom_color_directive_is_stripped() {
        // [Red] is a color directive — should be ignored, not emitted.
        let result = apply_format(123.0, 164, Some("[Red]#,##0"));
        assert!(!result.contains("Red"));
        assert!(result.contains("123"));
    }

    #[test]
    fn format_general_keeps_integers_unsuffixed() {
        // Whole-number floats render without ".0".
        assert_eq!(format_general(42.0), "42");
        assert_eq!(format_general(-7.0), "-7");
        assert_eq!(format_general(0.0), "0");
    }

    #[test]
    fn format_general_keeps_decimal_for_fraction() {
        assert_eq!(format_general(4.25), "4.25");
        assert_eq!(format_general(-2.5), "-2.5");
    }

    #[test]
    fn format_commas_negative_with_decimals() {
        assert_eq!(format_commas(-1234.5, 2), "-1,234.50");
    }

    #[test]
    fn format_commas_zero() {
        assert_eq!(format_commas(0.0, 0), "0");
        assert_eq!(format_commas(0.0, 2), "0.00");
    }

    #[test]
    fn format_percent_negative() {
        assert_eq!(format_percent(-0.5, 0), "-50%");
    }

    #[test]
    fn format_percent_zero_decimals() {
        // 50% with 0 decimals.
        assert_eq!(format_percent(0.5, 0), "50%");
    }
}
