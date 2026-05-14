//! Excel number format rendering.
//!
//! Applies a numeric format string (or built-in format ID) to an f64 value
//! and returns the display string. Covers the cases that matter in practice:
//! integers, fixed decimals, thousands separators, percentages, currency,
//! and scientific notation. Complex conditions/colors are stripped gracefully.

/// Apply an Excel number format to a numeric value.
pub fn apply_format(n: f64, fmt_id: u32, fmt_str: Option<&str>) -> String {
    if n.is_nan() || n.is_infinite() {
        return String::new();
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
        if !fmt.is_empty() && fmt != "General" && fmt != "@" {
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

    // Round to the required number of decimal places first.
    let factor = 10f64.powi(decimals as i32);
    let rounded = (abs * factor).round() / factor;

    let int_part = rounded.trunc() as u64;
    let int_str = insert_commas(int_part);

    let sign = if negative { "-" } else { "" };

    if decimals == 0 {
        format!("{}{}", sign, int_str)
    } else {
        let frac = ((rounded.fract()) * factor).round() as u64;
        format!("{}{}.{:0>width$}", sign, int_str, frac, width = decimals as usize)
    }
}

fn format_currency(n: f64, symbol: &str, decimals: u8) -> String {
    format!("{}{}", symbol, format_commas(n, decimals))
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
    // Multi-section: take the first section (positive numbers).
    // Second section = negatives, third = zero, fourth = text.
    let section = fmt.split(';').next().unwrap_or(fmt);

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
            // Quoted literal text — collect as suffix
            '"' => {
                for ch in chars.by_ref() {
                    if ch == '"' {
                        break;
                    }
                    suffix.push(ch);
                }
            },
            // Escape: next char is literal
            '\\' => {
                chars.next();
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
                // Comma between '#'/'0' chars = thousands separator.
                // Comma at end of number part = scale-by-1000 (rare, skip for now).
                if in_num_part {
                    has_comma_in_num = true;
                }
            },
            'E' | 'e' => {
                has_scientific = true;
                // Skip the +/- and exponent digits
                chars.next(); // '+' or '-'
                while chars.peek().is_some_and(|c| c.is_ascii_digit()) {
                    chars.next();
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
    let value = if has_percent { n * 100.0 } else { n };

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

    format!("{}{}{}{}", currency_prefix, body, suffix, pct_suffix)
}

// ── Tests ──────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn builtin_general() {
        assert_eq!(apply_format(42.0, 0, None), "42");
        assert_eq!(apply_format(4.25, 0, None), "4.25");
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
    fn nan_returns_empty() {
        assert_eq!(apply_format(f64::NAN, 0, None), "");
    }

    #[test]
    fn infinity_returns_empty() {
        assert_eq!(apply_format(f64::INFINITY, 0, None), "");
        assert_eq!(apply_format(f64::NEG_INFINITY, 0, None), "");
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
        assert_eq!(apply_format(-99.5, 7, None), "$-99.50");
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
