//! Excel number format rendering.
//!
//! Applies a numeric format string (or built-in format ID) to an f64 value
//! and returns the display string. Covers the cases that matter in practice:
//! integers, fixed decimals, thousands separators, percentages, currency,
//! and scientific notation. Complex conditions/colors are stripped gracefully.

/// The built-in id whose format code is exactly `code`, if any.
///
/// Writing an explicit `<numFmt>` for a code that is already built in gives
/// the cell a custom id (164+) for no reason, which shows up as an IR
/// difference on the second write→parse cycle.
pub fn builtin_id_for_code(code: &str) -> Option<u32> {
    (0..=49).find(|&id| builtin_format_code(id) == Some(code))
}

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
    if let Some(s) = apply_builtin(n, fmt_id) {
        return s;
    }
    // Custom format string (IDs 164+), or a workbook's explicit override of a
    // built-in id. Callers must pass the *declared* code here, never one
    // resolved from the built-in table: apply_custom is not a general format
    // engine and garbles codes the match in `apply_builtin` already declined
    // to handle.
    match fmt_str.and_then(compile_custom) {
        Some(compiled) => compiled.render(n),
        None => format_general(n),
    }
}

/// [`apply_format`] with the declared custom code already compiled by
/// [`compile_custom`] — what a renderer formatting many cells under one
/// `<numFmt>` uses, so the code is parsed once rather than once per cell.
pub(crate) fn apply_format_compiled(
    n: f64,
    fmt_id: u32,
    compiled: Option<&CompiledFormat>,
) -> String {
    if let Some(s) = apply_builtin(n, fmt_id) {
        return s;
    }
    match compiled {
        Some(c) => c.render(n),
        None => format_general(n),
    }
}

/// Non-finite values, and the built-in format ids rendered without their
/// code; `None` when `fmt_id` needs its declared code.
fn apply_builtin(n: f64, fmt_id: u32) -> Option<String> {
    if n.is_nan() {
        return Some("NaN".to_string());
    }
    if n.is_infinite() {
        return Some(if n < 0.0 {
            "-Infinity".to_string()
        } else {
            "Infinity".to_string()
        });
    }

    // Built-in format IDs per OOXML spec §18.8.30.
    Some(match fmt_id {
        0 | 49 => format_general(n),         // General / @
        1 => format_integer(n),              // 0
        2 => format_fixed(n, 2),             // 0.00
        3 => format_commas(n, 0),            // #,##0
        4 => format_commas(n, 2),            // #,##0.00
        5 | 6 => format_currency(n, "$", 0), // $#,##0
        7 | 8 => format_currency(n, "$", 2), // $#,##0.00
        9 => format_percent(n, 0),           // 0%
        10 => format_percent(n, 2),          // 0.00%
        11 => format_scientific(n),          // 0.00E+00
        12 => format_general(n),             // # ?/? (fractions — approx)
        13 => format_general(n),             // # ??/??
        // Accounting/comma built-ins wrap negatives in parentheses rather
        // than using a leading minus — the codes this file's own
        // `builtin_format_code` declares for them (`#,##0 ;(#,##0)`) say so,
        // per ECMA-376 §18.8.30.
        37 | 38 => format_accounting(n, 0), // #,##0 accounting variants
        39 | 40 => format_accounting(n, 2), // #,##0.00 accounting variants
        41..=44 => format_accounting(n, 2), // _(* ...) accounting
        _ => return None,
    })
}

/// Compile a declared format code, or `None` when it is General/text and
/// renders as General. The General/text sentinels are matched
/// case-insensitively, and after leading `[...]` directives are stripped:
/// real workbooks declare `numFmt formatCode="GENERAL"` and
/// `"[DBNum1][$-804]General"`, and treating either as a literal format code
/// dropped the cell's value and printed the code.
pub(crate) fn compile_custom(fmt: &str) -> Option<CompiledFormat> {
    let fmt = fmt.trim();
    let bare = strip_leading_directives(fmt);
    if fmt.is_empty() || bare.eq_ignore_ascii_case("General") || fmt == "@" {
        return None;
    }
    Some(CompiledFormat::compile(fmt))
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

    // One allocation per number: this runs once per formatted cell, and
    // the `to_string` + `format!` pair it replaced was a measurable share
    // of a large workbook's text extraction.
    let mut out = String::with_capacity(28);
    out.push_str(sign);
    if decimals == 0 {
        push_commas(&mut out, scaled_int);
    } else {
        let divisor = factor as u64;
        push_commas(&mut out, scaled_int / divisor);
        out.push('.');
        // Zero-padded to `decimals` digits, written in place.
        use std::fmt::Write;
        let frac = scaled_int % divisor;
        write!(out, "{frac:0width$}", width = decimals as usize).ok();
    }
    out
}

/// Format a number the way Excel's accounting/comma built-ins (ids 37-44)
/// do: negatives are parenthesised rather than signed with a leading minus.
pub fn format_accounting(n: f64, decimals: u8) -> String {
    if n < 0.0 {
        format!("({})", format_commas(n.abs(), decimals))
    } else {
        format_commas(n, decimals)
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
    // Excel's `0.00E+00` always signs the exponent and pads it to two
    // digits. Rust's `{:E}` does neither (`1.23E4`, `1.23E-3`), so the
    // exponent is reassembled by hand.
    let s = format!("{:.2E}", n);
    let Some((mantissa, exp)) = s.split_once('E') else {
        return s;
    };
    let (sign, digits) = match exp.strip_prefix('-') {
        Some(d) => ('-', d),
        None => ('+', exp.strip_prefix('+').unwrap_or(exp)),
    };
    format!("{mantissa}E{sign}{digits:0>2}")
}

/// Append `n` with thousands separators, no intermediate allocation.
fn push_commas(out: &mut String, mut n: u64) {
    let mut digits = [0u8; 20];
    let mut len = 0;
    loop {
        digits[len] = b'0' + (n % 10) as u8;
        len += 1;
        n /= 10;
        if n == 0 {
            break;
        }
    }
    for i in (0..len).rev() {
        out.push(digits[i] as char);
        if i > 0 && i.is_multiple_of(3) {
            out.push(',');
        }
    }
}

// ── Custom format string interpreter ──────────────────────────────────────

/// Simplified parser for Excel format strings. Handles the common cases:
/// thousands separators, decimal places, percentages, currency symbols,
/// and scientific notation. Strips color/condition brackets and literals.
#[cfg(test)]
fn apply_custom(n: f64, fmt: &str) -> String {
    CompiledFormat::compile(fmt).render(n)
}

/// A custom format code parsed once: its sections, each with the
/// condition it carries and the placeholders/literals it is made of.
/// Rendering many cells under one `<numFmt>` re-split and re-parsed the
/// code for every cell.
#[derive(Debug, Clone)]
pub(crate) struct CompiledFormat {
    /// `[<op><number>]`-conditioned sections are tested in order.
    conditional: bool,
    /// The sections (positive;negative;zero;text) and their conditions.
    sections: Vec<(Option<(&'static str, f64)>, CompiledSection)>,
    /// The whole code as one section: used when it has no sections, or a
    /// conditional code has none that applies.
    whole: CompiledSection,
}

impl CompiledFormat {
    fn compile(fmt: &str) -> Self {
        let sections: Vec<_> = split_format_sections(fmt)
            .into_iter()
            .map(|s| (section_condition(s), CompiledSection::compile(s)))
            .collect();
        Self {
            conditional: sections.iter().any(|(c, _)| c.is_some()),
            sections,
            whole: CompiledSection::compile(fmt),
        }
    }

    pub(crate) fn render(&self, n: f64) -> String {
        // Sections are positive;negative;zero;text. Pick the one that
        // applies to this value: a format like `#,##0;[Red](#,##0)` renders
        // -1234 with the *second* section, and a two-section `"yes";"no"`
        // is how a boolean flag column is written.
        //
        // A format whose sections carry `[<op><number>]` conditions is a
        // *conditional* format, not positive;negative;zero: sections are
        // tested in order and the first match wins, with an unconditioned
        // section as the fallback (ECMA-376 §18.8.31).
        // `[>999999]#,,"M";[>999]#,"K";#` is how a sheet renders 1.02 as
        // `1` and 102102 as `102K`; reading it positionally scaled every
        // value by 1e6 and produced `0M`.
        let (section, use_magnitude) = if self.conditional {
            let chosen = self
                .sections
                .iter()
                .find(|(cond, _)| match cond {
                    Some((op, v)) => condition_holds(n, op, *v),
                    None => true,
                })
                .map_or(&self.whole, |(_, s)| s);
            (chosen, false)
        } else {
            let s = &self.sections;
            let chosen = match s.len() {
                0 => &self.whole,
                1 => &s[0].1,
                2 => {
                    if n < 0.0 {
                        &s[1].1
                    } else {
                        &s[0].1
                    }
                },
                _ => {
                    if n < 0.0 {
                        &s[1].1
                    } else if n == 0.0 {
                        &s[2].1
                    } else {
                        &s[0].1
                    }
                },
            };
            // The negative section supplies its own sign (usually
            // parentheses or a literal '-'), so format its magnitude.
            (chosen, s.len() >= 2 && n < 0.0)
        };
        let n = if use_magnitude { n.abs() } else { n };
        section.render(n)
    }
}

/// One `;`-separated section of a custom format code, parsed.
#[derive(Debug, Clone)]
struct CompiledSection {
    currency_prefix: String,
    prefix_literal: String,
    suffix: String,
    has_percent: bool,
    has_comma_in_num: bool,
    decimals: u8,
    has_scientific: bool,
    in_decimal: bool,
    in_num_part: bool,
    scale_divisor: f64,
    has_forced_integer_digit: bool,
    /// For a section with no digit placeholder: whether it is a date/time
    /// code (rendered as the plain number) rather than literal text.
    literal_is_date: bool,
}

impl CompiledSection {
    fn compile(section: &str) -> Self {
        // ── Parse the section ────────────────────────────────────────────────
        let mut currency_prefix = String::new();
        let mut suffix = String::new(); // literal text after the number
        let mut has_percent = false;
        let mut has_comma_in_num = false;
        let mut decimal_zeros = 0u8; // '0' chars after '.'
        let mut has_scientific = false;
        let mut in_decimal = false;
        let mut in_num_part = false;
        // Divisor accumulated from commas trailing the digit placeholders.
        let mut scale_divisor = 1.0f64;
        // Whether the integer part contains a `0` placeholder, which forces a
        // digit to be shown even when the value rounds to zero.
        let mut has_forced_integer_digit = false;
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
                // `0` forces a digit; `#` and `?` are *optional* digit
                // placeholders (`?` pads with a space instead of nothing).
                '0' => {
                    in_num_part = true;
                    if in_decimal {
                        decimal_zeros += 1;
                    } else {
                        has_forced_integer_digit = true;
                    }
                },
                '#' | '?' => {
                    in_num_part = true;
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
                // ...and after it, a suffix. Dropping these lost the closing
                // paren of a custom `#,##0;(#,##0)` and any bare trailing
                // literal, so the rendered string didn't match the format.
                c => {
                    if in_num_part {
                        suffix.push(c);
                    }
                },
            }
        }

        let literal_is_date = !in_num_part && super::date::is_date_format_string(section);
        Self {
            currency_prefix,
            prefix_literal,
            suffix,
            has_percent,
            has_comma_in_num,
            decimals: decimal_zeros,
            has_scientific,
            in_decimal,
            in_num_part,
            scale_divisor,
            has_forced_integer_digit,
            literal_is_date,
        }
    }

    fn render(&self, n: f64) -> String {
        let Self {
            currency_prefix,
            prefix_literal,
            suffix,
            has_percent,
            has_comma_in_num,
            decimals,
            has_scientific,
            in_decimal,
            in_num_part,
            scale_divisor,
            has_forced_integer_digit,
            literal_is_date,
        } = self;
        let (has_percent, has_comma_in_num, decimals, has_scientific, in_decimal, in_num_part) = (
            *has_percent,
            *has_comma_in_num,
            *decimals,
            *has_scientific,
            *in_decimal,
            *in_num_part,
        );
        let (scale_divisor, has_forced_integer_digit) = (*scale_divisor, *has_forced_integer_digit);

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
            if *literal_is_date {
                return format_general(n);
            }
            return format!("{currency_prefix}{prefix_literal}{suffix}");
        }

        let value = if has_percent { n * 100.0 } else { n } / scale_divisor;

        // A value that rounds to zero renders as *nothing* when the integer part
        // has only optional placeholders (`#`/`?`) — which is exactly how the
        // accounting formats' zero section, `_-* "-"??_-`, shows a bare dash.
        // Forcing a digit there printed `0` where Excel prints nothing.
        let rounds_to_zero = decimals == 0 && value.round() == 0.0;
        let body = if has_scientific {
            format_scientific(value)
        } else if rounds_to_zero && !has_forced_integer_digit && in_num_part {
            String::new()
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

        if currency_prefix.is_empty() && prefix_literal.is_empty() && suffix.is_empty() {
            let mut body = body;
            body.push_str(pct_suffix);
            return body;
        }
        let mut out = String::with_capacity(
            currency_prefix.len() + prefix_literal.len() + body.len() + suffix.len() + 1,
        );
        for part in [
            currency_prefix.as_str(),
            prefix_literal,
            &body,
            suffix,
            pct_suffix,
        ] {
            out.push_str(part);
        }
        out
    }
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
    fn test_conditional_sections_select_by_comparison_not_by_position() {
        let fmt = Some(r#"[>999999]#,,"M";[>999]#,"K";#"#);
        assert_eq!(apply_format(1.02, 164, fmt), "1");
        assert_eq!(apply_format(102.0, 164, fmt), "102");
        assert_eq!(apply_format(1021.02, 164, fmt), "1K");
        assert_eq!(apply_format(102102.102, 164, fmt), "102K");
        assert_eq!(apply_format(1_500_000.0, 164, fmt), "2M");
    }

    /// The accounting formats are the most common custom code in real
    /// spreadsheets, and their zero section is `_-* "-"??_-`: `?` is an
    /// *optional* digit placeholder, so a zero renders as a bare dash.
    /// Not handling `?` sent the section down the literal path and printed
    /// `??-` — 3,950 cells in one corpus file.
    #[test]
    fn test_the_accounting_zero_section_renders_a_bare_dash() {
        let fmt = Some(r#"_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-"#);
        let out = apply_format(0.0, 164, fmt);
        assert!(!out.contains('?'), "digit placeholders leaked into the value: {out}");
        assert!(out.contains('-'), "expected the dash literal: {out}");
    }

    /// `?` is a digit placeholder wherever it appears, not a literal.
    #[test]
    fn test_question_mark_is_a_digit_placeholder() {
        assert_eq!(apply_format(42.0, 164, Some("??")), "42");
        assert_eq!(apply_format(7.0, 164, Some("???")), "7");
    }

    /// A value that rounds to zero shows nothing when the integer part has
    /// only optional placeholders, and shows `0` when a `0` forces it.
    #[test]
    fn test_optional_placeholders_suppress_a_zero_that_a_forced_digit_keeps() {
        assert_eq!(apply_format(0.0, 164, Some("#")), "");
        assert_eq!(apply_format(0.0, 164, Some("0")), "0");
    }

    #[test]
    fn test_a_colour_or_locale_directive_is_not_a_condition() {
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
    fn test_a_general_format_code_is_not_a_literal() {
        for code in ["GENERAL", "General", "general"] {
            assert_eq!(apply_format(70.0, 164, Some(code)), "70");
            assert_eq!(apply_format(3.5, 164, Some(code)), "3.5");
        }
    }

    /// `[DBNum1][$-804]General` is the General format behind two directives.
    /// Treating the whole code as a literal printed `General` and dropped
    /// the cell's value.
    #[test]
    fn test_general_behind_bracket_directives_is_still_general() {
        assert_eq!(apply_format(12323.0, 180, Some("[DBNum1][$-804]General")), "12323");
        assert_eq!(apply_format(1.5, 180, Some("[$-409]General")), "1.5");
    }

    /// A date/time code has no digit placeholder either, but it denotes a
    /// value rather than a literal: echoing its letters printed the format
    /// code where the data should be.
    #[test]
    fn test_a_time_format_does_not_echo_its_own_code() {
        let out = apply_format(0.5555671296296296, 179, Some(r#"h"时"mm"分"ss"秒";@"#));
        assert!(
            !out.contains('h') && !out.contains('时'),
            "the format code leaked into the value: {out}"
        );
        assert!(out.starts_with("0.55"), "expected the numeric value, got {out}");
    }

    #[test]
    fn test_an_unrecognised_literal_format_keeps_its_text() {
        // Unquoted literal characters accumulate separately from quoted
        // ones; dropping them turned the cell into an empty string.
        assert_eq!(apply_format(1.0, 164, Some("ABC")), "ABC");
    }

    #[test]
    fn test_builtin_general() {
        assert_eq!(apply_format(42.0, 0, None), "42");
        assert_eq!(apply_format(4.25, 0, None), "4.25");
    }

    #[test]
    fn test_builtin_format_code_lookup() {
        assert_eq!(builtin_format_code(0), Some("General"));
        assert_eq!(builtin_format_code(4), Some("#,##0.00"));
        assert_eq!(builtin_format_code(14), Some("m/d/yyyy"));
        assert_eq!(builtin_format_code(49), Some("@"));
        // Custom-format ID range has no built-in code.
        assert_eq!(builtin_format_code(164), None);
        assert_eq!(builtin_format_code(200), None);
    }

    #[test]
    fn test_builtin_integer() {
        assert_eq!(apply_format(42.7, 1, None), "43");
    }

    #[test]
    fn test_builtin_fixed_two() {
        assert_eq!(apply_format(4.25678, 2, None), "4.26");
    }

    #[test]
    fn test_builtin_commas_zero() {
        assert_eq!(apply_format(1234567.0, 3, None), "1,234,567");
    }

    #[test]
    fn test_builtin_commas_two() {
        assert_eq!(apply_format(1234567.891, 4, None), "1,234,567.89");
    }

    #[test]
    fn test_builtin_percent_zero() {
        assert_eq!(apply_format(0.75, 9, None), "75%");
    }

    #[test]
    fn test_builtin_percent_two() {
        assert_eq!(apply_format(0.1234, 10, None), "12.34%");
    }

    #[test]
    fn test_builtin_currency_usd() {
        assert_eq!(apply_format(1234.5, 7, None), "$1,234.50");
    }

    #[test]
    fn test_custom_thousands() {
        assert_eq!(apply_format(1234567.0, 164, Some("#,##0")), "1,234,567");
    }

    #[test]
    fn test_custom_thousands_two_decimals() {
        assert_eq!(apply_format(1234.5, 164, Some("#,##0.00")), "1,234.50");
    }

    #[test]
    fn test_custom_percent() {
        assert_eq!(apply_format(0.5, 164, Some("0%")), "50%");
    }

    #[test]
    fn test_custom_percent_decimals() {
        assert_eq!(apply_format(0.1256, 164, Some("0.00%")), "12.56%");
    }

    #[test]
    fn test_custom_euro() {
        let result = apply_format(1234.5, 164, Some("[$€-407]#,##0.00"));
        assert!(result.contains("€"), "expected euro symbol, got: {result}");
        assert!(result.contains("1,234.50"), "expected formatted number, got: {result}");
    }

    #[test]
    fn test_custom_dollar_prefix() {
        assert_eq!(apply_format(99.9, 164, Some("$#,##0.00")), "$99.90");
    }

    #[test]
    fn test_negative_commas() {
        assert_eq!(apply_format(-1234.5, 4, None), "-1,234.50");
    }

    #[test]
    fn test_zero_percent() {
        assert_eq!(apply_format(0.0, 9, None), "0%");
    }

    #[test]
    fn test_large_commas() {
        assert_eq!(apply_format(1_000_000_000.0, 3, None), "1,000,000,000");
    }

    // ── Edge cases ──────────────────────────────────────────────────────

    #[test]
    fn test_nan_renders_as_label() {
        // Returning the literal "NaN" rather than an empty string keeps
        // anomalous cells visible in extracted text so they're not
        // mistaken for empty data.
        assert_eq!(apply_format(f64::NAN, 0, None), "NaN");
    }

    #[test]
    fn test_infinity_renders_as_label() {
        assert_eq!(apply_format(f64::INFINITY, 0, None), "Infinity");
        assert_eq!(apply_format(f64::NEG_INFINITY, 0, None), "-Infinity");
    }

    #[test]
    fn test_zero_renders_uniformly() {
        assert_eq!(apply_format(0.0, 0, None), "0");
        assert_eq!(apply_format(0.0, 2, None), "0.00");
        assert_eq!(apply_format(0.0, 4, None), "0.00");
    }

    #[test]
    fn test_negative_percent() {
        assert_eq!(apply_format(-0.25, 9, None), "-25%");
        assert_eq!(apply_format(-0.1234, 10, None), "-12.34%");
    }

    #[test]
    fn test_negative_currency() {
        assert_eq!(apply_format(-99.5, 7, None), "-$99.50");
    }

    #[test]
    fn test_scientific_builtin() {
        // Format id 11 = 0.00E+00 → uses Rust's "{:.2E}" wrapper.
        let s = apply_format(12345.6789, 11, None);
        assert!(s.contains('E'), "scientific got: {s}");
    }

    #[test]
    fn test_accounting_alias() {
        // 37–40 map to comma formats matching #,##0 family.
        assert_eq!(apply_format(1234.0, 37, None), "1,234");
        assert_eq!(apply_format(1234.5, 39, None), "1,234.50");
    }

    #[test]
    fn test_accounting_paren_range() {
        // 41..=44 are accounting variants → commas with 2 decimals.
        for id in 41u32..=44 {
            assert_eq!(apply_format(1234.5, id, None), "1,234.50", "fmt id {id}");
        }
    }

    #[test]
    fn test_fraction_falls_back_to_general() {
        // Fraction formats (12,13) currently render as general.
        assert_eq!(apply_format(1.5, 12, None), "1.5");
        assert_eq!(apply_format(2.0, 13, None), "2");
    }

    #[test]
    fn test_custom_general_falls_through_to_default() {
        // "General" and "@" should fall back to General formatting.
        assert_eq!(apply_format(42.5, 164, Some("General")), "42.5");
        assert_eq!(apply_format(42.0, 164, Some("@")), "42");
    }

    #[test]
    fn test_custom_blank_falls_back_to_general() {
        assert_eq!(apply_format(4.25, 164, Some("")), "4.25");
        assert_eq!(apply_format(4.25, 164, Some("   ")), "4.25");
    }

    #[test]
    fn test_custom_multi_section_uses_first() {
        // Multi-section format: positives use first section only.
        assert_eq!(apply_format(1234.5, 164, Some("#,##0.00;-#,##0.00")), "1,234.50");
    }

    #[test]
    fn test_custom_with_quoted_literal_suffix() {
        let result = apply_format(42.0, 164, Some(r#"0" units""#));
        assert!(result.contains("42"), "got: {result}");
        assert!(result.contains("units"), "got: {result}");
    }

    #[test]
    fn test_custom_color_directive_is_stripped() {
        // [Red] is a color directive — should be ignored, not emitted.
        let result = apply_format(123.0, 164, Some("[Red]#,##0"));
        assert!(!result.contains("Red"));
        assert!(result.contains("123"));
    }

    #[test]
    fn test_format_general_keeps_integers_unsuffixed() {
        // Whole-number floats render without ".0".
        assert_eq!(format_general(42.0), "42");
        assert_eq!(format_general(-7.0), "-7");
        assert_eq!(format_general(0.0), "0");
    }

    #[test]
    fn test_format_general_keeps_decimal_for_fraction() {
        assert_eq!(format_general(4.25), "4.25");
        assert_eq!(format_general(-2.5), "-2.5");
    }

    #[test]
    fn test_format_commas_negative_with_decimals() {
        assert_eq!(format_commas(-1234.5, 2), "-1,234.50");
    }

    #[test]
    fn test_format_commas_zero() {
        assert_eq!(format_commas(0.0, 0), "0");
        assert_eq!(format_commas(0.0, 2), "0.00");
    }

    #[test]
    fn test_format_percent_negative() {
        assert_eq!(format_percent(-0.5, 0), "-50%");
    }

    #[test]
    fn test_format_percent_zero_decimals() {
        // 50% with 0 decimals.
        assert_eq!(format_percent(0.5, 0), "50%");
    }

    /// Built-in accounting/comma ids 37-44 parenthesise negatives, as the
    /// format codes this file's own `builtin_format_code` declares for them
    /// specify (`#,##0 ;(#,##0)`). The fast path emitted a leading minus
    /// instead.
    #[test]
    fn test_accounting_builtins_parenthesize_negatives() {
        assert_eq!(apply_format(-1234.0, 37, None), "(1,234)");
        assert_eq!(apply_format(-1234.0, 38, None), "(1,234)");
        assert_eq!(apply_format(-1234.5, 39, None), "(1,234.50)");
        assert_eq!(apply_format(-1234.5, 40, None), "(1,234.50)");
        assert_eq!(apply_format(-1234.5, 41, None), "(1,234.50)");
        assert_eq!(apply_format(-1234.5, 44, None), "(1,234.50)");
        // Positives and zero are untouched.
        assert_eq!(apply_format(1234.0, 37, None), "1,234");
        assert_eq!(apply_format(0.0, 37, None), "0");
    }

    /// A literal character after the digit placeholders is part of the
    /// output. The interpreter dropped everything past the number part it
    /// did not recognise as a format token, losing the closing paren of an
    /// explicitly-declared `#,##0;(#,##0)`.
    #[test]
    fn test_custom_format_keeps_trailing_literal_characters() {
        assert_eq!(apply_format(-1234.0, 164, Some("#,##0;(#,##0)")), "(1,234)");
        assert_eq!(apply_format(42.0, 164, Some(r#"0"x")"#)), "42x)");
    }

    /// Excel's `0.00E+00` always signs the exponent and pads it to two
    /// digits; Rust's `{:E}` does neither.
    #[test]
    fn test_scientific_exponent_is_signed_and_padded() {
        assert_eq!(apply_format(12345.6789, 11, None), "1.23E+04");
        assert_eq!(apply_format(0.0012345, 11, None), "1.23E-03");
        assert_eq!(apply_format(1.5, 11, None), "1.50E+00");
        assert_eq!(apply_format(1.23e120, 11, None), "1.23E+120");
    }
}

#[cfg(test)]
mod builtin_code_tests {
    use super::*;

    /// `apply_format`'s `fmt_str` branch is for codes a workbook *declares*.
    /// Feeding it a code resolved from the built-in table sends it to
    /// `apply_custom`, which is not a general format engine: id 47's
    /// `mm:ss.0` came out as the literal `mm:ss0.6`.
    #[test]
    fn test_a_builtin_code_is_not_fed_back_in_as_a_custom_format() {
        let v = 0.563_138_888_888_888_9;
        // No declared override: the built-in table decides, and id 47 is not
        // one apply_format renders, so the raw value must survive.
        let out = apply_format(v, 47, None);
        assert!(!out.contains("mm:ss"), "format code leaked into the value: {out}");
        assert!(out.starts_with("0.56"), "value lost: {out}");

        // A genuine declared override is still honoured.
        let out = apply_format(0.5, 164, Some("0.00\" kg\""));
        assert!(out.contains("kg"), "declared custom format ignored: {out}");
    }
}
