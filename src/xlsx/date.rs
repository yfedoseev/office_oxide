use std::fmt;

use super::styles::StyleSheet;

/// A date/time value converted from an Excel serial number.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DateTimeValue {
    /// Year component (e.g., 2024).
    pub year: i32,
    /// Month component (1–12).
    pub month: u32,
    /// Day component (1–31).
    pub day: u32,
    /// Hour component (0–23).
    pub hour: u32,
    /// Minute component (0–59).
    pub minute: u32,
    /// Second component (0–59).
    pub second: u32,
    /// Millisecond component (0–999).
    pub millisecond: u32,
}

impl fmt::Display for DateTimeValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if self.hour == 0 && self.minute == 0 && self.second == 0 && self.millisecond == 0 {
            write!(f, "{:04}-{:02}-{:02}", self.year, self.month, self.day)
        } else {
            write!(
                f,
                "{:04}-{:02}-{:02} {:02}:{:02}:{:02}",
                self.year, self.month, self.day, self.hour, self.minute, self.second
            )
        }
    }
}

/// Largest serial day count `DateTimeValue::from_serial` will convert.
///
/// Excel's own last representable date, 9999-12-31, is serial 2,958,465;
/// this sits well past it while keeping the calendar walk in
/// `serial_to_date_1900`/`_1904` bounded to a few thousand iterations.
pub const MAX_DATE_SERIAL: f64 = 5_000_000.0;

impl DateTimeValue {
    /// Format as ISO 8601 string.
    pub fn to_iso_string(&self) -> String {
        self.to_string()
    }

    /// Convert an Excel serial number to a date/time value.
    ///
    /// Excel stores dates as floating-point days since a base date:
    /// - 1900 system (default): Day 1 = Jan 1, 1900
    /// - 1904 system (Mac): Day 0 = Jan 1, 1904
    ///
    /// The 1900 system has a known bug: serial 60 = Feb 29, 1900 (which doesn't exist).
    /// Serials 1-59 correspond to Jan 1 – Feb 28, 1900.
    /// Serials >= 61 are off by one day compared to reality.
    pub fn from_serial(serial: f64, date1904: bool) -> Option<Self> {
        if !(0.0..=MAX_DATE_SERIAL).contains(&serial) {
            // NaN, negative, and out-of-calendar-range magnitudes all land
            // here. The upper bound matters for more than tidiness: the
            // year-by-year loops below are linear in the serial, and an `as
            // i64` cast saturates rather than erroring, so a cell holding
            // 1e300 under a date-classified style used to spin for
            // ~2.5e16 iterations.
            return None;
        }
        // `serial_to_date_*` walk the calendar a year at a time, so the work
        // they do is proportional to the input. `serial.trunc() as i64`
        // saturates rather than erroring for out-of-range floats, so a cell
        // holding 1e300 asked for ~2.5e16 iterations — a hang, not a slow
        // answer, on ordinary (non-adversarial) scientific-notation values
        // that a misclassifying caller routed here. Bound the input itself
        // so no caller, present or future, can reach that loop with a value
        // outside any real calendar date (MAX_DATE_SERIAL is comfortably
        // past 9999-12-31, which Excel itself caps at serial 2,958,465).
        if !serial.is_finite() || serial > MAX_DATE_SERIAL {
            return None;
        }

        // Split into integer days and fractional time
        let day_serial = serial.trunc() as i64;
        let mut time_frac = serial - serial.trunc();
        let mut day_serial = day_serial;
        // Rounding to the nearest second can reach a full day. Carry it into
        // the date rather than emitting hour 24, which no date library accepts.
        if (time_frac * 86400.0).round() as u64 >= 86_400 {
            time_frac = 0.0;
            day_serial += 1;
        }

        let (year, month, day) = if date1904 {
            // 1904 system: day 0 = Jan 1, 1904
            serial_to_date_1904(day_serial)?
        } else {
            // 1900 system: day 1 = Jan 1, 1900
            serial_to_date_1900(day_serial)?
        };

        // Convert fractional day to hours/minutes/seconds
        let total_seconds = (time_frac * 86400.0).round() as u64;
        let hour = (total_seconds / 3600) as u32;
        let minute = ((total_seconds % 3600) / 60) as u32;
        let second = (total_seconds % 60) as u32;
        let millisecond = ((time_frac * 86_400_000.0).round() as u64 % 1000) as u32;

        Some(DateTimeValue {
            year,
            month,
            day,
            hour,
            minute,
            second,
            millisecond,
        })
    }
}

/// Days in each month for non-leap and leap years.
const DAYS_IN_MONTH: [[u32; 12]; 2] = [
    [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31], // non-leap
    [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31], // leap
];

fn is_leap_year(y: i32) -> bool {
    (y % 4 == 0 && y % 100 != 0) || y % 400 == 0
}

/// Convert 1900-system serial to (year, month, day).
fn serial_to_date_1900(serial: i64) -> Option<(i32, u32, u32)> {
    if serial < 1 || serial > MAX_DATE_SERIAL as i64 {
        return None;
    }
    // Serial 60 is Excel's phantom "29 February 1900", kept for Lotus 1-2-3
    // compatibility. That date does not exist — 1900 was not a leap year —
    // and emitting it produces a string that any strict date parser
    // downstream (chrono, Python `datetime`) rejects. Refuse it instead;
    // callers fall back to rendering the raw serial, which is at least true.
    if serial == 60 {
        return None;
    }

    // Adjust for the bug: serials >= 61 are one day ahead
    let adjusted = if serial > 60 { serial - 1 } else { serial };

    // Day 1 = Jan 1, 1900 → convert to 0-based days since Jan 1, 1900
    let mut remaining = adjusted - 1;

    let mut year = 1900i32;
    loop {
        let days_in_year = if is_leap_year(year) { 366 } else { 365 };
        if remaining < days_in_year {
            break;
        }
        remaining -= days_in_year;
        year += 1;
    }

    let leap = if is_leap_year(year) { 1 } else { 0 };
    let mut month = 0u32;
    for (m, &dim) in DAYS_IN_MONTH[leap].iter().enumerate() {
        let dim = dim as i64;
        if remaining < dim {
            month = m as u32 + 1;
            break;
        }
        remaining -= dim;
    }

    let day = remaining as u32 + 1;
    Some((year, month, day))
}

/// Convert 1904-system serial to (year, month, day).
fn serial_to_date_1904(serial: i64) -> Option<(i32, u32, u32)> {
    if serial < 0 || serial > MAX_DATE_SERIAL as i64 {
        return None;
    }
    // Day 0 = Jan 1, 1904
    let mut remaining = serial;
    let mut year = 1904i32;

    loop {
        let days_in_year = if is_leap_year(year) { 366 } else { 365 };
        if remaining < days_in_year {
            break;
        }
        remaining -= days_in_year;
        year += 1;
    }

    let leap = if is_leap_year(year) { 1 } else { 0 };
    let mut month = 0u32;
    for (m, &dim) in DAYS_IN_MONTH[leap].iter().enumerate() {
        let dim = dim as i64;
        if remaining < dim {
            month = m as u32 + 1;
            break;
        }
        remaining -= dim;
    }

    let day = remaining as u32 + 1;
    Some((year, month, day))
}

/// Built-in number format IDs that represent dates/times.
/// These are the standard Excel built-in date format IDs.
pub fn is_date_format_id(id: u32) -> bool {
    matches!(
        id,
        14..=22 | 27..=36 | 45..=47 | 50..=58
    )
}

/// Check if a custom number format string indicates a date/time format.
///
/// Scans for date/time tokens (y, m, d, h, s, AM/PM) while ignoring
/// escaped characters and quoted sections.
pub fn is_date_format_string(format: &str) -> bool {
    let mut chars = format.chars().peekable();
    let mut has_date_token = false;

    while let Some(c) = chars.next() {
        match c {
            // Skip escaped character
            '\\' => {
                chars.next();
            },
            // Skip quoted section
            '"' => {
                for ch in chars.by_ref() {
                    if ch == '"' {
                        break;
                    }
                }
            },
            // Skip bracketed sections like [Red], [$-409]
            '[' => {
                for ch in chars.by_ref() {
                    if ch == ']' {
                        break;
                    }
                }
            },
            // Date/time tokens
            'y' | 'Y' | 'd' | 'D' | 'h' | 'H' | 's' | 'S' => {
                has_date_token = true;
            },
            // 'm' is ambiguous (month or minute) — consider it date-like
            'm' | 'M' => {
                has_date_token = true;
            },
            // AM/PM marker
            'A' | 'a' if (chars.peek() == Some(&'M') || chars.peek() == Some(&'m')) => {
                has_date_token = true;
            },
            _ => {},
        }
    }

    has_date_token
}

/// Check if a cell should be treated as a date cell given its style.
pub fn is_date_cell(style_index: Option<u32>, styles: Option<&StyleSheet>) -> bool {
    let Some(idx) = style_index else {
        return false;
    };
    let Some(styles) = styles else {
        return false;
    };
    let Some(fmt_id) = styles.number_format_id_for(idx) else {
        return false;
    };

    // A workbook may redefine a built-in id — [ECMA-376] §18.8.30 permits
    // ids 0-163 to be overridden — so an explicit <numFmt> wins over the
    // built-in meaning of its id. Testing the id first made a cell formatted
    // `0.00" kg"` under id 14 render as a 1900 date.
    if let Some(fmt_str) = styles.number_format_override_for(idx) {
        return is_date_format_string(fmt_str);
    }

    is_date_format_id(fmt_id)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_serial_to_date_1900_basic() {
        // Serial 1 = Jan 1, 1900
        let dt = DateTimeValue::from_serial(1.0, false).unwrap();
        assert_eq!(dt.year, 1900);
        assert_eq!(dt.month, 1);
        assert_eq!(dt.day, 1);
    }

    #[test]
    fn test_serial_to_date_1900_feb28() {
        // Serial 59 = Feb 28, 1900
        let dt = DateTimeValue::from_serial(59.0, false).unwrap();
        assert_eq!(dt.year, 1900);
        assert_eq!(dt.month, 2);
        assert_eq!(dt.day, 28);
    }

    #[test]
    fn test_serial_60_is_refused_rather_than_emitting_an_impossible_date() {
        // Excel's phantom leap day. 1900-02-29 never existed, so returning
        // it hands downstream date parsers a string they must reject.
        assert!(DateTimeValue::from_serial(60.0, false).is_none());
        // The serials on either side are unaffected.
        assert!(DateTimeValue::from_serial(59.0, false).is_some());
        assert!(DateTimeValue::from_serial(61.0, false).is_some());
    }

    #[test]
    fn test_serial_to_date_1900_mar1() {
        // Serial 61 = Mar 1, 1900
        let dt = DateTimeValue::from_serial(61.0, false).unwrap();
        assert_eq!(dt.year, 1900);
        assert_eq!(dt.month, 3);
        assert_eq!(dt.day, 1);
    }

    #[test]
    fn test_serial_to_date_2024_jan_15() {
        // Jan 15, 2024 = serial 45306
        let dt = DateTimeValue::from_serial(45306.0, false).unwrap();
        assert_eq!(dt.year, 2024);
        assert_eq!(dt.month, 1);
        assert_eq!(dt.day, 15);
    }

    #[test]
    fn test_serial_to_date_with_time() {
        // Serial 45306.5 = Jan 15, 2024 at 12:00:00
        let dt = DateTimeValue::from_serial(45306.5, false).unwrap();
        assert_eq!(dt.year, 2024);
        assert_eq!(dt.month, 1);
        assert_eq!(dt.day, 15);
        assert_eq!(dt.hour, 12);
        assert_eq!(dt.minute, 0);
        assert_eq!(dt.second, 0);
    }

    #[test]
    fn test_serial_to_date_1904_system() {
        // Day 0 in 1904 system = Jan 1, 1904
        let dt = DateTimeValue::from_serial(0.0, true).unwrap();
        assert_eq!(dt.year, 1904);
        assert_eq!(dt.month, 1);
        assert_eq!(dt.day, 1);
    }

    #[test]
    fn test_iso_string_date_only() {
        let dt = DateTimeValue {
            year: 2024,
            month: 1,
            day: 15,
            hour: 0,
            minute: 0,
            second: 0,
            millisecond: 0,
        };
        assert_eq!(dt.to_iso_string(), "2024-01-15");
    }

    #[test]
    fn test_iso_string_with_time() {
        let dt = DateTimeValue {
            year: 2024,
            month: 1,
            day: 15,
            hour: 14,
            minute: 30,
            second: 45,
            millisecond: 0,
        };
        assert_eq!(dt.to_iso_string(), "2024-01-15 14:30:45");
    }

    #[test]
    fn test_builtin_date_format_ids() {
        assert!(is_date_format_id(14));
        assert!(is_date_format_id(22));
        assert!(is_date_format_id(45));
        assert!(!is_date_format_id(0));
        assert!(!is_date_format_id(1));
        assert!(!is_date_format_id(13));
    }

    #[test]
    fn test_custom_date_format_detection() {
        assert!(is_date_format_string("yyyy-mm-dd"));
        assert!(is_date_format_string("dd/mm/yyyy"));
        assert!(is_date_format_string("h:mm:ss AM/PM"));
        assert!(is_date_format_string("m/d/yy"));
        assert!(!is_date_format_string("#,##0.00"));
        assert!(!is_date_format_string("0%"));
        assert!(!is_date_format_string("General"));
    }

    #[test]
    fn test_date_format_ignores_quoted() {
        // Quoted text should not trigger date detection
        assert!(!is_date_format_string("\"day\""));
        assert!(!is_date_format_string("#,##0.00\" days\""));
    }

    #[test]
    fn test_negative_serial_returns_none() {
        assert!(DateTimeValue::from_serial(-1.0, false).is_none());
    }

    /// The year-by-year scan is linear in the serial and `as i64`
    /// saturates rather than erroring, so an unbounded input meant an
    /// effectively infinite loop. The bound lives in the converter itself,
    /// so a future caller that misclassifies a cell as a date can't
    /// reintroduce the hang.
    #[test]
    fn test_out_of_range_serial_is_rejected_promptly() {
        let started = std::time::Instant::now();
        for serial in [
            MAX_DATE_SERIAL + 1.0,
            1e12,
            1e300,
            f64::MAX,
            f64::INFINITY,
            f64::NAN,
        ] {
            assert!(
                DateTimeValue::from_serial(serial, false).is_none(),
                "{serial} is not a calendar date"
            );
            assert!(
                DateTimeValue::from_serial(serial, true).is_none(),
                "{serial} is not a calendar date (1904)"
            );
        }
        assert!(
            started.elapsed() < std::time::Duration::from_millis(500),
            "rejection must not loop: took {:?}",
            started.elapsed()
        );

        // Every real Excel date still converts — the cap sits well above
        // Excel's own maximum of 2_958_465 (9999-12-31).
        let last = DateTimeValue::from_serial(2_958_465.0, false).expect("9999-12-31 converts");
        assert_eq!((last.year, last.month, last.day), (9999, 12, 31));
    }
}

#[cfg(test)]
mod override_tests {
    use super::*;

    /// Rounding the time fraction to the nearest second can reach a whole
    /// day. The day was never carried, so the value came out as hour 24 —
    /// which `chrono` and Python's `datetime` both reject.
    #[test]
    fn test_a_time_that_rounds_up_to_a_full_day_carries_into_the_date() {
        let v = DateTimeValue::from_serial(45000.9999999, false).expect("valid serial");
        assert!(v.hour < 24, "hour must stay in 0..=23, got {}", v.hour);
        assert_eq!((v.hour, v.minute, v.second), (0, 0, 0));

        let prev = DateTimeValue::from_serial(45000.0, false).unwrap();
        assert!(
            (v.year, v.month, v.day) > (prev.year, prev.month, prev.day),
            "the rounded-up day must advance the date"
        );
    }

    /// The calendar walk costs one iteration per year, and `as i64`
    /// saturates rather than erroring, so a cell holding 1e300 asked for
    /// ~2.5e16 iterations and never returned. Out-of-calendar serials are
    /// refused up front; callers then render the raw value.
    #[test]
    fn test_out_of_calendar_serials_are_refused_instead_of_hanging() {
        for serial in [1e300, 1e12, 1e10, f64::MAX, MAX_DATE_SERIAL + 1.0] {
            let started = std::time::Instant::now();
            assert!(
                DateTimeValue::from_serial(serial, false).is_none(),
                "{serial} is not a calendar date"
            );
            assert!(DateTimeValue::from_serial(serial, true).is_none());
            assert!(
                started.elapsed() < std::time::Duration::from_secs(1),
                "{serial} took {:?}",
                started.elapsed()
            );
        }
        assert!(DateTimeValue::from_serial(f64::NAN, false).is_none());
        assert!(DateTimeValue::from_serial(f64::INFINITY, false).is_none());

        // Every date Excel itself can hold still converts: 9999-12-31 is
        // serial 2,958,465, comfortably inside the bound.
        let last = DateTimeValue::from_serial(2_958_465.0, false).expect("9999-12-31");
        assert_eq!((last.year, last.month, last.day), (9999, 12, 31));
    }
}
