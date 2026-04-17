use std::fmt;

use super::styles::StyleSheet;

/// A date/time value converted from an Excel serial number.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DateTimeValue {
    pub year: i32,
    pub month: u32,
    pub day: u32,
    pub hour: u32,
    pub minute: u32,
    pub second: u32,
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
        if serial < 0.0 {
            return None;
        }

        // Split into integer days and fractional time
        let day_serial = serial.trunc() as i64;
        let time_frac = serial - serial.trunc();

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
    if serial < 1 {
        return None;
    }
    // Handle the Lotus 1-2-3 bug: serial 60 = Feb 29, 1900 (fictitious)
    if serial == 60 {
        return Some((1900, 2, 29));
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
    if serial < 0 {
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

    // Check built-in date format IDs first
    if is_date_format_id(fmt_id) {
        return true;
    }

    // Check custom format string
    if let Some(fmt_str) = styles.number_format_for(idx) {
        return is_date_format_string(fmt_str);
    }

    false
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn serial_to_date_1900_basic() {
        // Serial 1 = Jan 1, 1900
        let dt = DateTimeValue::from_serial(1.0, false).unwrap();
        assert_eq!(dt.year, 1900);
        assert_eq!(dt.month, 1);
        assert_eq!(dt.day, 1);
    }

    #[test]
    fn serial_to_date_1900_feb28() {
        // Serial 59 = Feb 28, 1900
        let dt = DateTimeValue::from_serial(59.0, false).unwrap();
        assert_eq!(dt.year, 1900);
        assert_eq!(dt.month, 2);
        assert_eq!(dt.day, 28);
    }

    #[test]
    fn serial_to_date_1900_bug_feb29() {
        // Serial 60 = Feb 29, 1900 (the Lotus 1-2-3 bug)
        let dt = DateTimeValue::from_serial(60.0, false).unwrap();
        assert_eq!(dt.year, 1900);
        assert_eq!(dt.month, 2);
        assert_eq!(dt.day, 29);
    }

    #[test]
    fn serial_to_date_1900_mar1() {
        // Serial 61 = Mar 1, 1900
        let dt = DateTimeValue::from_serial(61.0, false).unwrap();
        assert_eq!(dt.year, 1900);
        assert_eq!(dt.month, 3);
        assert_eq!(dt.day, 1);
    }

    #[test]
    fn serial_to_date_2024_jan_15() {
        // Jan 15, 2024 = serial 45306
        let dt = DateTimeValue::from_serial(45306.0, false).unwrap();
        assert_eq!(dt.year, 2024);
        assert_eq!(dt.month, 1);
        assert_eq!(dt.day, 15);
    }

    #[test]
    fn serial_to_date_with_time() {
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
    fn serial_to_date_1904_system() {
        // Day 0 in 1904 system = Jan 1, 1904
        let dt = DateTimeValue::from_serial(0.0, true).unwrap();
        assert_eq!(dt.year, 1904);
        assert_eq!(dt.month, 1);
        assert_eq!(dt.day, 1);
    }

    #[test]
    fn iso_string_date_only() {
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
    fn iso_string_with_time() {
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
    fn builtin_date_format_ids() {
        assert!(is_date_format_id(14));
        assert!(is_date_format_id(22));
        assert!(is_date_format_id(45));
        assert!(!is_date_format_id(0));
        assert!(!is_date_format_id(1));
        assert!(!is_date_format_id(13));
    }

    #[test]
    fn custom_date_format_detection() {
        assert!(is_date_format_string("yyyy-mm-dd"));
        assert!(is_date_format_string("dd/mm/yyyy"));
        assert!(is_date_format_string("h:mm:ss AM/PM"));
        assert!(is_date_format_string("m/d/yy"));
        assert!(!is_date_format_string("#,##0.00"));
        assert!(!is_date_format_string("0%"));
        assert!(!is_date_format_string("General"));
    }

    #[test]
    fn date_format_ignores_quoted() {
        // Quoted text should not trigger date detection
        assert!(!is_date_format_string("\"day\""));
        assert!(!is_date_format_string("#,##0.00\" days\""));
    }

    #[test]
    fn negative_serial_returns_none() {
        assert!(DateTimeValue::from_serial(-1.0, false).is_none());
    }
}
