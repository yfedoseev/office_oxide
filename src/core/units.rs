use serde::{Deserialize, Serialize};

/// English Metric Unit (914,400 per inch). Used for positions/dimensions in DrawingML.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct Emu(pub i64);

impl Emu {
    /// Number of EMUs per inch.
    pub const PER_INCH: i64 = 914_400;
    /// Number of EMUs per centimetre.
    pub const PER_CM: i64 = 360_000;
    /// Number of EMUs per point.
    pub const PER_PT: i64 = 12_700;
    /// Number of EMUs per pixel at 96 DPI.
    pub const PER_PIXEL_96DPI: i64 = 9_525;

    /// Create an `Emu` from a measurement in inches.
    pub fn from_inches(inches: f64) -> Self {
        Self((inches * Self::PER_INCH as f64) as i64)
    }

    /// Create an `Emu` from a measurement in centimetres.
    pub fn from_cm(cm: f64) -> Self {
        Self((cm * Self::PER_CM as f64) as i64)
    }

    /// Create an `Emu` from a measurement in points.
    pub fn from_pt(pt: f64) -> Self {
        Self((pt * Self::PER_PT as f64) as i64)
    }

    /// Convert to inches.
    pub fn to_inches(self) -> f64 {
        self.0 as f64 / Self::PER_INCH as f64
    }

    /// Convert to centimetres.
    pub fn to_cm(self) -> f64 {
        self.0 as f64 / Self::PER_CM as f64
    }

    /// Convert to points.
    pub fn to_pt(self) -> f64 {
        self.0 as f64 / Self::PER_PT as f64
    }

    /// Convert to twips. 1 twip = 914400/1440 = 635 EMU.
    pub fn to_twip(self) -> Twip {
        Twip((self.0 / 635) as i32)
    }
}

/// Twip = 1/20 of a point = 1/1440 of an inch. Used in WordprocessingML.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct Twip(pub i32);

impl Twip {
    /// Number of twips per inch.
    pub const PER_INCH: i32 = 1_440;
    /// Number of twips per point.
    pub const PER_PT: i32 = 20;

    /// Create a `Twip` from a measurement in points.
    pub fn from_pt(pt: f64) -> Self {
        Self((pt * Self::PER_PT as f64) as i32)
    }

    /// Create a `Twip` from a measurement in inches.
    pub fn from_inches(inches: f64) -> Self {
        Self((inches * Self::PER_INCH as f64) as i32)
    }

    /// Convert to `Emu` (approximate; loses sub-twip precision).
    pub fn to_emu(self) -> Emu {
        Emu(self.0 as i64 * 635)
    }

    /// Convert to inches.
    pub fn to_inches(self) -> f64 {
        self.0 as f64 / Self::PER_INCH as f64
    }

    /// Convert to points.
    pub fn to_pt(self) -> f64 {
        self.0 as f64 / Self::PER_PT as f64
    }
}

/// Half-point = 1/2 of a point. Used for font sizes in OOXML.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct HalfPoint(pub u32);

impl HalfPoint {
    /// Convert to points.
    pub fn to_points(self) -> f64 {
        self.0 as f64 / 2.0
    }

    /// Create a `HalfPoint` from a point value.
    pub fn from_points(pt: f64) -> Self {
        Self((pt * 2.0) as u32)
    }
}

/// Percentage * 1000 (e.g., 50% = 50_000, 100% = 100_000). ST_Percentage in OOXML.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct Percentage1000(pub i32);

impl Percentage1000 {
    /// Convert to a fraction in the range 0.0–1.0.
    pub fn to_fraction(self) -> f64 {
        self.0 as f64 / 100_000.0
    }

    /// Convert to a percentage value (e.g., 50.0 for 50%).
    pub fn to_percent(self) -> f64 {
        self.0 as f64 / 1_000.0
    }

    /// Create from a percentage value (e.g., `50.0` for 50%).
    pub fn from_percent(pct: f64) -> Self {
        Self((pct * 1_000.0) as i32)
    }
}

/// Angle in 60,000ths of a degree. ST_Angle in DrawingML.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct Angle60k(pub i32);

impl Angle60k {
    /// Convert to degrees.
    pub fn to_degrees(self) -> f64 {
        self.0 as f64 / 60_000.0
    }

    /// Create from a degree value.
    pub fn from_degrees(deg: f64) -> Self {
        Self((deg * 60_000.0) as i32)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn emu_conversions() {
        let one_inch = Emu::from_inches(1.0);
        assert_eq!(one_inch.0, 914_400);
        assert!((one_inch.to_inches() - 1.0).abs() < f64::EPSILON);

        let one_cm = Emu::from_cm(1.0);
        assert_eq!(one_cm.0, 360_000);

        let one_pt = Emu::from_pt(1.0);
        assert_eq!(one_pt.0, 12_700);
    }

    #[test]
    fn emu_to_twip_roundtrip() {
        let emu = Emu::from_inches(1.0);
        let twip = emu.to_twip();
        assert_eq!(twip.0, 1440);
        // Round-trip loses some precision due to integer division
        let back = twip.to_emu();
        assert!((back.0 - emu.0).abs() < 635);
    }

    #[test]
    fn twip_conversions() {
        let us_letter_width = Twip(12240); // 8.5 inches
        assert!((us_letter_width.to_inches() - 8.5).abs() < f64::EPSILON);

        let twelve_pt = Twip::from_pt(12.0);
        assert_eq!(twelve_pt.0, 240);
    }

    #[test]
    fn half_point_conversions() {
        let twelve_pt = HalfPoint(24);
        assert!((twelve_pt.to_points() - 12.0).abs() < f64::EPSILON);

        let from = HalfPoint::from_points(10.0);
        assert_eq!(from.0, 20);
    }

    #[test]
    fn percentage_conversions() {
        let fifty = Percentage1000(50_000);
        assert!((fifty.to_percent() - 50.0).abs() < f64::EPSILON);
        assert!((fifty.to_fraction() - 0.5).abs() < f64::EPSILON);
    }

    #[test]
    fn angle_conversions() {
        let right_angle = Angle60k(5_400_000);
        assert!((right_angle.to_degrees() - 90.0).abs() < f64::EPSILON);

        let from = Angle60k::from_degrees(45.0);
        assert_eq!(from.0, 2_700_000);
    }
}
