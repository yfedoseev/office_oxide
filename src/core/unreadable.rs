//! A bounded record of the parts of a package that could not be read.
//!
//! A reader that skips a part it cannot use reports the loss; what it
//! reports is shaped by the file, so the report is bounded the same way for
//! every format: each part once (however many references name it), at most
//! [`MAX_UNREADABLE_PARTS`] kept, a part name or reason cut at
//! [`MAX_REPORT_CHARS`] characters (both can echo values the file controls),
//! and the number of the rest counted.

use std::collections::HashSet;
use std::hash::{Hash, Hasher};

/// Most parts recorded as unreadable; the rest are counted.
pub(crate) const MAX_UNREADABLE_PARTS: usize = 100;
/// Longest part name or reason kept in a record, in characters.
pub(crate) const MAX_REPORT_CHARS: usize = 256;

/// The parts that could not be read: each once, at most
/// [`MAX_UNREADABLE_PARTS`] kept, and the number of the rest.
#[derive(Debug, Clone, Default)]
pub(crate) struct UnreadableParts {
    /// Prefix of the log line, naming the reader.
    label: &'static str,
    listed: Vec<(String, String)>,
    omitted: usize,
    /// A hash of each part's normalised name seen so far, so that a name
    /// the file makes very long costs eight bytes, not its length.
    seen: HashSet<u64>,
}

impl UnreadableParts {
    /// An empty record for the reader called `label` (`"docx"`, `"pptx"`).
    pub(crate) fn new(label: &'static str) -> Self {
        Self {
            label,
            ..Self::default()
        }
    }

    /// Record that `part` could not be read, for `reason`. A part already
    /// recorded (compared without regard to case or `\` against `/`, as a
    /// zip lookup is) is not recorded again; it is logged once, when it is
    /// first met.
    pub(crate) fn record(&mut self, part: &str, reason: &dyn std::fmt::Display) {
        let mut hasher = std::collections::hash_map::DefaultHasher::new();
        for c in part.chars() {
            let c = if c == '\\' {
                '/'
            } else {
                c.to_ascii_lowercase()
            };
            c.hash(&mut hasher);
        }
        if !self.seen.insert(hasher.finish()) {
            return;
        }
        if self.listed.len() >= MAX_UNREADABLE_PARTS {
            self.omitted += 1;
            return;
        }
        let cut = |s: &str| s.chars().take(MAX_REPORT_CHARS).collect::<String>();
        let (part, reason) = (cut(part), cut(&reason.to_string()));
        log::warn!("{}: skipping unreadable part {part}: {reason}", self.label);
        self.listed.push((part, reason));
    }

    /// The first part recorded, with its reason.
    pub(crate) fn first(&self) -> Option<&(String, String)> {
        self.listed.first()
    }

    /// The parts kept and the number left out.
    pub(crate) fn finish(self) -> (Vec<(String, String)>, usize) {
        (self.listed, self.omitted)
    }
}

/// The warning for `omitted` parts left out of a list.
pub(crate) fn omitted_warning(omitted: usize) -> String {
    format!("and {omitted} more unreadable parts not listed")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_a_part_is_recorded_once_whatever_its_spelling() {
        let mut sink = UnreadableParts::new("test");
        sink.record("/word/header1.xml", &"bad");
        sink.record("/WORD/Header1.xml", &"bad");
        sink.record("\\word\\header1.xml", &"worse");
        assert_eq!(sink.finish(), (vec![("/word/header1.xml".into(), "bad".into())], 0));
    }

    #[test]
    fn test_the_list_is_capped_and_the_rest_counted() {
        let mut sink = UnreadableParts::new("test");
        for i in 0..MAX_UNREADABLE_PARTS + 7 {
            sink.record(&format!("/p{i}.xml"), &"bad");
            sink.record(&format!("/p{i}.xml"), &"bad");
        }
        let (listed, omitted) = sink.finish();
        assert_eq!((listed.len(), omitted), (MAX_UNREADABLE_PARTS, 7));
    }

    #[test]
    fn test_names_and_reasons_are_cut() {
        let mut sink = UnreadableParts::new("test");
        sink.record(&"n".repeat(1000), &"r".repeat(1000));
        let (listed, _) = sink.finish();
        assert_eq!(listed[0].0.chars().count(), MAX_REPORT_CHARS);
        assert_eq!(listed[0].1.chars().count(), MAX_REPORT_CHARS);
    }
}
