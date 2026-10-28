//! A bounded record of the parts of a package that could not be read.
//!
//! A reader that skips a part it cannot use reports the loss; what it
//! reports is shaped by the file, so the report is bounded the same way for
//! every format: each part once (however many references name it), at most
//! [`MAX_UNREADABLE_PARTS`] kept, a part name or reason cut at
//! [`MAX_REPORT_CHARS`] characters (both can echo values the file controls),
//! and the number of the rest counted.
//!
//! "Each part once" is decided on the exact normalised name, never on a
//! hash of it: a hash of a fixed width can be made to collide with another
//! part's, and the collision would hide that part's failure.

use std::collections::HashSet;

/// Most parts recorded as unreadable; the rest are counted.
pub(crate) const MAX_UNREADABLE_PARTS: usize = 100;
/// Longest part name or reason kept in a record, in characters.
pub(crate) const MAX_REPORT_CHARS: usize = 256;
/// Most bytes of normalised names remembered to tell a part already recorded
/// from a new one. Past it a name is not remembered, so it is treated as new
/// each time it is met: it can then be counted more than once among the rest,
/// but a part is never taken for one already recorded when it is not.
const MAX_SEEN_BYTES: usize = 1 << 20;

/// The parts that could not be read: each once, at most
/// [`MAX_UNREADABLE_PARTS`] kept, and the number of the rest.
#[derive(Debug, Clone, Default)]
pub(crate) struct UnreadableParts {
    /// Prefix of the log line, naming the reader.
    label: &'static str,
    listed: Vec<(String, String)>,
    omitted: usize,
    /// The normalised name of each part recorded so far (up to
    /// [`MAX_SEEN_BYTES`] of them).
    seen: HashSet<String>,
    seen_bytes: usize,
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
        let key: String = part
            .chars()
            .map(|c| {
                if c == '\\' {
                    '/'
                } else {
                    c.to_ascii_lowercase()
                }
            })
            .collect();
        if self.seen.contains(&key) {
            return;
        }
        if self.seen_bytes + key.len() <= MAX_SEEN_BYTES {
            self.seen_bytes += key.len();
            self.seen.insert(key);
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

    /// The record keeps the names themselves, so two parts can never be
    /// taken for one; past the memory bound a part is counted again, never
    /// dropped.
    #[test]
    fn test_parts_are_told_apart_by_their_names_not_by_a_hash() {
        let mut sink = UnreadableParts::new("test");
        sink.record("/Word\\A.xml", &"bad");
        assert!(sink.seen.contains("/word/a.xml"));
        let long = |i: usize| format!("{i:08}{}", "n".repeat(MAX_SEEN_BYTES / 4));
        for i in 0..6 {
            sink.record(&long(i), &"bad");
            sink.record(&long(i), &"bad");
        }
        let (listed, _) = sink.finish();
        // Every distinct part is listed; the ones past the bound may be
        // listed again, none is lost.
        for i in 0..6 {
            let prefix = &long(i)[..8];
            assert!(listed.iter().any(|(p, _)| p.starts_with(prefix)), "{i}");
        }
        assert!(listed.iter().any(|(p, _)| p == "/Word\\A.xml"));
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
