//! Warnings that carry what they mean.
//!
//! A reader or converter that works around a problem says so in one
//! human-readable line (`Metadata::warnings`, each reader's `warnings()`).
//! [`Document::damage`](crate::Document::damage) and strict mode have to tell
//! a loss from a harmless repair, which they once did by looking for words
//! such as "truncat" in those lines: a reworded or new warning was misfiled
//! with no test failing. Every warning is now recorded with its kind at the
//! point it is raised, through [`Notices::push`], the only way to add one;
//! the text is kept exactly as before.

use crate::options::DamageKind;

/// Warnings in the order they were raised, each with its kind.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct Notices {
    texts: Vec<String>,
    kinds: Vec<DamageKind>,
}

impl Notices {
    /// An empty list.
    pub(crate) fn new() -> Self {
        Self::default()
    }

    /// Record one warning of `kind`.
    pub(crate) fn push(&mut self, kind: DamageKind, text: impl Into<String>) {
        self.texts.push(text.into());
        self.kinds.push(kind);
    }

    /// Record each of `texts` as a warning of `kind`.
    pub(crate) fn push_all<S: Into<String>>(
        &mut self,
        kind: DamageKind,
        texts: impl IntoIterator<Item = S>,
    ) {
        for text in texts {
            self.push(kind, text);
        }
    }

    /// Append every warning of `other`, keeping its kinds.
    pub(crate) fn append(&mut self, other: Notices) {
        self.texts.extend(other.texts);
        self.kinds.extend(other.kinds);
    }

    /// The warning lines, in order.
    pub(crate) fn texts(&self) -> &[String] {
        &self.texts
    }

    /// Each warning with its kind, in order.
    pub(crate) fn iter(&self) -> impl Iterator<Item = (DamageKind, &str)> {
        self.kinds
            .iter()
            .copied()
            .zip(self.texts.iter().map(String::as_str))
    }

    /// The warning lines alone, for `Metadata::warnings`.
    pub(crate) fn into_texts(self) -> Vec<String> {
        self.texts
    }
}

impl FromIterator<(DamageKind, String)> for Notices {
    fn from_iter<I: IntoIterator<Item = (DamageKind, String)>>(iter: I) -> Self {
        let mut out = Notices::new();
        for (kind, text) in iter {
            out.push(kind, text);
        }
        out
    }
}

thread_local! {
    /// The warnings the last conversion on this thread produced, while a
    /// [`ReportScope`] is open: `None` outside one.
    static REPORT: std::cell::RefCell<Option<Option<Notices>>> =
        const { std::cell::RefCell::new(None) };
}

/// Collects the typed warnings of the conversions run on this thread while it
/// is open, for [`Document::damage`](crate::Document::damage): the IR carries
/// only the text of each warning. A conversion reports its document-level
/// warnings when it assembles them; a document embedded in it is converted
/// first, so the last report is the outer document's.
#[must_use = "the scope ends as soon as it is dropped"]
pub(crate) struct ReportScope {
    /// What was open on this thread before, restored when the scope ends.
    previous: Option<Option<Notices>>,
    finished: bool,
}

impl ReportScope {
    pub(crate) fn begin() -> Self {
        let previous = REPORT.with(|r| r.borrow_mut().replace(None));
        Self {
            previous,
            finished: false,
        }
    }

    /// The warnings the last conversion reported, if one did.
    pub(crate) fn finish(mut self) -> Option<Notices> {
        self.finished = true;
        let previous = self.previous.take();
        REPORT
            .with(|r| std::mem::replace(&mut *r.borrow_mut(), previous))
            .flatten()
    }
}

impl Drop for ReportScope {
    fn drop(&mut self) {
        if !self.finished {
            let previous = self.previous.take();
            REPORT.with(|r| *r.borrow_mut() = previous);
        }
    }
}

/// Record `notices` as the document-level warnings of the conversion that
/// just finished, if a [`ReportScope`] is open on this thread; and return
/// their text for `Metadata::warnings`.
pub(crate) fn report(notices: Notices) -> Vec<String> {
    REPORT.with(|r| {
        if let Some(slot) = r.borrow_mut().as_mut() {
            *slot = Some(notices.clone());
        }
    });
    notices.into_texts()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_notices_keep_order_and_kind() {
        let mut n = Notices::new();
        n.push(DamageKind::Recovered, "repaired");
        n.push(DamageKind::Truncated, String::from("cut"));
        let mut more = Notices::new();
        more.push_all(DamageKind::UnreadablePart, ["a", "b"]);
        n.append(more);
        assert_eq!(n.texts(), ["repaired", "cut", "a", "b"]);
        let kinds: Vec<_> = n.iter().map(|(k, _)| k).collect();
        assert_eq!(
            kinds,
            [
                DamageKind::Recovered,
                DamageKind::Truncated,
                DamageKind::UnreadablePart,
                DamageKind::UnreadablePart
            ]
        );
        assert_eq!(n.into_texts().len(), 4);
    }

    #[test]
    fn test_report_scope_keeps_the_last_report_and_restores_the_outer_scope() {
        let mut a = Notices::new();
        a.push(DamageKind::Recovered, "inner");
        let mut b = Notices::new();
        b.push(DamageKind::Truncated, "outer");
        // Outside a scope nothing is kept.
        assert_eq!(report(a.clone()), ["inner"]);
        let outer = ReportScope::begin();
        let _ = report(a.clone());
        {
            let nested = ReportScope::begin();
            let _ = report(b.clone());
            assert_eq!(nested.finish(), Some(b.clone()));
        }
        let _ = report(b.clone());
        assert_eq!(outer.finish(), Some(b));
        let empty = ReportScope::begin();
        assert_eq!(empty.finish(), None);
        // A scope dropped without finishing closes too.
        drop(ReportScope::begin());
        let _ = report(a);
        assert_eq!(ReportScope::begin().finish(), None);
    }
}
