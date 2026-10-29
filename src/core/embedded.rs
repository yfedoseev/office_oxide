//! Bounds on opening the packages a document embeds.
//!
//! A `.docx` can embed another OOXML package (`<o:OLEObject Type="Embed">`),
//! and that package can embed another. Opened as an independent document each
//! time, with its own limits, the nesting grows linearly in file size and
//! exponentially in work: entries that name one blob let a 20 KB file embed
//! ten copies of a package that embeds ten copies of a package, and so on, so
//! that the leaves number ten to the power of the depth.
//!
//! Everything an embedded package costs is therefore charged to the one
//! document that holds it, through a [`Session`] that lives as long as that
//! document's own embedded parts are being resolved:
//!
//! * **depth** — a package embedded this many levels deep is not opened
//!   ([`crate::limits::max_embedded_depth`]);
//! * **count** — the packages opened for one document are capped
//!   ([`crate::limits::max_embedded_packages`]), and a blob reached again by
//!   another entry name is recognised by its content and not opened again;
//! * **bytes** — the blobs' sizes, and what opening them inflates, share one
//!   allowance (a quarter of [`crate::limits::max_package_bytes`]), and an embedded
//!   package's own package-byte limit is what is left of it;
//! * **text** — the text the embedded packages yield is charged to the text
//!   budget ([`crate::limits::max_text_chars`]); while a package is open, that
//!   limit reads as what is left of the allowance.
//!
//! The state is thread-local and travels with the nested parse to the thread
//! that runs it (see `Document::from_reader_nested`), which never waits for a
//! parse slot: a nested parse that waited for a slot held by the parse that
//! is waiting for it would wait forever.

use std::any::Any;
use std::cell::RefCell;
use std::collections::HashMap;
use std::hash::BuildHasher;
use std::sync::Arc;

/// What is known of a document's embedded packages while it is parsed.
#[derive(Default)]
pub(crate) struct State {
    /// Levels of embedding the running parse is inside.
    depth: usize,
    /// Packages opened so far (not counting those answered from `cache`).
    packages: usize,
    /// Bytes of blobs opened and bytes their packages inflated.
    bytes: u64,
    /// Characters of text the embedded packages have yielded.
    text: usize,
    /// The packages not opened, with the reason, kept up to
    /// [`MAX_REFUSALS_KEPT`]; `refused_more` counts the rest.
    refused: Vec<(String, Refusal)>,
    refused_more: usize,
    /// What each distinct blob yielded, by its content.
    cache: HashMap<Fingerprint, Arc<dyn Any + Send + Sync>>,
}

/// Most refused packages kept to be reported; the rest are counted.
const MAX_REFUSALS_KEPT: usize = 100;

thread_local! {
    static STATE: RefCell<Option<State>> = const { RefCell::new(None) };
}

/// Run `f` on the state, if there is one.
fn with<T>(f: impl FnOnce(&mut State) -> T) -> Option<T> {
    STATE
        .try_with(|s| s.borrow_mut().as_mut().map(f))
        .ok()
        .flatten()
}

/// Take the state off this thread, to carry it to the thread that runs a
/// nested parse.
pub(crate) fn take_state() -> Option<State> {
    STATE.try_with(|s| s.borrow_mut().take()).ok().flatten()
}

/// Put a state taken with [`take_state`] on this thread.
pub(crate) fn install_state(state: Option<State>) {
    let _ = STATE.try_with(|s| *s.borrow_mut() = state);
}

/// A document's embedded-package account, alive while the guard is. The
/// outermost document's session owns the account and ends it on drop; the
/// session of a document opened inside it joins the same account.
#[must_use = "the account ends when the session is dropped"]
pub(crate) struct Session {
    owner: bool,
}

impl Session {
    /// Start the account for the document being parsed, or join the one
    /// of the document that embeds it.
    pub(crate) fn begin() -> Self {
        let owner = STATE
            .try_with(|s| {
                let mut s = s.borrow_mut();
                if s.is_none() {
                    *s = Some(State::default());
                    true
                } else {
                    false
                }
            })
            .unwrap_or(false);
        Self { owner }
    }

    /// Report each package that was not opened, anywhere beneath the
    /// document, through `skip` (name, reason). Only the session that
    /// started the account reports, once; a document embedded in another
    /// leaves its refusals to the document that holds it, as it is gone by
    /// the time anyone reads its report.
    pub(crate) fn report(&mut self, skip: &dyn Fn(&str, &dyn std::fmt::Display)) {
        if !self.owner {
            return;
        }
        let Some((refused, more)) =
            with(|s| (std::mem::take(&mut s.refused), std::mem::take(&mut s.refused_more)))
        else {
            return;
        };
        for (name, refusal) in &refused {
            skip(name, refusal);
        }
        if more > 0 {
            skip("embedded packages", &format!("{more} more embedded packages were not opened"));
        }
    }

    /// Leave the account if the session started it.
    fn end(&mut self) {
        if std::mem::take(&mut self.owner) {
            let _ = take_state();
        }
    }
}

impl Drop for Session {
    fn drop(&mut self) {
        self.end();
    }
}

/// Why an embedded package was not opened.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum Refusal {
    /// Embedded more levels deep than the limit.
    Depth(usize),
    /// More packages than the limit were opened for the document.
    Count(usize),
    /// The embedded packages together are past the byte allowance.
    Bytes(u64),
    /// The embedded packages together yield more text than the budget.
    Text(usize),
}

impl std::fmt::Display for Refusal {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "skipped embedded package: ")?;
        match self {
            Self::Depth(n) => write!(f, "embedded more than {n} levels deep"),
            Self::Count(n) => write!(f, "the document embeds more than {n} packages"),
            Self::Bytes(n) => write!(f, "the embedded packages together exceed {n} bytes"),
            Self::Text(n) => {
                write!(f, "the embedded packages together hold more than {n} characters of text")
            },
        }
    }
}

/// Record that the package `part` was not opened, for `refusal`, to be
/// reported by the document that started the account.
pub(crate) fn refuse(_session: &Session, part: &str, refusal: Refusal) {
    with(|s| {
        if s.refused.len() < MAX_REFUSALS_KEPT {
            let name = if s.depth == 0 {
                part.to_string()
            } else {
                format!("{part} (embedded {} levels deep)", s.depth)
            };
            s.refused.push((name, refusal));
        } else {
            s.refused_more += 1;
        }
    });
}

/// An embedded package being opened: one more level of depth until dropped.
#[must_use = "the level ends when the admission is dropped"]
pub(crate) struct Admission {
    _private: (),
}

impl Drop for Admission {
    fn drop(&mut self) {
        with(|s| s.depth = s.depth.saturating_sub(1));
    }
}

/// Admit the opening of an embedded package of `blob_len` bytes, or say why
/// not. The package counts, with its size, against the document's account
/// until the end of its session.
pub(crate) fn admit(_session: &Session, blob_len: u64) -> Result<Admission, Refusal> {
    with(|s| {
        let max_depth = crate::limits::max_embedded_depth();
        if s.depth >= max_depth {
            return Err(Refusal::Depth(max_depth));
        }
        let max_packages = crate::limits::max_embedded_packages();
        if s.packages >= max_packages {
            return Err(Refusal::Count(max_packages));
        }
        let max_bytes = crate::limits::max_embedded_bytes();
        let bytes = s.bytes.saturating_add(blob_len);
        if bytes > max_bytes {
            return Err(Refusal::Bytes(max_bytes));
        }
        s.depth += 1;
        s.packages += 1;
        s.bytes = bytes;
        Ok(Admission { _private: () })
    })
    // No account: nothing is opened that is not counted.
    .unwrap_or(Err(Refusal::Count(0)))
}

/// Charge `chars` of text an embedded package yielded to the document's
/// text budget.
pub(crate) fn charge_text(_session: &Session, chars: usize) -> Result<(), Refusal> {
    let max = crate::limits::configured_max_text_chars();
    with(|s| {
        let total = s.text.saturating_add(chars);
        s.text = total;
        if total > max {
            Err(Refusal::Text(max))
        } else {
            Ok(())
        }
    })
    .unwrap_or(Err(Refusal::Text(0)))
}

/// The bytes an embedded package's own package-byte limit may be, or `None`
/// outside an embedded package.
pub(crate) fn package_bytes_cap() -> Option<u64> {
    with(|s| (s.depth > 0).then(|| crate::limits::max_embedded_bytes().saturating_sub(s.bytes)))
        .flatten()
}

/// The characters of text an embedded package may yield, or `None` outside
/// an embedded package.
pub(crate) fn text_cap() -> Option<usize> {
    with(|s| {
        (s.depth > 0).then(|| crate::limits::configured_max_text_chars().saturating_sub(s.text))
    })
    .flatten()
}

/// Charge `bytes` an embedded package's zip account inflated, when the
/// package being dropped is one.
pub(crate) fn charge_inflated(bytes: u64) {
    with(|s| {
        if s.depth > 0 {
            s.bytes = s.bytes.saturating_add(bytes);
        }
    });
}

/// A blob identified by its content: its length and two hashes keyed
/// randomly for the process, so that no input can be made to collide with
/// another's.
#[derive(Clone, Copy, PartialEq, Eq, Hash, Debug)]
pub(crate) struct Fingerprint(u64, u64, u64);

impl Fingerprint {
    pub(crate) fn of(blob: &[u8]) -> Self {
        static KEYS: std::sync::OnceLock<(std::hash::RandomState, std::hash::RandomState)> =
            std::sync::OnceLock::new();
        let (a, b) =
            KEYS.get_or_init(|| (std::hash::RandomState::new(), std::hash::RandomState::new()));
        Self(blob.len() as u64, a.hash_one(blob), b.hash_one(blob))
    }
}

/// What the blob with this fingerprint yielded when it was opened for this
/// document, if it was.
pub(crate) fn cached<T: Clone + 'static>(_session: &Session, fp: Fingerprint) -> Option<T> {
    with(|s| {
        s.cache
            .get(&fp)
            .and_then(|v| v.downcast_ref::<T>().cloned())
    })
    .flatten()
}

/// Remember what the blob with this fingerprint yielded.
pub(crate) fn remember<T: Send + Sync + 'static>(_session: &Session, fp: Fingerprint, value: T) {
    with(|s| {
        s.cache.insert(fp, Arc::new(value));
    });
}

/// What stands for an embedded document or package whose text was already
/// given, when it is found again under another name: one line, in place of a
/// second copy of the text.
pub(crate) const REPEATED_NOTICE: &str =
    "[embedded package repeated: its text is given where it first appears]";

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_depth_count_and_bytes_are_admitted_up_to_their_limits() {
        let session = Session::begin();
        let a = admit(&session, 1).unwrap();
        let b = admit(&session, 1).unwrap();
        let c = admit(&session, 1).unwrap();
        assert_eq!(admit(&session, 1).err(), Some(Refusal::Depth(3)));
        drop((c, b, a));
        // Depth is back, the three packages stay counted.
        assert!(admit(&session, 1).is_ok());
        assert_eq!(package_bytes_cap(), None);
    }

    #[test]
    fn test_a_joined_session_does_not_end_the_account() {
        let outer = Session::begin();
        let _a = admit(&outer, 10).unwrap();
        {
            let inner = Session::begin();
            drop(inner);
        }
        assert!(package_bytes_cap().is_some());
        drop(outer);
        assert_eq!(package_bytes_cap(), None);
        assert_eq!(text_cap(), None);
    }

    #[test]
    fn test_fingerprints_tell_blobs_apart_by_content_and_length() {
        assert_eq!(Fingerprint::of(b"abc"), Fingerprint::of(b"abc"));
        assert_ne!(Fingerprint::of(b"abc"), Fingerprint::of(b"abd"));
        assert_ne!(Fingerprint::of(b"abc"), Fingerprint::of(b"abc\0"));
    }
}
