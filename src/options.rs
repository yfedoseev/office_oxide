//! Per-call options: the limits, which text is kept, and what to do when the
//! result is known to be incomplete.
//!
//! The `set_max_*` functions in [`limits`](crate::limits) and the
//! `set_include_*` functions in [`read_options`](crate::read_options) change
//! the process-wide defaults. [`ParseOptions`] belongs to one call instead:
//! [`Document::open_with`](crate::Document::open_with) keeps it with the
//! document, and every conversion of that document (text, Markdown, HTML, IR)
//! runs under it, on whichever thread it is called. Nothing global changes,
//! so documents for different tenants, or tests, can run side by side.

use crate::limits::Limits;

/// What to do when a limit cut the result or the file was damaged, so that
/// the content returned is known to be incomplete.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[non_exhaustive]
pub enum OnIncomplete {
    /// Return what was read and say what is missing: in the text, in
    /// `Metadata::warnings` and `Metadata::text_truncated`, and in
    /// [`Document::damage`](crate::Document::damage). The default, and what
    /// the process-wide functions do.
    #[default]
    Report,
    /// Refuse the document: [`Document::open_with`](crate::Document::open_with)
    /// returns an error naming the first losses. For callers who would
    /// rather have nothing than a partial result (indexing, compliance).
    /// Costs one extra conversion of the document at open.
    Error,
}

/// Options for one document: see the [module documentation](self).
///
/// Build it with [`ParseOptions::default`] (the process-wide settings now in
/// force) or [`ParseOptions::strict`], then change fields.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub struct ParseOptions {
    /// The bounds on memory, output and nesting.
    pub limits: Limits,
    /// Keep text formatted as hidden (see
    /// [`read_options::include_hidden_text`](crate::read_options::include_hidden_text)).
    pub include_hidden_text: bool,
    /// Keep text a tracked change deleted (see
    /// [`read_options::include_deleted_text`](crate::read_options::include_deleted_text)).
    pub include_deleted_text: bool,
    /// What to do when the result is known to be incomplete.
    pub on_incomplete: OnIncomplete,
}

impl Default for ParseOptions {
    /// The process-wide limits and read settings in force now, reporting
    /// incomplete results.
    fn default() -> Self {
        Self {
            limits: Limits::current(),
            include_hidden_text: crate::read_options::include_hidden_text(),
            include_deleted_text: crate::read_options::include_deleted_text(),
            on_incomplete: OnIncomplete::Report,
        }
    }
}

impl ParseOptions {
    /// For input from a source nobody vouches for: [`Limits::strict`], and
    /// an incomplete result is an error.
    pub fn strict() -> Self {
        Self {
            limits: Limits::strict(),
            on_incomplete: OnIncomplete::Error,
            ..Self::default()
        }
    }

    /// Set an option by name, for callers that cannot name a Rust field (the
    /// command line, the C interface, the language bindings).
    ///
    /// `name` is a limit (see [`Limits::NAMES`]) with a number as its value, or
    /// one of `include_hidden_text` and `include_deleted_text` (0 or 1), or
    /// `on_incomplete` (0 reports, 1 is an error).
    pub fn set_option(&mut self, name: &str, value: u64) -> Result<(), UnknownOption> {
        match name {
            "include_hidden_text" => self.include_hidden_text = value != 0,
            "include_deleted_text" => self.include_deleted_text = value != 0,
            "on_incomplete" => {
                self.on_incomplete = if value == 0 {
                    OnIncomplete::Report
                } else {
                    OnIncomplete::Error
                }
            },
            limit => self
                .limits
                .set(limit, value)
                .map_err(|_| UnknownOption(name.to_string()))?,
        }
        Ok(())
    }

    pub(crate) fn read_settings(&self) -> (bool, bool) {
        (self.include_hidden_text, self.include_deleted_text)
    }
}

/// An option name [`ParseOptions::set_option`] does not know.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UnknownOption(pub String);

impl std::fmt::Display for UnknownOption {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "unknown option '{}'; the options are: include_hidden_text, include_deleted_text, \
             on_incomplete, {}",
            self.0,
            Limits::NAMES.join(", ")
        )
    }
}

impl std::error::Error for UnknownOption {}

/// How a document came to be incomplete.
#[derive(Debug, Clone, Copy, PartialEq, Eq, serde::Serialize)]
#[serde(rename_all = "snake_case")]
#[non_exhaustive]
pub enum DamageKind {
    /// A part of the file could not be read and its content is missing (an
    /// unsupported compression method, an encrypted entry, a part that is
    /// not well-formed at all).
    UnreadablePart,
    /// Output was cut: a limit was reached, or the file's own structure
    /// shows that text is missing.
    Truncated,
    /// Something was repaired or guessed and the content is probably intact
    /// (a checksum mismatch, a header field out of range, a code page
    /// inferred). Informational; it does not make the result incomplete.
    Recovered,
}

/// One thing wrong with a document, found while reading it.
#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize)]
#[non_exhaustive]
pub struct Damage {
    /// What kind of loss it is.
    pub kind: DamageKind,
    /// The part (or stream) concerned, when the reader knows it.
    pub part: Option<String>,
    /// What happened, as the reader reports it.
    pub detail: String,
}

impl Damage {
    /// Whether content is missing: every kind but [`DamageKind::Recovered`].
    pub fn is_loss(&self) -> bool {
        self.kind != DamageKind::Recovered
    }
}

impl std::fmt::Display for Damage {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match &self.part {
            Some(part) => write!(f, "{part}: {}", self.detail),
            None => f.write_str(&self.detail),
        }
    }
}
