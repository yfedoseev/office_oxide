//! Writing a rendering into a [`std::io::Write`] as it is produced.
//!
//! The renderers produce their output through [`std::fmt::Write`], so the
//! same code fills a `String` and streams into a file or a socket. This
//! adapts one to the other, keeping the I/O error that `fmt::Error` cannot
//! carry.

use std::fmt;
use std::io;

/// A [`fmt::Write`] over an [`io::Write`].
pub(crate) struct IoSink<'a, W: io::Write + ?Sized> {
    inner: &'a mut W,
    error: Option<io::Error>,
}

impl<W: io::Write + ?Sized> fmt::Write for IoSink<'_, W> {
    fn write_str(&mut self, s: &str) -> fmt::Result {
        self.inner.write_all(s.as_bytes()).map_err(|e| {
            self.error = Some(e);
            fmt::Error
        })
    }
}

/// Run `render` against a [`fmt::Write`] that writes to `out`, returning the
/// I/O error that stopped it, if any.
pub(crate) fn render_to_io<W: io::Write + ?Sized>(
    out: &mut W,
    render: impl FnOnce(&mut IoSink<'_, W>) -> fmt::Result,
) -> io::Result<()> {
    let mut sink = IoSink {
        inner: out,
        error: None,
    };
    match render(&mut sink) {
        Ok(()) => Ok(()),
        Err(_) => Err(sink
            .error
            .take()
            .unwrap_or_else(|| io::Error::other("rendering failed"))),
    }
}

/// Writes `prefix` before the first non-empty string written through it, and
/// nothing if none is: the separator in front of a piece that may turn out
/// to be empty.
pub(crate) struct Prefixed<'a, W: fmt::Write + ?Sized> {
    out: &'a mut W,
    prefix: &'a str,
    pending: bool,
}

impl<'a, W: fmt::Write + ?Sized> Prefixed<'a, W> {
    pub(crate) fn new(out: &'a mut W, prefix: &'a str) -> Self {
        Self {
            out,
            prefix,
            pending: true,
        }
    }
}

impl<W: fmt::Write + ?Sized> fmt::Write for Prefixed<'_, W> {
    fn write_str(&mut self, s: &str) -> fmt::Result {
        if s.is_empty() {
            return Ok(());
        }
        if self.pending {
            self.pending = false;
            self.out.write_str(self.prefix)?;
        }
        self.out.write_str(s)
    }
}
