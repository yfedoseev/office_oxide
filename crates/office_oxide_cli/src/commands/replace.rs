pub fn run(
    file: &str,
    find: &str,
    replace: &str,
    output: Option<&str>,
) -> Result<(), Box<dyn std::error::Error>> {
    // An empty search string is never what the caller meant: `str::replace`
    // treats it as "every position", so the replacement is interleaved
    // between every character of every run and the command still reports
    // success. Reject it before anything is opened or written.
    if find.is_empty() {
        return Err("search string cannot be empty".into());
    }
    let mut doc = office_oxide::edit::EditableDocument::open(file)?;
    // Fail before writing: an unsupported format used to report
    // "0 occurrences" and rewrite the file anyway.
    let count = doc.replace_text(find, replace)?;
    let out = output.unwrap_or(file);
    doc.save(out)?;
    eprintln!("replaced {count} occurrence(s); wrote {out}");
    Ok(())
}

#[cfg(test)]
mod tests {
    /// An empty `find` used to interleave the replacement between every
    /// character of the document and report "replaced N occurrence(s)".
    /// The guard must fire before the file is even opened, so a path that
    /// does not exist still yields the empty-search error.
    #[test]
    fn test_replace_rejects_empty_find() {
        let err = super::run("/nonexistent/does-not-exist.docx", "", "X", None)
            .expect_err("empty find must be rejected");
        assert!(
            err.to_string().contains("search string cannot be empty"),
            "unexpected error: {err}"
        );
    }
}
