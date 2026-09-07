pub fn run(
    file: &str,
    find: &str,
    replace: &str,
    output: Option<&str>,
) -> Result<(), Box<dyn std::error::Error>> {
    let mut doc = office_oxide::edit::EditableDocument::open(file)?;
    let count = doc.replace_text(find, replace);
    let out = output.unwrap_or(file);
    doc.save(out)?;
    eprintln!("replaced {count} occurrence(s); wrote {out}");
    Ok(())
}
