pub fn run(file: &str) -> Result<(), Box<dyn std::error::Error>> {
    let doc = super::open(file)?;
    super::emit_with(|out| doc.write_html(out))?;
    Ok(())
}
