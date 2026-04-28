pub fn run(file: &str) -> Result<(), Box<dyn std::error::Error>> {
    let doc = office_oxide::Document::open(file)?;
    print!("{}", doc.to_markdown());
    Ok(())
}
