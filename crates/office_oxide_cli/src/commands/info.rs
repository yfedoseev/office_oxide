use office_oxide::Document;

pub fn run(file: &str) -> Result<(), Box<dyn std::error::Error>> {
    let doc = Document::open(file)?;
    let ir = doc.to_ir();

    println!("Format: {:?}", ir.metadata.format);
    if let Some(ref title) = ir.metadata.title {
        println!("Title: {title}");
    }
    println!("Sections: {}", ir.sections.len());

    for (i, section) in ir.sections.iter().enumerate() {
        let title = section.title.as_deref().unwrap_or("(untitled)");
        println!("  [{i}] {title} — {} elements", section.elements.len());
    }

    Ok(())
}
