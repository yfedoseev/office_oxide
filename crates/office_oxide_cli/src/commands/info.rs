use office_oxide::Document;

pub fn run(file: &str) -> Result<(), Box<dyn std::error::Error>> {
    let doc = Document::open(file)?;
    let ir = doc.to_ir();

    let meta = &ir.metadata;
    println!("Format: {:?}", meta.format);
    for (label, value) in meta.properties() {
        println!("{label}: {value}");
    }
    for p in &meta.custom_properties {
        println!("Custom property: {} = {}", p.name, p.value);
    }
    if meta.has_macros {
        println!("Macros: yes");
    }
    if meta.has_digital_signature {
        println!("Digitally signed: yes");
    }
    if meta.thumbnail.is_some() {
        println!("Thumbnail: yes");
    }
    if ir.metadata.text_truncated {
        println!(
            "Warning: text extraction is incomplete — the source file's own structure disagrees \
             with itself about how much text there is, and the gap could not be safely recovered"
        );
    }
    println!("Sections: {}", ir.sections.len());

    for (i, section) in ir.sections.iter().enumerate() {
        let title = section.title.as_deref().unwrap_or("(untitled)");
        println!("  [{i}] {title} — {} elements", section.elements.len());
    }

    Ok(())
}
