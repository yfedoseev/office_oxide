pub fn run(file: &str, embed_images: bool) -> Result<(), Box<dyn std::error::Error>> {
    let doc = office_oxide::Document::open(file)?;
    if embed_images {
        use office_oxide::ir_render::{ImageEmbed, MarkdownOptions};
        print!(
            "{}",
            doc.to_markdown_with(MarkdownOptions {
                image_embed: ImageEmbed::Base64,
            })
        );
    } else {
        print!("{}", doc.to_markdown());
    }
    Ok(())
}
