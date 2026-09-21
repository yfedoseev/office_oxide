//! Scratch probe: dump plain text for a document, to check chart text.
fn main() {
    for p in std::env::args().skip(1) {
        println!("=== {p}");
        let doc = office_oxide::Document::open(&p).unwrap();
        println!("--- plain_text\n{}", doc.plain_text());
        println!("--- ir text\n{}", doc.to_ir().plain_text());
    }
}
