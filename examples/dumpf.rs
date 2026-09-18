fn main() {
    for p in std::env::args().skip(1) {
        let doc = office_oxide::xlsx::XlsxDocument::open(&p).unwrap();
        println!("== {p}");
        for ws in &doc.worksheets {
            for row in &ws.rows {
                for c in &row.cells {
                    if let Some(f) = &c.formula {
                        println!("  {} = {}", c.reference, f);
                    }
                }
            }
        }
    }
}
