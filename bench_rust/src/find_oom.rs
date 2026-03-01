// Standalone binary to test calamine on a single file
use calamine::{open_workbook_auto, Data, Reader};
use std::path::Path;

fn main() {
    let path = std::env::args().nth(1).expect("path required");
    let path = Path::new(&path);

    let mut wb = match open_workbook_auto(path) {
        Ok(wb) => wb,
        Err(e) => {
            eprintln!("OPEN_ERR: {}", e);
            std::process::exit(1);
        }
    };

    let names: Vec<String> = wb.sheet_names().to_vec();
    for name in names {
        match wb.worksheet_range(&name) {
            Ok(range) => {
                let (rows, cols) = range.get_size();
                eprintln!("Sheet '{}': {}x{}", name, rows, cols);
            }
            Err(e) => {
                eprintln!("RANGE_ERR on '{}': {}", name, e);
            }
        }
    }
    eprintln!("OK");
}
