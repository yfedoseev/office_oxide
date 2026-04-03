//! Validate a document file by parsing and extracting text/markdown/IR.
//! Outputs a JSON line with results for each file argument.
//!
//! Usage: validate FILE [FILE ...]

use std::time::Instant;

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    if args.is_empty() {
        eprintln!("Usage: validate FILE|DIR [FILE|DIR ...]");
        std::process::exit(1);
    }

    // Collect all files, expanding directories recursively
    let mut files = Vec::new();
    for arg in &args {
        let path = std::path::Path::new(arg);
        if path.is_dir() {
            collect_files(path, &mut files);
        } else {
            files.push(arg.clone());
        }
    }

    let total = files.len();
    let mut ok = 0usize;
    let mut fail = 0usize;
    let wall_start = Instant::now();

    // Print JSON array start
    println!("[");
    for (i, path) in files.iter().enumerate() {
        let result = validate_file(path);
        if result.contains("\"parse_ok\": true") {
            ok += 1;
        } else {
            fail += 1;
        }
        let comma = if i + 1 < total { "," } else { "" };
        println!("{result}{comma}");
    }
    println!("]");

    let wall_secs = wall_start.elapsed().as_secs_f64();
    let pct = if total > 0 {
        ok as f64 / total as f64 * 100.0
    } else {
        0.0
    };
    eprintln!(
        "\n=== Validation Summary ===\nTotal: {total}  OK: {ok}  FAIL: {fail}  Rate: {pct:.1}%  Wall: {wall_secs:.1}s"
    );
}

fn collect_files(dir: &std::path::Path, out: &mut Vec<String>) {
    let Ok(entries) = std::fs::read_dir(dir) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            collect_files(&path, out);
        } else if let Some(ext) = path.extension().and_then(|e| e.to_str()) {
            match ext.to_ascii_lowercase().as_str() {
                "docx" | "xlsx" | "pptx" | "doc" | "xls" | "ppt" => {
                    out.push(path.to_string_lossy().into_owned());
                }
                _ => {}
            }
        }
    }
}

fn validate_file(path: &str) -> String {
    let start = Instant::now();

    let doc = match office_oxide::Document::open(path) {
        Ok(d) => d,
        Err(e) => {
            let elapsed = start.elapsed().as_millis();
            let err_msg = format!("{e}").replace('\\', "\\\\").replace('"', "\\\"");
            return format!(
                r#"  {{"path": "{path}", "parse_ok": false, "error": "{err_msg}", "parse_time_ms": {elapsed}}}"#,
            );
        }
    };
    let parse_ms = start.elapsed().as_millis();

    let plain_text_len = doc.plain_text().len();
    let markdown_len = doc.to_markdown().len();
    let ir = doc.to_ir();
    let section_count = ir.sections.len();

    let total_ms = start.elapsed().as_millis();
    let format_name = format!("{:?}", doc.format());

    format!(
        r#"  {{"path": "{path}", "parse_ok": true, "format": "{format_name}", "parse_time_ms": {parse_ms}, "total_time_ms": {total_ms}, "plain_text_len": {plain_text_len}, "markdown_len": {markdown_len}, "sections": {section_count}}}"#,
    )
}
