use std::collections::HashMap;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::time::Instant;

use dotext::MsDoc;

fn collect_files(dir: &Path) -> Vec<PathBuf> {
    let mut files = Vec::new();
    collect_recursive(dir, &mut files);
    files.sort();
    files
}

fn collect_recursive(dir: &Path, out: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(dir) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            collect_recursive(&path, out);
        } else if let Some(ext) = path.extension().and_then(|e| e.to_str()) {
            match ext.to_ascii_lowercase().as_str() {
                "docx" | "xlsx" | "pptx" => out.push(path),
                _ => {}
            }
        }
    }
}

struct Stats {
    ok: usize,
    fail: usize,
    total_ms: f64,
    errors: HashMap<String, usize>,
}

impl Stats {
    fn new() -> Self {
        Stats { ok: 0, fail: 0, total_ms: 0.0, errors: HashMap::new() }
    }
    fn record_ok(&mut self, ms: f64) { self.ok += 1; self.total_ms += ms; }
    fn record_err(&mut self, ms: f64, e: String) {
        self.fail += 1;
        self.total_ms += ms;
        let key = if e.len() > 80 { format!("{}...", &e[..80]) } else { e };
        *self.errors.entry(key).or_insert(0) += 1;
    }
}

fn try_calamine(path: &Path) -> Result<String, String> {
    use calamine::{open_workbook_auto, Data, Reader};
    // Skip known OOM file (calamine allocates 8TB dense matrix for sparse sheet)
    if path.file_name().is_some_and(|n| n == "lo_too-many-cols-rows.xlsx") {
        return Err("SKIP: known OOM file".to_string());
    }
    let mut wb = open_workbook_auto(path).map_err(|e| format!("{}", e))?;
    let mut parts = Vec::new();
    let names: Vec<String> = wb.sheet_names().to_vec();
    for name in names {
        if let Ok(range) = wb.worksheet_range(&name) {
            for row in range.rows() {
                let cells: Vec<String> = row.iter().map(|c| match c {
                    Data::String(s) => s.clone(),
                    Data::Float(f) => f.to_string(),
                    Data::Int(i) => i.to_string(),
                    Data::Bool(b) => b.to_string(),
                    Data::Empty => String::new(),
                    _ => format!("{:?}", c),
                }).collect();
                parts.push(cells.join("\t"));
            }
        }
    }
    Ok(parts.join("\n"))
}

fn try_dotext_docx(path: &Path) -> Result<String, String> {
    let mut file = dotext::Docx::open(path).map_err(|e| format!("{}", e))?;
    let mut text = String::new();
    file.read_to_string(&mut text).map_err(|e| format!("{}", e))?;
    Ok(text)
}

fn try_dotext_xlsx(path: &Path) -> Result<String, String> {
    let mut file = dotext::Xlsx::open(path).map_err(|e| format!("{}", e))?;
    let mut text = String::new();
    file.read_to_string(&mut text).map_err(|e| format!("{}", e))?;
    Ok(text)
}

fn try_dotext_pptx(path: &Path) -> Result<String, String> {
    let mut file = dotext::Pptx::open(path).map_err(|e| format!("{}", e))?;
    let mut text = String::new();
    file.read_to_string(&mut text).map_err(|e| format!("{}", e))?;
    Ok(text)
}

fn try_docx_rs(path: &Path) -> Result<String, String> {
    let data = std::fs::read(path).map_err(|e| format!("{}", e))?;
    let docx = docx_rs::read_docx(&data).map_err(|e| format!("{}", e))?;
    let mut parts = Vec::new();
    for child in &docx.document.children {
        if let docx_rs::DocumentChild::Paragraph(p) = child {
            let mut para_text = String::new();
            for pc in &p.children {
                if let docx_rs::ParagraphChild::Run(r) = pc {
                    for rc in &r.children {
                        if let docx_rs::RunChild::Text(t) = rc {
                            para_text.push_str(&t.text);
                        }
                    }
                }
            }
            parts.push(para_text);
        }
    }
    Ok(parts.join("\n"))
}

type LibFn = fn(&Path) -> Result<String, String>;

fn run_lib(lib_name: &str, ext_filter: &str, func: LibFn, files: &[PathBuf]) -> Stats {
    let mut stats = Stats::new();
    for path in files {
        let ext = path.extension().and_then(|e| e.to_str()).unwrap_or("").to_ascii_lowercase();
        if ext != ext_filter { continue; }

        let t0 = Instant::now();
        let path_clone = path.clone();
        let result = std::panic::catch_unwind(move || func(&path_clone));
        let elapsed = t0.elapsed().as_secs_f64() * 1000.0;

        match result {
            Ok(Ok(_)) => stats.record_ok(elapsed),
            Ok(Err(e)) => stats.record_err(elapsed, e),
            Err(_) => stats.record_err(elapsed, "PANIC".to_string()),
        }
    }
    stats
}

fn print_stats(lib_name: &str, ext: &str, s: &Stats) {
    let total = s.ok + s.fail;
    let pct = if total > 0 { s.ok as f64 / total as f64 * 100.0 } else { 0.0 };
    let wall = s.total_ms / 1000.0;
    println!("{} ({}):", lib_name, ext);
    println!("  Total: {}  OK: {}  FAIL: {}  Rate: {:.1}%  Wall: {:.1}s", total, s.ok, s.fail, pct, wall);
    if !s.errors.is_empty() {
        let mut errs: Vec<(&String, &usize)> = s.errors.iter().collect();
        errs.sort_by(|a, b| b.1.cmp(a.1));
        for (err, count) in errs.iter().take(10) {
            println!("    {}: {}", err, count);
        }
    }
    println!();
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    if args.len() < 2 {
        eprintln!("Usage: bench_rust --lib <calamine|dotext|docx-rs|all> DIR");
        std::process::exit(1);
    }

    let lib_arg = if args[0] == "--lib" { &args[1] } else { &args[0] };
    let dir_arg = args.last().unwrap();
    let dir = Path::new(dir_arg);
    let files = collect_files(dir);
    eprintln!("Found {} files", files.len());

    println!("\n=== Rust Library Benchmark Results ===\n");

    let run_all = lib_arg == "all";

    if run_all || lib_arg == "calamine" {
        let s = run_lib("calamine", "xlsx", try_calamine, &files);
        print_stats("calamine", "xlsx", &s);
    }
    if run_all || lib_arg == "docx-rs" {
        let s = run_lib("docx-rs", "docx", try_docx_rs, &files);
        print_stats("docx-rs", "docx", &s);
    }
    if run_all || lib_arg == "dotext" {
        let s1 = run_lib("dotext", "docx", try_dotext_docx, &files);
        print_stats("dotext", "docx", &s1);
        let s2 = run_lib("dotext", "xlsx", try_dotext_xlsx, &files);
        print_stats("dotext", "xlsx", &s2);
        let s3 = run_lib("dotext", "pptx", try_dotext_pptx, &files);
        print_stats("dotext", "pptx", &s3);
    }
}
