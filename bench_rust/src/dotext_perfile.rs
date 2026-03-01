use std::io::Read;
use std::path::{Path, PathBuf};
use dotext::MsDoc;

fn collect_files(dir: &Path) -> Vec<PathBuf> {
    let mut files = Vec::new();
    collect_recursive(dir, &mut files);
    files.sort();
    files
}

fn collect_recursive(dir: &Path, out: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(dir) else { return };
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

fn try_dotext(path: &Path) -> Result<(), String> {
    let ext = path.extension().and_then(|e| e.to_str()).unwrap_or("").to_ascii_lowercase();
    match ext.as_str() {
        "docx" => {
            let mut f = dotext::Docx::open(path).map_err(|e| format!("{}", e))?;
            let mut t = String::new();
            f.read_to_string(&mut t).map_err(|e| format!("{}", e))?;
            Ok(())
        }
        "xlsx" => {
            let mut f = dotext::Xlsx::open(path).map_err(|e| format!("{}", e))?;
            let mut t = String::new();
            f.read_to_string(&mut t).map_err(|e| format!("{}", e))?;
            Ok(())
        }
        "pptx" => {
            let mut f = dotext::Pptx::open(path).map_err(|e| format!("{}", e))?;
            let mut t = String::new();
            f.read_to_string(&mut t).map_err(|e| format!("{}", e))?;
            Ok(())
        }
        _ => Err("unsupported".to_string()),
    }
}

fn main() {
    let dir = std::env::args().nth(1).expect("DIR required");
    let files = collect_files(Path::new(&dir));
    for path in &files {
        let path_str = path.to_string_lossy();
        let result = std::panic::catch_unwind(|| try_dotext(path));
        match result {
            Ok(Ok(())) => println!("OK\t{}", path_str),
            Ok(Err(e)) => println!("FAIL\t{}\t{}", path_str, e),
            Err(_) => println!("FAIL\t{}\tPANIC", path_str),
        }
    }
}
