//! Open every file given on the command line and run the full pipeline:
//! `open` -> `to_ir` -> `plain_text` -> `to_markdown` -> `to_html`, and for
//! DOCX/XLSX/PPTX also the write round trip: `to_ir` -> `create_from_ir_to_writer`
//! -> re-`open` (`from_reader`). Prints one JSON line per *step* so a hard
//! crash (abort / stack overflow / SIGSEGV — which kills the whole process)
//! can still be pinned to the exact file and step it happened on.
//!
//! Usage: `cargo run --release --example corpus_validate -- FILE...`
//!        `cargo run --release --example corpus_validate -- --list files.txt`
//!
//! The format is taken from the extension *family* (`.docm`/`.dotx` -> DOCX,
//! `.xlsm`/`.xlsb`/`.xltx` -> XLSX, `.pps`/`.pot` -> PPT, ...) so variant
//! extensions that `Document::open` does not map still get exercised.
//!
//! Output, one line per step:
//! ```text
//! {"path":...,"step":"open|to_ir|plain_text|to_markdown|to_html|rt_write|rt_reopen","status":"start"}
//! {"path":...,"step":...,"status":"ok|error|panic","error":...,"ms":...}
//! ```
//! `Document::open`/`from_reader` already run inside a guarded thread
//! (`with_parse_stack`), so a panic during `open` surfaces as `status:"panic"`
//! via the normal `Err` path. The rendering calls (`to_ir`, `plain_text`,
//! `to_markdown`, `to_html`) and the write round trip do NOT go through that
//! guard, so they are wrapped here in `catch_unwind` explicitly. A default
//! panic hook is installed that swallows the backtrace print (10k files times
//! a panic each would flood stderr); the message is still recovered from the
//! `catch_unwind` payload.

use office_oxide::create::create_from_ir_to_writer;
use office_oxide::format::DocumentFormat;
use office_oxide::{Document, OfficeError};
use std::io::{Cursor, Write};
use std::panic::{self, AssertUnwindSafe};
use std::path::Path;
use std::time::Instant;

/// Words (three or more letters) lost from `before` and gained in `after`,
/// as multisets, plus `before`'s total — the same measure the release
/// sweep's `compare_tree.py` uses.
/// (lost, gained, total, sample of lost words) — word multisets of the
/// original IR's text against the reread's.
fn word_diff(before: &str, after: &str) -> (usize, usize, usize, Vec<String>) {
    use std::collections::HashMap;
    fn words(s: &str) -> HashMap<String, i64> {
        let mut m = HashMap::new();
        for w in s
            .split(|c: char| !c.is_alphabetic())
            .filter(|w| w.chars().count() >= 3)
        {
            *m.entry(w.to_lowercase()).or_insert(0) += 1;
        }
        m
    }
    let b = words(before);
    let a = words(after);
    let total: i64 = b.values().sum();
    let lost: i64 = b
        .iter()
        .map(|(w, n)| (n - a.get(w).copied().unwrap_or(0)).max(0))
        .sum();
    let gained: i64 = a
        .iter()
        .map(|(w, n)| (n - b.get(w).copied().unwrap_or(0)).max(0))
        .sum();
    let mut sample: Vec<(String, i64)> = b
        .iter()
        .filter_map(|(w, n)| {
            let d = n - a.get(w).copied().unwrap_or(0);
            (d > 0).then(|| (w.clone(), d))
        })
        .collect();
    sample.sort_by(|x, y| y.1.cmp(&x.1).then_with(|| x.0.cmp(&y.0)));
    sample.truncate(10);
    (
        lost as usize,
        gained as usize,
        total as usize,
        sample.into_iter().map(|(w, _)| w).collect(),
    )
}

fn family(path: &Path) -> Option<DocumentFormat> {
    let ext = path.extension()?.to_str()?.to_ascii_lowercase();
    Some(match ext.as_str() {
        "doc" | "dot" | "wbk" => DocumentFormat::Doc,
        "xls" | "xlt" | "xla" | "xlw" => DocumentFormat::Xls,
        "ppt" | "pps" | "pot" => DocumentFormat::Ppt,
        "docx" | "docm" | "dotx" | "dotm" => DocumentFormat::Docx,
        "xlsx" | "xlsm" | "xltx" | "xltm" | "xlsb" | "xlam" => DocumentFormat::Xlsx,
        "pptx" | "pptm" | "potx" | "potm" | "ppsx" | "ppsm" => DocumentFormat::Pptx,
        _ => return None,
    })
}

fn esc(s: &str) -> String {
    serde_json::to_string(s).unwrap_or_else(|_| "\"?\"".into())
}

fn panic_msg(payload: Box<dyn std::any::Any + Send>) -> String {
    payload
        .downcast_ref::<&str>()
        .map(|s| (*s).to_string())
        .or_else(|| payload.downcast_ref::<String>().cloned())
        .unwrap_or_else(|| "unknown panic payload".to_string())
}

/// Run `step` for `path`, printing a `"start"` line first (flushed) and a
/// result line after. Returns the value on success so later steps can chain.
fn step<T>(
    out: &mut impl Write,
    path: &str,
    name: &str,
    f: impl FnOnce() -> T + std::panic::UnwindSafe,
) -> Option<T> {
    let _ = writeln!(out, "{{\"path\":{},\"step\":\"{name}\",\"status\":\"start\"}}", esc(path));
    let _ = out.flush();
    let start = Instant::now();
    match panic::catch_unwind(f) {
        Ok(v) => {
            let _ = writeln!(
                out,
                "{{\"path\":{},\"step\":\"{name}\",\"status\":\"ok\",\"ms\":{}}}",
                esc(path),
                start.elapsed().as_millis()
            );
            Some(v)
        },
        Err(payload) => {
            let msg = panic_msg(payload);
            let _ = writeln!(
                out,
                "{{\"path\":{},\"step\":\"{name}\",\"status\":\"panic\",\"error\":{},\"ms\":{}}}",
                esc(path),
                esc(&msg),
                start.elapsed().as_millis()
            );
            None
        },
    }
}

fn main() {
    // Silence the default panic backtrace print: with 10k+ files and dozens
    // of panicking ones expected, the raw hook output would dominate stderr.
    // The message itself is still recovered from the catch_unwind payload.
    panic::set_hook(Box::new(|_info| {}));

    let mut args: Vec<String> = std::env::args().skip(1).collect();
    if args.first().map(String::as_str) == Some("--list") {
        let list = args.get(1).expect("--list FILE");
        let text = std::fs::read_to_string(list).expect("read list");
        args = text
            .lines()
            .filter(|l| !l.trim().is_empty())
            .map(str::to_owned)
            .collect();
    }
    let stdout = std::io::stdout();

    for a in args {
        let path = Path::new(&a);
        let mut out = stdout.lock();
        let _ = writeln!(out, "{{\"path\":{},\"status\":\"running\"}}", esc(&a));
        let _ = out.flush();

        let Some(fmt) = family(path) else {
            let _ = writeln!(
                out,
                "{{\"path\":{},\"step\":\"open\",\"status\":\"error\",\"error\":\"unsupported extension\",\"ms\":0}}",
                esc(&a)
            );
            continue;
        };

        // --- open ---
        let open_start = Instant::now();
        let _ = writeln!(out, "{{\"path\":{},\"step\":\"open\",\"status\":\"start\"}}", esc(&a));
        let _ = out.flush();
        let doc = match Document::open(path) {
            Ok(d) => {
                let _ = writeln!(
                    out,
                    "{{\"path\":{},\"step\":\"open\",\"status\":\"ok\",\"ms\":{}}}",
                    esc(&a),
                    open_start.elapsed().as_millis()
                );
                d
            },
            Err(e) => {
                let status = if matches!(e, OfficeError::Panic(_)) {
                    "panic"
                } else {
                    "error"
                };
                let _ = writeln!(
                    out,
                    "{{\"path\":{},\"step\":\"open\",\"status\":\"{status}\",\"error\":{},\"ms\":{}}}",
                    esc(&a),
                    esc(&e.to_string()),
                    open_start.elapsed().as_millis()
                );
                continue;
            },
        };

        // --- to_ir (kept for later steps; also independently re-derived by
        // to_markdown/to_html internally, but we call it once here as its
        // own step so a to_ir-only panic is distinguishable) ---
        let ir = step(&mut out, &a, "to_ir", AssertUnwindSafe(|| doc.to_ir()));

        let _ = step(&mut out, &a, "plain_text", AssertUnwindSafe(|| doc.plain_text()));
        let _ = step(&mut out, &a, "to_markdown", AssertUnwindSafe(|| doc.to_markdown()));
        let _ = step(&mut out, &a, "to_html", AssertUnwindSafe(|| doc.to_html()));

        // --- round trip (DOCX/XLSX/PPTX only): to_ir -> write -> re-open ---
        if matches!(fmt, DocumentFormat::Docx | DocumentFormat::Xlsx | DocumentFormat::Pptx) {
            if let Some(ir) = ir.as_ref() {
                let bytes = step(
                    &mut out,
                    &a,
                    "rt_write",
                    AssertUnwindSafe(|| {
                        let mut buf = Cursor::new(Vec::new());
                        create_from_ir_to_writer(ir, fmt, &mut buf).map(|()| buf.into_inner())
                    }),
                );
                if let Some(Ok(bytes)) = bytes {
                    let reopened = step(
                        &mut out,
                        &a,
                        "rt_reopen",
                        AssertUnwindSafe(|| Document::from_reader(Cursor::new(bytes), fmt)),
                    );
                    // --- fidelity: the words of the original IR against the
                    // reread's. A write path that drops or duplicates content
                    // succeeds at both steps above; only this sees it.
                    if let Some(Ok(reopened)) = reopened {
                        let diff = step(
                            &mut out,
                            &a,
                            "rt_compare",
                            AssertUnwindSafe(|| {
                                word_diff(&ir.plain_text(), &reopened.to_ir().plain_text())
                            }),
                        );
                        if let Some((lost, gained, total, sample)) = diff {
                            let sample =
                                serde_json::to_string(&sample).unwrap_or_else(|_| "[]".into());
                            let _ = writeln!(
                                out,
                                "{{\"path\":{},\"step\":\"rt_words\",\"status\":\"ok\",\"lost\":{lost},\"gained\":{gained},\"total\":{total},\"lost_sample\":{sample}}}",
                                esc(&a)
                            );
                        }
                    }
                } else if let Some(Err(e)) = bytes {
                    let _ = writeln!(
                        out,
                        "{{\"path\":{},\"step\":\"rt_write\",\"status\":\"error\",\"error\":{},\"ms\":0}}",
                        esc(&a),
                        esc(&e.to_string())
                    );
                }
            }
        }

        let _ = writeln!(out, "{{\"path\":{},\"step\":\"done\",\"status\":\"ok\"}}", esc(&a));
        let _ = out.flush();
    }
}
