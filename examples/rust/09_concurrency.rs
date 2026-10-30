//! 09_concurrency: open and read 32 documents from 8 workers at once, with logging on.
//!
//! Run: `cargo run --example 09_concurrency`

use log::{LevelFilter, Log, Metadata, Record};
use office_oxide::create::create_from_markdown_to_writer;
use office_oxide::{Document, DocumentFormat};
use std::io::Cursor;
use std::sync::{Arc, mpsc};
use std::thread;
use std::time::Duration;

const MARKDOWN: &str = "# Quarterly Report\n\nRevenue grew in every region.\n\n\
| Region | Q1  | Q2  |\n|--------|-----|-----|\n| North  | 120 | 135 |\n| South  | 95  | 110 |\n";

fn markdown_to_bytes(markdown: &str, format: DocumentFormat) -> Vec<u8> {
    let mut buf = Cursor::new(Vec::new());
    create_from_markdown_to_writer(markdown, format, &mut buf).expect("create from Markdown");
    buf.into_inner()
}

fn build_docx_bytes(title: &str) -> Vec<u8> {
    markdown_to_bytes(&MARKDOWN.replace("Quarterly Report", title), DocumentFormat::Docx)
}

const WORKERS: usize = 8;
const DOCUMENTS: usize = 32;

/// Parsing runs on a helper thread that may log, so a logger must never block
/// on a lock the parsing caller holds. This one only takes stderr's own lock
/// for one line.
struct StderrLogger;

impl Log for StderrLogger {
    fn enabled(&self, _: &Metadata) -> bool {
        true
    }
    fn log(&self, record: &Record) {
        eprintln!("{} {}", thread::current().name().unwrap_or("worker"), record.args());
    }
    fn flush(&self) {}
}

static LOGGER: StderrLogger = StderrLogger;

fn check(condition: bool, message: &str) {
    if !condition {
        eprintln!("FAIL: {message}");
        std::process::exit(1);
    }
}

fn main() {
    println!("step 1: build 32 DOCX files in memory");
    let inputs: Arc<Vec<Vec<u8>>> = Arc::new(
        (0..DOCUMENTS)
            .map(|i| build_docx_bytes(&format!("Quarterly Report {i}")))
            .collect(),
    );
    println!("  documents: {DOCUMENTS}");

    println!("step 2: switch logging on");
    log::set_logger(&LOGGER).expect("install the logger");
    log::set_max_level(LevelFilter::Debug);

    println!("step 3: parse them on 8 workers");
    let (sender, receiver) = mpsc::channel();
    for worker in 0..WORKERS {
        let (inputs, sender) = (Arc::clone(&inputs), sender.clone());
        thread::spawn(move || {
            for i in (worker..DOCUMENTS).step_by(WORKERS) {
                log::debug!("document {i}: parse start");
                let doc =
                    Document::from_reader(Cursor::new(inputs[i].clone()), DocumentFormat::Docx)
                        .expect("open DOCX");
                let text = doc.plain_text();
                log::debug!("document {i}: parse done");
                sender.send((i, text)).expect("the main thread is waiting");
            }
        });
    }
    drop(sender);
    let mut texts = vec![String::new(); DOCUMENTS];
    for _ in 0..DOCUMENTS {
        // A deadlock fails the example instead of hanging it.
        let (i, text) = receiver
            .recv_timeout(Duration::from_secs(60))
            .unwrap_or_else(|_| {
                eprintln!("FAIL: a parse did not finish: possible deadlock");
                std::process::exit(1)
            });
        texts[i] = text;
    }
    println!("  workers: {WORKERS}");

    println!("step 4: check every result");
    check(texts.iter().all(|t| !t.is_empty()), "not every document was read");
    for (i, text) in texts.iter().enumerate() {
        check(
            text.contains(&format!("Quarterly Report {i}")),
            &format!("document {i} came back with the wrong text"),
        );
    }

    println!("OK 09_concurrency");
}
