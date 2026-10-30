//! 09_concurrency: open and read 32 documents from 8 threads with logging on.
//!
//! Parsing runs on a helper thread that may log. A logger must therefore never
//! block on a lock the parsing caller holds while it waits for the parse. The
//! logger here only takes its own short-lived lock, and the workers hold no
//! lock while they call into the library.
//!
//! Run: `cargo run --example 09_concurrency`

use std::io::Cursor;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::{Arc, Mutex};
use std::thread;

use log::{Level, LevelFilter, Log, Metadata, Record};
use office_oxide::create::create_from_markdown_to_writer;
use office_oxide::{Document, DocumentFormat};

const MARKDOWN: &str = "# Quarterly Report\n\nRevenue grew in every region.\n\n\
| Region | Q1  | Q2  |\n|--------|-----|-----|\n| North  | 120 | 135 |\n| South  | 95  | 110 |\n";

fn build_docx_bytes(markdown: &str) -> Vec<u8> {
    let mut buf = Cursor::new(Vec::new());
    create_from_markdown_to_writer(markdown, DocumentFormat::Docx, &mut buf).expect("create DOCX");
    buf.into_inner()
}

/// Assert with a message; a failure exits non-zero.
fn check(condition: bool, message: &str) {
    if !condition {
        eprintln!("FAIL: {message}");
        std::process::exit(1);
    }
}

static LOG_LINES: AtomicUsize = AtomicUsize::new(0);

/// Counts lines and keeps the last one under a private lock.
struct CountingLogger(Mutex<String>);

impl Log for CountingLogger {
    fn enabled(&self, metadata: &Metadata) -> bool {
        metadata.level() <= Level::Trace
    }
    fn log(&self, record: &Record) {
        LOG_LINES.fetch_add(1, Ordering::Relaxed);
        *self.0.lock().unwrap() = format!("{}: {}", record.target(), record.args());
    }
    fn flush(&self) {}
}

static LOGGER: CountingLogger = CountingLogger(Mutex::new(String::new()));

fn main() {
    println!("step 1: switch logging on");
    log::set_logger(&LOGGER).expect("install logger");
    log::set_max_level(LevelFilter::Trace);

    println!("step 2: build 32 DOCX documents in memory");
    let documents: Arc<Vec<Vec<u8>>> = Arc::new(
        (0..32)
            .map(|i| {
                build_docx_bytes(
                    &MARKDOWN.replace("Quarterly Report", &format!("Quarterly Report {i}")),
                )
            })
            .collect(),
    );

    println!("step 3: read them from 8 workers");
    let workers: Vec<_> = (0..8)
        .map(|worker| {
            let documents = Arc::clone(&documents);
            thread::spawn(move || {
                (worker..32)
                    .step_by(8)
                    .map(|i| {
                        let doc = Document::from_reader(
                            Cursor::new(documents[i].clone()),
                            DocumentFormat::Docx,
                        )
                        .expect("open DOCX");
                        (i, doc.plain_text())
                    })
                    .collect::<Vec<_>>()
            })
        })
        .collect();
    let mut read = 0;
    for handle in workers {
        for (i, text) in handle.join().expect("a worker panicked") {
            check(
                text.contains(&format!("Quarterly Report {i}")),
                "each document must read back its own title",
            );
            read += 1;
        }
    }
    println!("  workers: 8");
    println!("  documents: {read}");

    println!("step 4: check the results");
    check(read == 32, "every document must be read");
    println!("OK 09_concurrency");
}
