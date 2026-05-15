//! `office-oxide` — command-line front-end to the `office_oxide` library.
//!
//! Extracts text, converts to Markdown / HTML / IR, and inspects DOCX,
//! XLSX, PPTX, DOC, XLS, and PPT files. See `office-oxide --help` for
//! the full subcommand list.

#![warn(missing_docs)]

mod commands;

use clap::Parser;
use std::process;

#[derive(Parser)]
#[command(name = "office-oxide", about = "Fast Office document processing")]
struct Cli {
    #[command(subcommand)]
    command: commands::Command,
}

fn main() {
    let cli = Cli::parse();
    if let Err(e) = commands::run(cli.command) {
        eprintln!("error: {e}");
        process::exit(1);
    }
}
