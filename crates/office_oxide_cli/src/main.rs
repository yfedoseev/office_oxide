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
#[command(
    name = "office-oxide",
    version,
    about = "Fast Office document processing"
)]
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

#[cfg(test)]
mod tests {
    use super::*;
    use clap::CommandFactory;

    /// `--version` must be wired up: the CLI shipped without it, so
    /// `office-oxide --version` failed with clap's generic
    /// "unexpected argument" error.
    #[test]
    fn test_cli_has_version_flag() {
        let cmd = Cli::command();
        assert_eq!(cmd.get_version(), Some(env!("CARGO_PKG_VERSION")));
        let err = match Cli::try_parse_from(["office-oxide", "--version"]) {
            Err(e) => e,
            Ok(_) => panic!("--version should short-circuit parsing"),
        };
        assert_eq!(err.kind(), clap::error::ErrorKind::DisplayVersion);
        assert!(err.to_string().contains(env!("CARGO_PKG_VERSION")));
    }
}
