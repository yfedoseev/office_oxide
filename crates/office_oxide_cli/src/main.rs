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
    // Rust ignores SIGPIPE at startup, so `office-oxide text f.docx | head`
    // panicked with "failed printing to stdout: Broken pipe" once `head`
    // closed the pipe. Restore the default disposition, as ripgrep and fd
    // do: a closed stdout ends the process quietly, like every other
    // text tool.
    #[cfg(unix)]
    // SAFETY: `signal` with SIG_DFL only resets a disposition; it is
    // called once, before any other thread exists.
    unsafe {
        libc::signal(libc::SIGPIPE, libc::SIG_DFL);
    }
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

    /// The root README's CLI section documents every subcommand and every
    /// flag the binary has — it once omitted `replace` and
    /// `--embed-images`. Reading the list off clap keeps the two from
    /// drifting apart again.
    #[test]
    fn test_readme_cli_section_lists_every_subcommand_and_flag() {
        let readme =
            std::fs::read_to_string(concat!(env!("CARGO_MANIFEST_DIR"), "/../../README.md"))
                .unwrap();
        let start = readme.find("\n## CLI\n").expect("README has a CLI section");
        let end = readme[start + 1..]
            .find("\n## ")
            .map_or(readme.len(), |i| start + 1 + i);
        let section = &readme[start..end];
        for sub in Cli::command().get_subcommands() {
            let name = sub.get_name();
            assert!(
                section.contains(&format!("office-oxide {name} ")),
                "README's CLI section does not show `office-oxide {name}`"
            );
            for arg in sub.get_arguments() {
                if let Some(long) = arg.get_long() {
                    if long == "help" || long == "version" {
                        continue;
                    }
                    assert!(
                        section.contains(&format!("--{long}")),
                        "README's CLI section does not mention `office-oxide {name} --{long}`"
                    );
                }
            }
        }
    }

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
