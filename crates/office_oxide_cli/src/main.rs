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
    /// Keep text formatted as hidden (left out by default)
    #[arg(long, global = true)]
    include_hidden: bool,
    /// Keep text a tracked change deleted (left out by default)
    #[arg(long, global = true)]
    include_deleted: bool,
    /// For input nobody vouches for: tighter limits, and a document that
    /// could not be read completely is an error
    #[arg(long, global = true)]
    strict: bool,
    /// Set an option for this run: a limit such as `text_chars=1000000` or
    /// `part_bytes=1048576` (`max` for none), or `include_hidden_text=1`;
    /// repeatable
    #[arg(long = "limit", value_name = "NAME=VALUE", global = true)]
    limits: Vec<String>,
    #[command(subcommand)]
    command: commands::Command,
}

/// The options this run reads documents under, from the flags.
fn parse_options(cli: &Cli) -> Result<office_oxide::options::ParseOptions, String> {
    let mut options = if cli.strict {
        office_oxide::options::ParseOptions::strict()
    } else {
        office_oxide::options::ParseOptions::default()
    };
    if cli.include_hidden {
        options.include_hidden_text = true;
    }
    if cli.include_deleted {
        options.include_deleted_text = true;
    }
    for setting in &cli.limits {
        let (name, value) = setting
            .split_once('=')
            .ok_or_else(|| format!("--limit expects NAME=VALUE, not '{setting}'"))?;
        let value = if matches!(value, "max" | "none" | "unlimited") {
            u64::MAX
        } else {
            value
                .parse::<u64>()
                .map_err(|_| format!("--limit {name}: '{value}' is not a number or 'max'"))?
        };
        options.set_option(name, value).map_err(|e| e.to_string())?;
    }
    Ok(options)
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
    match parse_options(&cli) {
        Ok(options) => commands::set_options(options),
        Err(e) => {
            eprintln!("error: {e}");
            process::exit(2);
        },
    }
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
