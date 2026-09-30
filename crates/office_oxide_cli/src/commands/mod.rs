mod html;
mod info;
mod ir;
mod markdown;
mod replace;
mod text;

use clap::Subcommand;

#[derive(Subcommand)]
pub enum Command {
    /// Extract plain text from a document
    Text {
        /// Path to the document file
        file: String,
    },
    /// Convert a document to markdown
    Markdown {
        /// Path to the document file
        file: String,
        /// Embed each image inline as `[image-base64:<data>]` at its
        /// position in the document flow. Images are otherwise dropped
        /// from markdown entirely.
        #[arg(long)]
        embed_images: bool,
    },
    /// Convert a document to HTML
    Html {
        /// Path to the document file
        file: String,
    },
    /// Show document metadata
    Info {
        /// Path to the document file
        file: String,
    },
    /// Dump the document IR as JSON
    Ir {
        /// Path to the document file
        file: String,
    },
    /// Replace text in a document, preserving every other part of the file
    Replace {
        /// Path to the document file
        file: String,
        /// Text to search for
        find: String,
        /// Replacement text
        replace: String,
        /// Where to write the result (default: overwrite the input file)
        #[arg(short, long)]
        output: Option<String>,
    },
}

/// Write a rendered result to stdout.
///
/// `print!` panics on any write error other than a closed pipe (which
/// `main` turns back into a quiet `SIGPIPE` exit): `office-oxide text f >
/// /dev/full` exited 101 with a panic message. Propagating the error gives
/// the usual `error: …` and exit status 1 instead.
fn emit(s: &str) -> std::io::Result<()> {
    use std::io::Write;
    let mut out = std::io::stdout().lock();
    out.write_all(s.as_bytes())?;
    out.flush()
}

pub fn run(cmd: Command) -> Result<(), Box<dyn std::error::Error>> {
    match cmd {
        Command::Text { file } => text::run(&file),
        Command::Markdown { file, embed_images } => markdown::run(&file, embed_images),
        Command::Html { file } => html::run(&file),
        Command::Info { file } => info::run(&file),
        Command::Ir { file } => ir::run(&file),
        Command::Replace {
            file,
            find,
            replace,
            output,
        } => replace::run(&file, &find, &replace, output.as_deref()),
    }
}
