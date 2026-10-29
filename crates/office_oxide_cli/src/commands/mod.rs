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

/// Write a rendering to stdout as the document produces it, so a large
/// spreadsheet is never held whole in memory.
///
/// `print!` panics on any write error other than a closed pipe (which
/// `main` turns back into a quiet `SIGPIPE` exit): `office-oxide text f >
/// /dev/full` exited 101 with a panic message. Propagating the error gives
/// the usual `error: …` and exit status 1 instead.
fn emit_with(
    render: impl FnOnce(&mut dyn std::io::Write) -> std::io::Result<()>,
) -> std::io::Result<()> {
    use std::io::Write;
    // Unbuffered, every row would be a write to the terminal or pipe.
    let mut out = std::io::BufWriter::with_capacity(1 << 16, std::io::stdout().lock());
    render(&mut out)?;
    out.flush()
}

static OPTIONS: std::sync::OnceLock<office_oxide::options::ParseOptions> =
    std::sync::OnceLock::new();

/// The options every document of this run is read under (set once, by `main`).
pub fn set_options(options: office_oxide::options::ParseOptions) {
    let _ = OPTIONS.set(options);
}

/// Open `file` under this run's options.
pub fn open(file: &str) -> office_oxide::Result<office_oxide::Document> {
    match OPTIONS.get() {
        Some(options) => office_oxide::Document::open_with(file, options),
        None => office_oxide::Document::open(file),
    }
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
