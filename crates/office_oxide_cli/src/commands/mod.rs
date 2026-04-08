mod text;
mod markdown;
mod html;
mod info;
mod ir;

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
}

pub fn run(cmd: Command) -> Result<(), Box<dyn std::error::Error>> {
    match cmd {
        Command::Text { file } => text::run(&file),
        Command::Markdown { file } => markdown::run(&file),
        Command::Html { file } => html::run(&file),
        Command::Info { file } => info::run(&file),
        Command::Ir { file } => ir::run(&file),
    }
}
