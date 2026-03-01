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
