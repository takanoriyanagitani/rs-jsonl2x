use clap::Parser;
use std::io;

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Cli {
    /// Output Excel file path
    #[arg(short, long)]
    output: String,

    /// Sheet name in the Excel file
    #[arg(short, long, default_value = "Sheet1")]
    sheet: String,

    /// Optional: Limit the number of rows to process (max 65536)
    #[arg(short, long, default_value_t = 65536)]
    row_limit: u32,
}

fn main() {
    let cli = Cli::parse();

    if cli.row_limit > 65536 {
        eprintln!("Error: Row limit cannot exceed 65536.");
        std::process::exit(1);
    }

    let stdin = io::stdin();

    if let Err(err) = rs_jsonl2x::run(stdin.lock(), &cli.output, &cli.sheet, cli.row_limit) {
        eprintln!("Error: {}", err);
        std::process::exit(1);
    }
}
