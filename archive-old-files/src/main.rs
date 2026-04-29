//! Purpose: CLI entry for age-based file archival (Rust port of Archive-OldFiles.ps1).
//! Author: Doug Hesseltine
//! Created: 2026-03-28
//! Modified: 2026-03-28
//! Version: See Cargo.toml (`CARGO_PKG_VERSION` exposed as `cli::VERSION`).

mod age_basis;
mod cli;
mod config;
mod email;
mod format_util;
mod html_report;
mod job;
mod path_util;
mod run_log;
mod runner;
mod win_domain;
mod win_dpapi;
mod win_owner;
mod win_shares;

use clap::Parser;

fn main() {
    let cli = cli::Cli::parse();
    if let Err(e) = entry(cli) {
        eprintln!("Error: {:#}", e);
        std::process::exit(1);
    }
}

fn entry(cli: cli::Cli) -> anyhow::Result<()> {
    #[cfg(not(windows))]
    {
        let _ = cli;
        anyhow::bail!(
            "archive-old-files runs on Windows file servers.\n\
             Cross-compile from macOS/Linux:\n\
               rustup target add x86_64-pc-windows-msvc\n\
               cargo build --release --target x86_64-pc-windows-msvc"
        );
    }
    #[cfg(windows)]
    {
        runner::run(&cli)
    }
}
