//! Command-line interface (maps to PowerShell parameters).

use clap::Parser;
use std::path::PathBuf;

pub const VERSION: &str = env!("CARGO_PKG_VERSION");

#[derive(Parser, Debug)]
#[command(name = "archive-old-files")]
#[command(
    version = VERSION,
    about = "Age-based file archival — Rust port of Archive-OldFiles.ps1 (Windows)."
)]
pub struct Cli {
    /// Root folder to scan (InputPath)
    #[arg(long, value_name = "PATH")]
    pub input: Option<String>,

    /// Archive root (ArchivePath)
    #[arg(long, value_name = "PATH")]
    pub archive: Option<String>,

    /// Age threshold in years (decimals allowed)
    #[arg(long, value_name = "N")]
    pub years: Option<String>,

    /// Perform moves (commit)
    #[arg(long)]
    pub commit: bool,

    /// Text or HTML
    #[arg(long, value_name = "FORMAT")]
    pub output: Option<String>,

    /// Yes or No (after commit)
    #[arg(long, value_name = "Yes|No")]
    pub remove_empty_folders: Option<String>,

    /// Config JSON path (default: next to executable)
    #[arg(long, value_name = "PATH")]
    pub config: Option<PathBuf>,

    #[arg(long)]
    pub no_save_config: bool,

    /// Default SMB share name used to resolve archive path
    #[arg(long, default_value = "Archive")]
    pub archive_share_name: String,

    #[arg(long)]
    pub skip_share_menu: bool,

    /// LastWriteTime, LastAccessTime, LatestWriteOrAccess, CreationTime, Earliest (aliases like PS)
    #[arg(long)]
    pub age_basis: Option<String>,

    #[arg(long)]
    pub no_run_log: bool,

    /// Preview all published disk shares (HTML per share; no moves)
    #[arg(long)]
    pub all: bool,

    #[arg(long)]
    pub email_to: Option<String>,

    #[arg(long)]
    pub email_from: Option<String>,

    #[arg(long)]
    pub smtp_profile: Option<String>,

    #[arg(long)]
    pub smtp_host: Option<String>,

    #[arg(long)]
    pub smtp_port: Option<u16>,

    #[arg(long)]
    pub smtp_use_ssl: Option<String>,

    #[arg(long)]
    pub smtp_user: Option<String>,

    #[arg(long)]
    pub no_send_email: bool,

    /// Send test SMTP using saved config (no scan)
    #[arg(long)]
    pub test_email: bool,
}
