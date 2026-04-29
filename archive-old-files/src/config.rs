//! JSON config compatible with Archive-OldFiles.config.json from PowerShell.

use crate::age_basis::{normalize_age_basis_param, AgeBasis};
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "PascalCase")]
pub struct EmailConfig {
    #[serde(default)]
    pub enabled: bool,
    #[serde(default)]
    pub to: String,
    #[serde(default)]
    pub from: String,
    #[serde(default)]
    pub profile: String,
    #[serde(default)]
    pub smtp_host: String,
    #[serde(default = "default_smtp_port")]
    pub smtp_port: u16,
    #[serde(default = "default_true")]
    pub use_ssl: bool,
    #[serde(default)]
    pub user_name: String,
    #[serde(default)]
    pub password_protected_base64: Option<String>,
}

fn default_smtp_port() -> u16 {
    587
}

fn default_true() -> bool {
    true
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "PascalCase")]
pub struct AppConfig {
    #[serde(default)]
    pub schema_version: Option<i32>,
    #[serde(default)]
    pub script_version: Option<String>,
    #[serde(default)]
    pub saved_at: Option<String>,
    #[serde(default)]
    pub input_path: String,
    #[serde(default)]
    pub archive_path: String,
    #[serde(default)]
    pub archive_share_name: Option<String>,
    #[serde(default)]
    pub years: Option<f64>,
    #[serde(default)]
    pub age_basis: Option<String>,
    #[serde(default)]
    pub commit: Option<bool>,
    #[serde(default)]
    pub output: Option<String>,
    #[serde(default)]
    pub remove_empty_folders: Option<String>,
    #[serde(default)]
    pub all_shares: Option<bool>,
    #[serde(default)]
    pub email: Option<EmailConfig>,
}

pub fn normalize_output_format(raw: Option<&str>) -> Option<String> {
    let s = raw?.trim();
    let lower = s.to_ascii_lowercase();
    match lower.as_str() {
        "text" => Some("Text".to_string()),
        "html" | "csv" => Some("HTML".to_string()),
        _ => None,
    }
}

pub fn normalize_smtp_profile(raw: &str) -> String {
    let s = raw.trim();
    let lower = s.to_ascii_lowercase();
    if lower.contains("gmail") {
        return "GmailAppPassword".to_string();
    }
    if lower.contains("microsoft365")
        || lower.contains("office365")
        || lower.contains("m365")
        || lower.contains("outlook")
    {
        return "Microsoft365AppPassword".to_string();
    }
    if lower.contains("smtp2go") {
        return "Smtp2Go".to_string();
    }
    "Custom".to_string()
}

pub fn apply_smtp_preset(profile: &str, host: &mut String, port: &mut u16, use_ssl: &mut bool) {
    match profile {
        "GmailAppPassword" => {
            *host = "smtp.gmail.com".to_string();
            *port = 587;
            *use_ssl = true;
        }
        "Microsoft365AppPassword" => {
            *host = "smtp.office365.com".to_string();
            *port = 587;
            *use_ssl = true;
        }
        "Smtp2Go" => {
            *host = "mail.smtp2go.com".to_string();
            *port = 2525;
            *use_ssl = true;
        }
        _ => {}
    }
}

/// Runtime email settings (hashtable equivalent).
#[derive(Debug, Clone)]
pub struct EmailRuntime {
    pub enabled: bool,
    pub to: String,
    pub from: String,
    pub profile: String,
    pub smtp_host: String,
    pub smtp_port: u16,
    pub use_ssl: bool,
    pub user_name: String,
    pub password_protected_base64: Option<String>,
}

impl Default for EmailRuntime {
    fn default() -> Self {
        Self {
            enabled: false,
            to: String::new(),
            from: String::new(),
            profile: "Custom".to_string(),
            smtp_host: String::new(),
            smtp_port: 587,
            use_ssl: true,
            user_name: String::new(),
            password_protected_base64: None,
        }
    }
}

pub fn email_from_json(node: Option<&EmailConfig>) -> EmailRuntime {
    let mut e = EmailRuntime::default();
    let Some(n) = node else {
        return e;
    };
    e.enabled = n.enabled;
    if !n.to.trim().is_empty() {
        e.to = n.to.trim().to_string();
    }
    if !n.from.trim().is_empty() {
        e.from = n.from.trim().to_string();
    }
    if !n.profile.trim().is_empty() {
        e.profile = normalize_smtp_profile(&n.profile);
    }
    if !n.smtp_host.trim().is_empty() {
        e.smtp_host = n.smtp_host.trim().to_string();
    }
    if n.smtp_port > 0 {
        e.smtp_port = n.smtp_port;
    }
    e.use_ssl = n.use_ssl;
    if !n.user_name.trim().is_empty() {
        e.user_name = n.user_name.trim().to_string();
    }
    e.password_protected_base64 = n
        .password_protected_base64
        .as_ref()
        .filter(|s| !s.trim().is_empty())
        .map(|s| s.trim().to_string());

    if e.profile != "Custom" && e.smtp_host.is_empty() {
        apply_smtp_preset(&e.profile, &mut e.smtp_host, &mut e.smtp_port, &mut e.use_ssl);
    }
    if e.profile == "Smtp2Go"
        && e.smtp_port == 587
        && e.smtp_host.eq_ignore_ascii_case("mail.smtp2go.com")
    {
        e.smtp_port = 2525;
    }
    e
}

pub fn parse_age_basis_saved(raw: Option<&str>) -> Option<AgeBasis> {
    raw.and_then(|s| normalize_age_basis_param(s))
}

pub fn read_config(path: &std::path::Path) -> anyhow::Result<Option<AppConfig>> {
    if !path.exists() {
        return Ok(None);
    }
    let raw = std::fs::read_to_string(path)?;
    match serde_json::from_str::<AppConfig>(&raw) {
        Ok(c) => Ok(Some(c)),
        Err(err) => {
            eprintln!("Warning: could not read config file (ignored): {}", err);
            Ok(None)
        }
    }
}

pub fn save_config(
    path: &std::path::Path,
    version: &str,
    input_path: &str,
    archive_path: &str,
    archive_share_name: &str,
    years: f64,
    age_basis: AgeBasis,
    commit: bool,
    output: &str,
    remove_empty: Option<&str>,
    all_shares: bool,
    email: &EmailRuntime,
) -> anyhow::Result<()> {
    let now = chrono::Local::now().to_rfc3339();
    let payload = serde_json::json!({
        "SchemaVersion": 2,
        "ScriptVersion": version,
        "SavedAt": now,
        "InputPath": input_path,
        "ArchivePath": archive_path,
        "ArchiveShareName": archive_share_name,
        "Years": years,
        "AgeBasis": age_basis.as_str(),
        "Commit": commit,
        "Output": output,
        "RemoveEmptyFolders": remove_empty,
        "AllShares": all_shares,
        "Email": {
            "Enabled": email.enabled,
            "To": email.to,
            "From": email.from,
            "Profile": email.profile,
            "SmtpHost": email.smtp_host,
            "SmtpPort": email.smtp_port,
            "UseSsl": email.use_ssl,
            "UserName": email.user_name,
            "PasswordProtectedBase64": email.password_protected_base64,
        }
    });
    let s = serde_json::to_string_pretty(&payload)?;
    std::fs::write(path, s)?;
    Ok(())
}
