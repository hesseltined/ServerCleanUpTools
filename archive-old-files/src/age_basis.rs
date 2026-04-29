//! Age basis parsing and per-file timestamp selection (matches PowerShell script).

use chrono::{DateTime, Local};
use std::fs::Metadata;
use std::time::SystemTime;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AgeBasis {
    LastWriteTime,
    LastAccessTime,
    LatestWriteOrAccess,
    CreationTime,
    Earliest,
}

impl AgeBasis {
    pub fn as_str(&self) -> &'static str {
        match self {
            AgeBasis::LastWriteTime => "LastWriteTime",
            AgeBasis::LastAccessTime => "LastAccessTime",
            AgeBasis::LatestWriteOrAccess => "LatestWriteOrAccess",
            AgeBasis::CreationTime => "CreationTime",
            AgeBasis::Earliest => "Earliest",
        }
    }

    pub fn report_column_header(&self) -> &'static str {
        match self {
            AgeBasis::LastAccessTime => "Last opened (access)",
            AgeBasis::LatestWriteOrAccess => "Age date (max modified / access)",
            AgeBasis::CreationTime => "Created",
            AgeBasis::Earliest => "Older of created / modified",
            AgeBasis::LastWriteTime => "Last modified",
        }
    }
}

pub fn normalize_age_basis_param(raw: &str) -> Option<AgeBasis> {
    let ab = raw.trim();
    let lower = ab.to_ascii_lowercase();
    match lower.as_str() {
        "lastwritetime" | "lastmodified" | "modified" | "a" => Some(AgeBasis::LastWriteTime),
        "lastaccesstime" | "lastaccess" | "lastopened" | "opened" | "b" => {
            Some(AgeBasis::LastAccessTime)
        }
        "latestwriteoraccess" | "modifiedoropened" | "writeoraccess" | "latestactivity" | "c" => {
            Some(AgeBasis::LatestWriteOrAccess)
        }
        "creationtime" => Some(AgeBasis::CreationTime),
        "earliest" => Some(AgeBasis::Earliest),
        _ => None,
    }
}

fn system_time_to_local(st: std::io::Result<SystemTime>) -> DateTime<Local> {
    st.map(|t| DateTime::<Local>::from(t))
        .unwrap_or_else(|_| Local::now())
}

/// Uses metadata times; `accessed` may be unavailable on some platforms (falls back to now).
pub fn file_age_timestamp(meta: &Metadata, basis: AgeBasis) -> DateTime<Local> {
    let modified = system_time_to_local(meta.modified());
    let created = system_time_to_local(meta.created());
    let accessed = system_time_to_local(meta.accessed());

    match basis {
        AgeBasis::CreationTime => created,
        AgeBasis::LastAccessTime => accessed,
        AgeBasis::LastWriteTime => modified,
        AgeBasis::LatestWriteOrAccess => {
            if modified > accessed {
                modified
            } else {
                accessed
            }
        }
        AgeBasis::Earliest => {
            if modified < created {
                modified
            } else {
                created
            }
        }
    }
}

/// Cutoff: files with age timestamp strictly before this are candidates (same as PS: skip if >= cutoff).
pub fn cutoff_from_years(years: f64) -> DateTime<Local> {
    // Match .NET-style fractional years via day approximation (PS passes double to AddYears which truncates to int in .NET;
    // we use continuous days for predictable behavior with decimals.)
    let days = (years * 365.2425).round() as i64;
    Local::now() - chrono::Duration::days(days)
}
