//! Path helpers: normalization, relative paths, archive-under-input check.

use std::path::{Path, PathBuf};

/// Trim quotes and whitespace (matches PowerShell Normalize-UserPath).
pub fn normalize_user_path(s: &str) -> String {
    let mut p = s.trim().to_string();
    if p.len() >= 2 && p.starts_with('"') && p.ends_with('"') {
        p = p[1..p.len() - 1].trim().to_string();
    }
    if p.len() >= 2 && p.starts_with('\'') && p.ends_with('\'') {
        p = p[1..p.len() - 1].trim().to_string();
    }
    p
}

/// Strip trailing backslashes/slashes for display and comparisons.
pub fn trim_sep(p: &str) -> String {
    let mut s = p.to_string();
    while s.ends_with('\\') || s.ends_with('/') {
        s.pop();
    }
    s
}

#[cfg(windows)]
fn path_starts_with_ignore_case(full: &str, root: &str) -> bool {
    full.len() >= root.len() && full[..root.len()].eq_ignore_ascii_case(&root[..])
}

#[cfg(not(windows))]
fn path_starts_with_ignore_case(full: &str, root: &str) -> bool {
    full.starts_with(root)
}

/// True if archive is the same as or nested under input (Windows: case-insensitive).
pub fn archive_under_input(input_resolved: &str, archive_resolved: &str) -> bool {
    let input = trim_sep(input_resolved);
    let archive = trim_sep(archive_resolved);
    if archive.eq_ignore_ascii_case(&input) {
        return true;
    }
    let prefix = format!("{}\\", input);
    path_starts_with_ignore_case(&archive, &prefix)
}

/// Relative path from root to file; both should be canonicalized when possible.
pub fn relative_path_from_root(file_full: &Path, root_full: &Path) -> anyhow::Result<PathBuf> {
    let file_s = file_full.to_string_lossy();
    let root_s = root_full.to_string_lossy();
    let root_t = trim_sep(&root_s);
    let file_t = trim_sep(&file_s);
    if !file_t
        .to_lowercase()
        .starts_with(&root_t.to_lowercase())
    {
        anyhow::bail!("File is not under input root: {}", file_full.display());
    }
    if file_t.len() < root_t.len() {
        anyhow::bail!("File is not under input root: {}", file_full.display());
    }
    let tail = &file_t[root_t.len()..];
    let tail = tail.trim_start_matches(['\\', '/']);
    Ok(PathBuf::from(tail))
}

pub fn unique_destination_path(initial: &Path) -> PathBuf {
    if !initial.exists() {
        return initial.to_path_buf();
    }
    let parent = initial.parent().unwrap_or_else(|| Path::new("."));
    let file_name = initial.file_name().and_then(|n| n.to_str()).unwrap_or("file");
    let (base, ext) = if let Some(dot) = file_name.rfind('.') {
        if dot == 0 {
            (file_name, "")
        } else {
            (&file_name[..dot], &file_name[dot..])
        }
    } else {
        (file_name, "")
    };
    let mut n = 2u32;
    loop {
        let candidate = parent.join(format!("{}_{}{}", base, n, ext));
        if !candidate.exists() {
            return candidate;
        }
        n += 1;
    }
}

/// Try rename; on Windows cross-volume copy+delete.
pub fn move_file_best_effort(src: &Path, dst: &Path) -> std::io::Result<()> {
    match std::fs::rename(src, dst) {
        Ok(()) => Ok(()),
        Err(e) => {
            #[cfg(windows)]
            {
                use std::io::ErrorKind;
                // ERROR_NOT_SAME_DEVICE = 17
                if e.raw_os_error() == Some(17) {
                    std::fs::copy(src, dst)?;
                    std::fs::remove_file(src)?;
                    return Ok(());
                }
            }
            Err(e)
        }
    }
}
