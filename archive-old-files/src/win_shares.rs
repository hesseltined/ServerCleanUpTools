//! Published disk shares (Win32_Share Type 0, no trailing $).

#[derive(Clone, Debug)]
pub struct DiskShare {
    pub name: String,
    pub local_path: String,
    pub unc: String,
}

#[cfg(windows)]
pub fn published_disk_shares() -> Vec<DiskShare> {
    use serde::Deserialize;
    use wmi::{COMLibrary, WMIConnection};

    #[derive(Deserialize, Debug)]
    #[allow(non_snake_case)]
    struct Win32Share {
        Name: String,
        Path: String,
        Type: u32,
    }

    let Ok(com) = COMLibrary::new() else {
        eprintln!("Warning: COM init failed for share list.");
        return vec![];
    };
    let Ok(wmi) = WMIConnection::new(com) else {
        eprintln!("Warning: WMI connection failed for share list.");
        return vec![];
    };
    let Ok(mut rows): Result<Vec<Win32Share>, _> =
        wmi.raw_query("SELECT Name, Path, Type FROM Win32_Share")
    else {
        eprintln!("Warning: Win32_Share query failed.");
        return vec![];
    };

    rows.sort_by(|a, b| a.Name.cmp(&b.Name));
    let hostname = std::env::var("COMPUTERNAME").unwrap_or_default();

    let mut out = Vec::new();
    for s in rows {
        if s.Type != 0 {
            continue;
        }
        if s.Name.contains('$') {
            continue;
        }
        if s.Path.trim().is_empty() {
            continue;
        }
        let local = s.Path.trim_end_matches('\\').to_string();
        let unc = format!(r"\\{}\{}", hostname, s.Name);
        out.push(DiskShare {
            name: s.Name,
            local_path: local,
            unc,
        });
    }
    out
}

#[cfg(windows)]
pub fn local_path_for_share_name(share_name: &str) -> Option<String> {
    let name = share_name.trim();
    if name.is_empty() {
        return None;
    }
    for s in published_disk_shares() {
        if s.name.eq_ignore_ascii_case(name) {
            return Some(s.local_path.clone());
        }
    }
    None
}

#[cfg(not(windows))]
pub fn published_disk_shares() -> Vec<DiskShare> {
    vec![]
}

#[cfg(not(windows))]
pub fn local_path_for_share_name(_share_name: &str) -> Option<String> {
    None
}
