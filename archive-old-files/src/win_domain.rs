//! Domain / computer name helpers (Windows WMI).

#[cfg(windows)]
pub fn computer_system_domain() -> Option<String> {
    use serde::Deserialize;
    use wmi::{COMLibrary, WMIConnection};

    #[derive(Deserialize, Debug)]
    #[allow(non_snake_case)]
    struct Win32ComputerSystem {
        Domain: Option<String>,
    }

    let com = COMLibrary::new().ok()?;
    let wmi = WMIConnection::new(com).ok()?;
    let q: Vec<Win32ComputerSystem> = wmi
        .raw_query("SELECT Domain FROM Win32_ComputerSystem")
        .unwrap_or_default();
    q.into_iter().next()?.Domain.filter(|d| !d.is_empty())
}

#[cfg(not(windows))]
pub fn computer_system_domain() -> Option<String> {
    None
}
