//! Resolve file owner for HTML report (SID to account name via LookupAccountSidW).

#[cfg(windows)]
pub fn file_owner_for_report(path: &std::path::Path) -> String {
    use std::ffi::OsStr;
    use std::os::windows::ffi::OsStrExt;
    use windows::core::{PCWSTR, PWSTR};
    use windows::Win32::Foundation::LocalFree;
    use windows::Win32::Security::Authorization::{
        GetNamedSecurityInfoW, LookupAccountSidW, OWNER_SECURITY_INFORMATION, SE_FILE_OBJECT,
    };
    use windows::Win32::Security::SID_NAME_USE;

    let wide: Vec<u16> = OsStr::new(path.as_os_str())
        .encode_wide()
        .chain(std::iter::once(0))
        .collect();

    let mut psd: *mut std::ffi::c_void = std::ptr::null_mut();
    let mut owner_sid: *mut std::ffi::c_void = std::ptr::null_mut();

    let r = unsafe {
        GetNamedSecurityInfoW(
            PCWSTR(wide.as_ptr()),
            SE_FILE_OBJECT,
            OWNER_SECURITY_INFORMATION,
            &mut owner_sid,
            std::ptr::null_mut(),
            std::ptr::null_mut(),
            std::ptr::null_mut(),
            &mut psd,
        )
    };

    use windows::Win32::Foundation::ERROR_SUCCESS;
    if r != ERROR_SUCCESS || owner_sid.is_null() {
        unsafe {
            if !psd.is_null() {
                let _ = LocalFree(Some(psd as *const _));
            }
        }
        return String::new();
    }

    let mut name = [0u16; 512];
    let mut domain = [0u16; 256];
    let mut cch_name = name.len() as u32;
    let mut cch_domain = domain.len() as u32;
    let mut use_e = SID_NAME_USE::default();

    let ok = unsafe {
        LookupAccountSidW(
            PCWSTR::null(),
            owner_sid,
            PWSTR(name.as_mut_ptr()),
            &mut cch_name,
            PWSTR(domain.as_mut_ptr()),
            &mut cch_domain,
            &mut use_e,
        )
    };

    unsafe {
        if !psd.is_null() {
            let _ = LocalFree(Some(psd as *const _));
        }
    }

    if !ok.as_bool() {
        return "No active user".to_string();
    }

    let name_s = String::from_utf16_lossy(&name[..cch_name as usize]);
    name_s.trim_end_matches('\0').to_string()
}

#[cfg(not(windows))]
pub fn file_owner_for_report(_path: &std::path::Path) -> String {
    String::new()
}
