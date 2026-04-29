//! Windows DPAPI protect/unprotect (CurrentUser), compatible with PowerShell ProtectedData.

#[cfg(windows)]
pub fn protect_current_user(plain: &str) -> anyhow::Result<String> {
    use base64::{engine::general_purpose::STANDARD, Engine};
    use std::ffi::c_void;
    use windows::Win32::Foundation::LocalFree;
    use windows::Win32::Security::Cryptography::{
        CryptProtectData, CRYPTPROTECT_UI_FORBIDDEN, CRYPT_DATA_BLOB,
    };

    let mut bytes = plain.as_bytes().to_vec();
    let mut in_blob = CRYPT_DATA_BLOB {
        cbData: bytes.len() as u32,
        pbData: bytes.as_mut_ptr(),
    };
    let mut out_blob = CRYPT_DATA_BLOB::default();
    unsafe {
        CryptProtectData(
            &in_blob,
            windows::core::PCWSTR::null(),
            None,
            None,
            None,
            CRYPTPROTECT_UI_FORBIDDEN,
            &mut out_blob,
        )
        .map_err(|e| anyhow::anyhow!("CryptProtectData: {}", e))?;
        let sl =
            std::slice::from_raw_parts(out_blob.pbData, out_blob.cbData as usize).to_vec();
        let _ = LocalFree(Some(out_blob.pbData as *const c_void));
        Ok(STANDARD.encode(sl))
    }
}

#[cfg(windows)]
pub fn unprotect_current_user(b64: &str) -> anyhow::Result<String> {
    use base64::{engine::general_purpose::STANDARD, Engine};
    use std::ffi::c_void;
    use windows::Win32::Foundation::LocalFree;
    use windows::Win32::Security::Cryptography::{
        CryptUnprotectData, CRYPT_DATA_BLOB,
    };

    let mut raw = STANDARD.decode(b64.trim())?;
    let mut in_blob = CRYPT_DATA_BLOB {
        cbData: raw.len() as u32,
        pbData: raw.as_mut_ptr(),
    };
    let mut out_blob = CRYPT_DATA_BLOB::default();
    unsafe {
        CryptUnprotectData(&mut in_blob, None, None, None, None, 0, &mut out_blob).map_err(
            |e| {
                anyhow::anyhow!(
                    "Cannot decrypt SMTP password on this account/machine (Windows DPAPI). Re-enter the password in email setup. {}",
                    e
                )
            },
        )?;
        let out = String::from_utf8_lossy(std::slice::from_raw_parts(
            out_blob.pbData,
            out_blob.cbData as usize,
        ))
        .into_owned();
        let _ = LocalFree(Some(out_blob.pbData as *const c_void));
        Ok(out)
    }
}

#[cfg(not(windows))]
pub fn protect_current_user(_plain: &str) -> anyhow::Result<String> {
    anyhow::bail!("DPAPI is only available on Windows.")
}

#[cfg(not(windows))]
pub fn unprotect_current_user(_b64: &str) -> anyhow::Result<String> {
    anyhow::bail!("DPAPI is only available on Windows.")
}
