//! Human-readable byte sizes (matches PowerShell Format-DataSize).

pub fn format_data_size(bytes: u64) -> String {
    let b = bytes as f64;
    const KB: f64 = 1024.0;
    const MB: f64 = KB * 1024.0;
    const GB: f64 = MB * 1024.0;
    const TB: f64 = GB * 1024.0;
    if b >= TB {
        return format!("{:.2} TB", b / TB);
    }
    if b >= GB {
        return format!("{:.2} GB", b / GB);
    }
    if b >= MB {
        return format!("{:.2} MB", b / MB);
    }
    if b >= KB {
        return format!("{:.2} KB", b / KB);
    }
    format!("{} bytes", bytes)
}
