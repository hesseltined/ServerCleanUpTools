//! Single-tree archive scan: enumerate, age filter, plan/move, optional empty-dir prune, report.

use crate::age_basis::{cutoff_from_years, file_age_timestamp, AgeBasis};
use crate::format_util::format_data_size;
use crate::html_report::{self, ReportRow};
use crate::path_util::{move_file_best_effort, relative_path_from_root, unique_destination_path};
use crate::win_owner::file_owner_for_report;
use chrono::{DateTime, Local};
use dialoguer::Input;
use std::fs;
use std::path::{Path, PathBuf};
use walkdir::WalkDir;

#[derive(Clone, Debug)]
pub struct ArchiveRow {
    pub source_path: String,
    pub destination_path: String,
    pub compared_for_age: DateTime<Local>,
    pub length: u64,
    pub owner: String,
    pub status: String,
    pub message: String,
}

#[derive(Clone)]
pub struct JobResult {
    pub input_resolved: String,
    pub archive_resolved: String,
    pub results: Vec<ArchiveRow>,
    pub file_scan_count: usize,
    pub skipped_too_new: usize,
    pub planned_moved_count: usize,
    pub failed_count: usize,
    pub reclaim_bytes: u64,
    pub reclaim_display: String,
    pub html_report_path: Option<PathBuf>,
    pub out_format: String,
    pub top_level_entry_count: isize,
    pub depth1_file_count: isize,
    pub top_level_folder_count: isize,
    pub gci_error_count: usize,
    pub first_gci_error_log: String,
    pub remove_empty_choice: Option<String>,
    pub cutoff: DateTime<Local>,
    pub age_basis_effective: AgeBasis,
    pub years_num: f64,
}

#[allow(clippy::too_many_arguments)]
pub fn invoke_single_archive_job(
    input_resolved: &Path,
    archive_resolved: &Path,
    years_num: f64,
    age_basis: AgeBasis,
    do_commit: bool,
    out_format_in: &str,
    share_name_label: Option<&str>,
    bound_remove_empty: bool,
    remove_empty_param: &str,
    remove_empty_pref: &str,
    script_version: &str,
    interactive: bool,
) -> anyhow::Result<JobResult> {
    let input_s = path_util_trim(input_resolved);
    let archive_s = path_util_trim(archive_resolved);
    let cutoff = cutoff_from_years(years_num);

    println!();
    println!("========== ARCHIVE RUN ==========");
    println!("  Source tree (InputPath):  {}", input_s);
    println!("  Archive root:             {}", archive_s);
    println!(
        "  Age rule: older than {} year(s); using {} (must be strictly before {}).",
        years_num,
        age_basis.as_str(),
        cutoff.format("%Y-%m-%d %H:%M")
    );
    if age_basis == AgeBasis::LastAccessTime {
        println!("  Note: LastAccessTime depends on NTFS last-access tracking; if the volume disables it, dates may look stale.");
    }
    println!(
        "  Mode: {}",
        if do_commit {
            "COMMIT - files will be moved to the archive"
        } else {
            "PREVIEW - no moves; plan only"
        }
    );
    println!("=================================");
    println!();

    // Step 1: root listing
    let (top_level_entry_count, depth1_file_count, top_level_folder_count) =
        match fs::read_dir(input_resolved) {
            Ok(rd) => {
                let mut n = 0;
                let mut nf = 0;
                let mut nd = 0;
                for e in rd.flatten() {
                    n += 1;
                    let Ok(ft) = e.file_type() else { continue };
                    if ft.is_file() {
                        nf += 1;
                    } else if ft.is_dir() {
                        nd += 1;
                    }
                }
                (n as isize, nf as isize, nd as isize)
            }
            Err(e) => {
                anyhow::bail!(
                    "Could not list the root of InputPath (before recursive scan). Path: {}  {}",
                    input_s,
                    e
                );
            }
        };

    println!("--- Step 1: Can we read the source folder? (top level only) ---");
    println!(
        "  Entries at root: {} ({} files, {} subfolders here).",
        top_level_entry_count, depth1_file_count, top_level_folder_count
    );
    println!();

    let mut gci_errors: Vec<String> = Vec::new();
    let mut files: Vec<PathBuf> = Vec::new();

    for w in WalkDir::new(input_resolved).follow_links(false).into_iter() {
        match w {
            Ok(entry) => {
                if entry.file_type().is_file() {
                    files.push(entry.path().to_path_buf());
                }
            }
            Err(e) => {
                gci_errors.push(e.to_string());
            }
        }
    }

    let gci_err_count = gci_errors.len();
    if gci_err_count > 0 {
        println!();
        println!(
            "RECURSIVE SCAN REPORTED {} ERROR(S) - results may be incomplete (0 files can mean everything below was blocked). First errors:",
            gci_err_count
        );
        let show = gci_err_count.min(8);
        for (i, err) in gci_errors.iter().take(show).enumerate() {
            println!("  [{}] {}", i + 1, err);
        }
        if gci_err_count > show {
            println!(
                "  ... and {} more (see warnings above).",
                gci_err_count - show
            );
        }
        println!();
        for err in &gci_errors {
            eprintln!("Warning: Enumeration issue: {}", err);
        }
    }

    let first_gci_error_log = gci_errors
        .first()
        .map(|s| {
            let mut em = s.replace(['\r', '\n'], " ");
            if em.len() > 240 {
                em.truncate(237);
                em.push_str("...");
            }
            em
        })
        .unwrap_or_default();

    let file_scan_count = files.len();
    println!("--- Step 2: Recursive file list under source ---");
    println!("  Total files found (all ages): {}", file_scan_count);
    if file_scan_count == 0 {
        println!("  No files returned by recursive walk.");
        println!("  Check: empty tree, permissions on subfolders, or wrong InputPath.");
    }
    println!();

    println!("--- Step 3: Compare each file to the cutoff (older files are listed below) ---");
    println!();

    let input_canon = fs::canonicalize(input_resolved)?;
    let archive_canon = fs::canonicalize(archive_resolved)?;

    let mut results: Vec<ArchiveRow> = Vec::new();
    let mut skipped_too_new = 0usize;

    for file_path in files {
        let source_path = file_path.clone();
        let source_str = source_path.to_string_lossy().to_string();
        let meta = match fs::metadata(&file_path) {
            Ok(m) => m,
            Err(e) => {
                results.push(ArchiveRow {
                    source_path: source_str.clone(),
                    destination_path: String::new(),
                    compared_for_age: Local::now(),
                    length: 0,
                    owner: file_owner_for_report(&file_path),
                    status: "Failed".to_string(),
                    message: e.to_string(),
                });
                eprintln!("FAILED: {} - {}", source_str, e);
                continue;
            }
        };

        let compared = file_age_timestamp(&meta, age_basis);
        let length = meta.len();

        if compared >= cutoff {
            skipped_too_new += 1;
            continue;
        }

        let mut dest_path_str = String::new();
        let rel = match relative_path_from_root(&file_path, &input_canon) {
            Ok(r) => r,
            Err(e) => {
                results.push(ArchiveRow {
                    source_path: source_str.clone(),
                    destination_path: String::new(),
                    compared_for_age: compared,
                    length,
                    owner: file_owner_for_report(&file_path),
                    status: "Failed".to_string(),
                    message: e.to_string(),
                });
                eprintln!("FAILED: {} - {}", source_str, e);
                continue;
            }
        };

        let initial_dest = archive_canon.join(&rel);
        let dest_path = unique_destination_path(&initial_dest);
        dest_path_str = dest_path.to_string_lossy().to_string();

        if !do_commit {
            let owner_snap = file_owner_for_report(&file_path);
            results.push(ArchiveRow {
                source_path: source_str,
                destination_path: dest_path_str,
                compared_for_age: compared,
                length,
                owner: owner_snap,
                status: "Planned".to_string(),
                message: String::new(),
            });
            continue;
        }

        let owner_snap = file_owner_for_report(&file_path);
        if let Some(parent) = dest_path.parent() {
            fs::create_dir_all(parent)?;
        }

        match move_file_best_effort(&file_path, &dest_path) {
            Ok(()) => {
                results.push(ArchiveRow {
                    source_path: source_str,
                    destination_path: dest_path_str,
                    compared_for_age: compared,
                    length,
                    owner: owner_snap,
                    status: "Moved".to_string(),
                    message: String::new(),
                });
            }
            Err(e) => {
                results.push(ArchiveRow {
                    source_path: source_str.clone(),
                    destination_path: dest_path_str,
                    compared_for_age: compared,
                    length,
                    owner: owner_snap,
                    status: "Failed".to_string(),
                    message: e.to_string(),
                });
                eprintln!("FAILED: {} - {}", source_str, e);
            }
        }
    }

    let planned_moved_count = results
        .iter()
        .filter(|r| r.status == "Planned" || r.status == "Moved")
        .count();
    let failed_count = results.iter().filter(|r| r.status == "Failed").count();
    let reclaim_bytes: u64 = results
        .iter()
        .filter(|r| r.status == "Planned" || r.status == "Moved")
        .map(|r| r.length)
        .sum();
    let reclaim_display = format_data_size(reclaim_bytes);

    println!();
    println!("--- Step 4: Age filter summary ---");
    println!(
        "  Files skipped (too new, on or after cutoff): {}",
        skipped_too_new
    );
    println!(
        "  Files listed for archive (met age rule):     {}",
        results.len()
    );
    println!(
        "    - Planned or Moved (success path):         {}",
        planned_moved_count
    );
    println!(
        "    - Failed (met age but error on move/plan): {}",
        failed_count
    );
    println!(
        "  Total size of Planned+Moved (space no longer on source after a successful move): {} ({} bytes)",
        reclaim_display, reclaim_bytes
    );
    println!();

    let mut remove_empty_choice: Option<String> = None;
    if do_commit {
        let do_prune = if bound_remove_empty {
            remove_empty_param.eq_ignore_ascii_case("yes")
        } else if remove_empty_pref.eq_ignore_ascii_case("yes")
            || remove_empty_pref.eq_ignore_ascii_case("no")
        {
            remove_empty_pref.eq_ignore_ascii_case("yes")
        } else if interactive {
            let ans: String = Input::new()
                .with_prompt("Remove empty folders under the input tree (never the input root)? [y/N]")
                .allow_empty(true)
                .interact_text()?;
            ans.eq_ignore_ascii_case("y") || ans.eq_ignore_ascii_case("yes")
        } else {
            false
        };

        if do_prune {
            remove_empty_choice = Some("Yes".to_string());
            let mut dirs: Vec<PathBuf> = WalkDir::new(input_resolved)
                .min_depth(1)
                .follow_links(false)
                .into_iter()
                .filter_map(|e| e.ok())
                .filter(|e| e.file_type().is_dir())
                .map(|e| e.path().to_path_buf())
                .collect();
            dirs.sort_by_key(|p| std::cmp::Reverse(p.as_os_str().len()));
            for dir in dirs {
                if path_util_trim(&dir) == input_s {
                    continue;
                }
                let count = match fs::read_dir(&dir) {
                    Ok(rd) => rd.count(),
                    Err(_) => continue,
                };
                if count == 0 {
                    match fs::remove_dir(&dir) {
                        Ok(()) => println!("Removed empty folder: {}", dir.display()),
                        Err(e) => eprintln!(
                            "Warning: Could not remove folder {}: {}",
                            dir.display(),
                            e
                        ),
                    }
                }
            }
        } else {
            remove_empty_choice = Some("No".to_string());
        }
    }

    let mut out_format = crate::config::normalize_output_format(Some(out_format_in))
        .unwrap_or_default();
    if out_format.is_empty() && (out_format_in.is_empty() || interactive) {
        loop {
            let r: String = Input::new()
                .with_prompt("Output format: type Text or HTML")
                .interact_text()?;
            let t = r.trim();
            if let Some(n) = crate::config::normalize_output_format(Some(t)) {
                out_format = n;
                break;
            }
        }
    } else if out_format.is_empty() {
        out_format = "Text".to_string();
    }

    let html_report_path = if out_format == "HTML" {
        let domain = html_report::report_domain_label();
        let fname = html_report::archive_html_report_filename(
            &domain,
            &input_s,
            share_name_label,
        );
        let html_path = archive_canon.join(&fname);
        let report_rows: Vec<ReportRow> = results
            .iter()
            .map(|r| ReportRow {
                source_path: r.source_path.clone(),
                compared_for_age: r.compared_for_age,
                length: r.length,
                owner: r.owner.clone(),
                status: r.status.clone(),
                message: r.message.clone(),
            })
            .collect();
        html_report::write_archive_html_report(
            &html_path,
            &report_rows,
            &input_s,
            &archive_s,
            years_num,
            age_basis,
            cutoff,
            do_commit,
            script_version,
            file_scan_count,
            skipped_too_new,
            planned_moved_count,
            failed_count,
            reclaim_bytes,
            &reclaim_display,
            &domain,
        )?;
        println!();
        println!("HTML report written: {}", html_path.display());
        println!("  Rows (met age rule): {}", results.len());
        Some(html_path)
    } else {
        println!();
        println!("========== DETAIL: FILES THAT MET THE AGE RULE ==========");
        if results.is_empty() {
            println!("  (none)");
        } else {
            for r in &results {
                println!(
                    "{}\t{}\t{}\t{}\t{}",
                    r.source_path,
                    r.owner,
                    r.compared_for_age.format("%Y-%m-%d %H:%M"),
                    r.length,
                    r.status
                );
            }
        }
        println!("=========================================================");
        None
    };

    println!();
    println!("==============================================================================");
    println!(" FINAL SUMMARY (this source)");
    println!("==============================================================================");
    println!("  Script version:          {}", script_version);
    println!("  InputPath (resolved):    {}", input_s);
    println!("  ArchivePath (resolved):  {}", archive_s);
    println!(
        "  Years / AgeBasis:        {} / {}",
        years_num,
        age_basis.as_str()
    );
    println!("  Cutoff (exclusive):      {}", cutoff);
    println!(
        "  Mode:                    {}",
        if do_commit {
            "Commit (moves performed where successful)"
        } else {
            "Preview (no moves)"
        }
    );
    println!("  Output format this run:  {}", out_format);
    if let Some(ref p) = html_report_path {
        println!("  HTML report path:        {}", p.display());
    }
    println!("  Files scanned (all ages): {}", file_scan_count);
    println!("  Skipped (too new):        {}", skipped_too_new);
    println!("  Listed (met age rule):    {}", results.len());
    println!("    Planned + Moved:        {}", planned_moved_count);
    println!("    Failed:                 {}", failed_count);
    println!(
        "  Source space from Planned+Moved: {} ({} bytes)",
        reclaim_display, reclaim_bytes
    );
    println!("    (After commit, this much file data no longer lives under the source tree.)");
    println!("==============================================================================");
    println!();

    Ok(JobResult {
        input_resolved: input_s,
        archive_resolved: archive_s,
        results,
        file_scan_count,
        skipped_too_new,
        planned_moved_count,
        failed_count,
        reclaim_bytes,
        reclaim_display,
        html_report_path,
        out_format,
        top_level_entry_count,
        depth1_file_count,
        top_level_folder_count,
        gci_error_count,
        first_gci_error_log,
        remove_empty_choice,
        cutoff,
        age_basis_effective: age_basis,
        years_num,
    })
}

fn path_util_trim(p: &Path) -> String {
    crate::path_util::trim_sep(&p.to_string_lossy())
}
