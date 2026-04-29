//! Append run summary to Archive-OldFiles.run.log (same format as PowerShell).

use crate::job::ArchiveRow;
use crate::job::JobResult;
use chrono::Local;
use std::fs::OpenOptions;
use std::io::Write;
use std::path::Path;

#[allow(clippy::too_many_arguments)]
pub fn append_run_log(
    run_log_path: &Path,
    version: &str,
    input_path: &str,
    input_resolved: &str,
    archive_path: &str,
    archive_resolved: &str,
    years: f64,
    age_basis: &str,
    cutoff_display: &str,
    top_level_entry_count: isize,
    depth1_file_count: isize,
    top_level_folder_count: isize,
    gci_error_count: usize,
    first_gci_error: &str,
    all_shares: bool,
    file_scan_count: usize,
    skipped_too_new: usize,
    listed_for_archive: usize,
    out_format: &str,
    do_commit: bool,
    planned_moved_count: usize,
    failed_count: usize,
    reclaim_bytes: u64,
    reclaim_display: &str,
    results: &[ArchiveRow],
    share_jobs: Option<&[(String, Vec<ArchiveRow>)]>,
) -> anyhow::Result<()> {
    let tab = '\t';
    let mut block = String::new();
    block.push_str(&format!(
        "----- {} -----\n",
        Local::now().format("%Y-%m-%d %H:%M:%S")
    ));
    block.push_str(&format!("Version={}\n", version));
    block.push_str(&format!("InputPath={}\n", input_path));
    block.push_str(&format!("ResolvedInput={}\n", input_resolved));
    block.push_str(&format!("ArchivePath={}\n", archive_path));
    block.push_str(&format!("ResolvedArchive={}\n", archive_resolved));
    block.push_str(&format!(
        "Years={} AgeBasis={} Cutoff={}\n",
        years, age_basis, cutoff_display
    ));
    block.push_str(&format!(
        "Probe_TopLevelEntries={} Probe_RootFiles={} Probe_RootFolders={}\n",
        top_level_entry_count, depth1_file_count, top_level_folder_count
    ));
    block.push_str(&format!(
        "GciErrorCount={} FirstGciError={}\n",
        gci_error_count, first_gci_error
    ));
    if all_shares {
        block.push_str("Mode=AllSharesPreview Commit=false (forced)\n");
    }
    block.push_str(&format!(
        "FilesScanned={} SkippedNewerThanCutoff={} ListedForArchive={} Output={} Commit={}\n",
        file_scan_count, skipped_too_new, listed_for_archive, out_format, do_commit
    ));
    block.push_str(&format!(
        "PlannedMovedCount={} FailedCount={} ReclaimBytes_PlannedMoved={} ReclaimHuman={}\n",
        planned_moved_count, failed_count, reclaim_bytes, reclaim_display
    ));
    block.push_str("FILES_MEETING_AGE_RULE_TAB_SEPARATED_SourcePath_Bytes_Status_ComparedForAge_ISO\n");

    if let Some(jobs) = share_jobs {
        for (share_root, rows) in jobs {
            block.push_str(&format!("SHARE_BEGIN={}\n", share_root));
            for row in rows {
                block.push_str(&format!(
                    "{}{}{}{}{}{}{}{}\n",
                    row.source_path,
                    tab,
                    row.length,
                    tab,
                    row.status,
                    tab,
                    row.compared_for_age.format("%+").to_string()
                ));
            }
            block.push_str("SHARE_END\n");
        }
    } else {
        for row in results {
            block.push_str(&format!(
                "{}{}{}{}{}{}{}{}\n",
                row.source_path,
                tab,
                row.length,
                tab,
                row.status,
                tab,
                row.compared_for_age.format("%+").to_string()
            ));
        }
    }

    block.push_str("END_FILES_MEETING_AGE_RULE\n");
    block.push_str("-----\n");

    let mut f = OpenOptions::new()
        .create(true)
        .append(true)
        .open(run_log_path)?;
    f.write_all(block.as_bytes())?;
    Ok(())
}

/// Convenience: single-job log from `JobResult`.
pub fn append_run_log_single(
    run_log_path: &Path,
    version: &str,
    input_path: &str,
    archive_path: &str,
    job: &JobResult,
    do_commit: bool,
) -> anyhow::Result<()> {
    append_run_log(
        run_log_path,
        version,
        input_path,
        &job.input_resolved,
        archive_path,
        &job.archive_resolved,
        job.years_num,
        job.age_basis_effective.as_str(),
        &job.cutoff.format("%Y-%m-%d %H:%M").to_string(),
        job.top_level_entry_count,
        job.depth1_file_count,
        job.top_level_folder_count,
        job.gci_error_count,
        &job.first_gci_error_log,
        false,
        job.file_scan_count,
        job.skipped_too_new,
        job.results.len(),
        &job.out_format,
        do_commit,
        job.planned_moved_count,
        job.failed_count,
        job.reclaim_bytes,
        &job.reclaim_display,
        &job.results,
        None,
    )
}
