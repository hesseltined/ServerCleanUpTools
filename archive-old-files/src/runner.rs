//! Orchestration: config, validation, -All, save, run log, email.

use crate::age_basis::{normalize_age_basis_param, AgeBasis};
use crate::cli::Cli;
use crate::config::{
    apply_smtp_preset, email_from_json, normalize_output_format, normalize_smtp_profile,
    parse_age_basis_saved, read_config, save_config, EmailRuntime,
};
use crate::email::send_archive_email;
use crate::format_util::format_data_size;
use crate::html_report;
use crate::job::{self as archive_job, ArchiveRow, JobResult};
use crate::path_util::{archive_under_input, normalize_user_path};
use crate::run_log;
use crate::win_dpapi::protect_current_user;
use crate::win_shares::{local_path_for_share_name, published_disk_shares};
use dialoguer::{Confirm, Input, Select};
use std::fs;
use std::path::{Path, PathBuf};

fn default_config_path() -> PathBuf {
    std::env::current_exe()
        .ok()
        .and_then(|p| p.parent().map(|d| d.join("Archive-OldFiles.config.json")))
        .unwrap_or_else(|| PathBuf::from("Archive-OldFiles.config.json"))
}

fn parse_years(s: &str) -> anyhow::Result<f64> {
    let t = s.trim();
    t.parse::<f64>()
        .map_err(|_| anyhow::anyhow!("Cannot convert '-Years' to a number: {}", s))
}

fn bound_bool(s: Option<&str>, default: bool) -> bool {
    match s.map(|x| x.trim().to_ascii_lowercase()) {
        None => default,
        Some(ref v) if v == "true" || v == "yes" || v == "1" => true,
        Some(ref v) if v == "false" || v == "no" || v == "0" => false,
        _ => default,
    }
}

fn resolve_input_via_share_menu() -> anyhow::Result<String> {
    let shares = published_disk_shares();
    if shares.is_empty() {
        println!("No published disk shares were returned (check permissions or use --skip-share-menu).");
        let p: String = Input::new()
            .with_prompt("InputPath (full local or UNC path)")
            .interact_text()?;
        return Ok(normalize_user_path(&p));
    }
    println!();
    println!("Published folder shares on this server (select INPUT / source to scan):");
    let mut opts: Vec<String> = shares
        .iter()
        .enumerate()
        .map(|(i, s)| {
            format!(
                "{}) {}  ->  {}  ({})",
                i + 1,
                s.name,
                s.local_path,
                s.unc
            )
        })
        .collect();
    opts.push(format!("{}) Other (enter path manually)", shares.len() + 1));

    let sel = Select::new()
        .with_prompt("Choose input share")
        .items(&opts)
        .default(0)
        .interact()?;

    if sel == shares.len() {
        let p: String = Input::new()
            .with_prompt("Full path for InputPath")
            .interact_text()?;
        Ok(normalize_user_path(&p))
    } else {
        Ok(shares[sel].local_path.clone())
    }
}

fn merge_email_from_cli(
    mut e: EmailRuntime,
    cli: &Cli,
    interactive_wizard: bool,
) -> anyhow::Result<EmailRuntime> {
    if let Some(ref t) = cli.email_to {
        if !t.trim().is_empty() {
            e.to = t.trim().to_string();
            e.enabled = true;
        }
    }
    if let Some(ref f) = cli.email_from {
        if !f.trim().is_empty() {
            e.from = f.trim().to_string();
        }
    }
    if let Some(ref p) = cli.smtp_profile {
        if !p.trim().is_empty() {
            e.profile = normalize_smtp_profile(p);
            apply_smtp_preset(&e.profile, &mut e.smtp_host, &mut e.smtp_port, &mut e.use_ssl);
        }
    }
    if let Some(ref h) = cli.smtp_host {
        if !h.trim().is_empty() {
            e.smtp_host = h.trim().to_string();
        }
    }
    if let Some(port) = cli.smtp_port {
        if port > 0 {
            e.smtp_port = port;
        }
    }
    if let Some(ref s) = cli.smtp_use_ssl {
        e.use_ssl = bound_bool(Some(s.as_str()), e.use_ssl);
    }
    if let Some(ref u) = cli.smtp_user {
        if !u.trim().is_empty() {
            e.user_name = u.trim().to_string();
        }
    }

    if interactive_wizard {
        let ask_change = if !e.enabled || e.password_protected_base64.is_none() {
            Confirm::new()
                .with_prompt("Set up SMTP email notifications (saved to config)?")
                .default(false)
                .interact()?
        } else {
            Confirm::new()
                .with_prompt("Change SMTP email settings?")
                .default(false)
                .interact()?
        };

        if ask_change {
            e = email_setup_wizard(e)?;
        }
    }

    if e.enabled {
        if e.to.trim().is_empty() || e.from.trim().is_empty() {
            eprintln!("Warning: Email is enabled but To or From is missing; disabling email for this run.");
            e.enabled = false;
        }
        if e.smtp_host.trim().is_empty() {
            eprintln!("Warning: Email is enabled but SmtpHost is missing; disabling email for this run.");
            e.enabled = false;
        }
        if e.password_protected_base64.is_none() {
            eprintln!("Warning: Email is enabled but no saved SMTP password; run interactive setup or disable Email.");
            e.enabled = false;
        }
    }

    Ok(e)
}

fn email_setup_wizard(mut seed: EmailRuntime) -> anyhow::Result<EmailRuntime> {
    println!();
    println!("--- Email notification (SMTP) ---");
    println!("Passwords are stored with Windows DPAPI (CurrentUser) and only work for this user on this computer.");

    let enable = Confirm::new()
        .with_prompt("Send email when a run completes?")
        .default(false)
        .interact()?;
    if !enable {
        seed.enabled = false;
        return Ok(seed);
    }
    seed.enabled = true;

    seed.to = Input::new()
        .with_prompt("To (comma-separated)")
        .default(seed.to.clone())
        .interact_text()?;

    seed.from = Input::new()
        .with_prompt("From (sender address)")
        .default(seed.from.clone())
        .interact_text()?;

    let profiles = &[
        "Gmail (app password)",
        "Microsoft 365 / Outlook (app password)",
        "SMTP2GO",
        "Custom server",
    ];
    let pick = Select::new()
        .with_prompt("SMTP profile")
        .items(profiles)
        .default(0)
        .interact()?;

    seed.profile = match pick {
        1 => "Microsoft365AppPassword".to_string(),
        2 => "Smtp2Go".to_string(),
        3 => "Custom".to_string(),
        _ => "GmailAppPassword".to_string(),
    };
    apply_smtp_preset(
        &seed.profile,
        &mut seed.smtp_host,
        &mut seed.smtp_port,
        &mut seed.use_ssl,
    );

    if seed.profile == "Custom" {
        seed.smtp_host = Input::new()
            .with_prompt("SMTP host")
            .default(seed.smtp_host.clone())
            .interact_text()?;
        let port_s: String = Input::new()
            .with_prompt("SMTP port")
            .default(seed.smtp_port.to_string())
            .interact_text()?;
        if let Ok(p) = port_s.trim().parse::<u16>() {
            if p > 0 {
                seed.smtp_port = p;
            }
        }
        seed.use_ssl = Confirm::new()
            .with_prompt("Use TLS/SSL (recommended for ports 587/2525)?")
            .default(seed.use_ssl)
            .interact()?;
    }

    seed.user_name = Input::new()
        .with_prompt("SMTP user name (often your mailbox)")
        .default(seed.user_name.clone())
        .interact_text()?;

    let pass = dialoguer::Password::new()
        .with_prompt("SMTP password or app password")
        .interact()?;
    if !pass.trim().is_empty() {
        seed.password_protected_base64 = Some(protect_current_user(pass.trim())?);
    }

    println!("--- End email setup ---");
    println!();
    Ok(seed)
}

fn validate_issues(
    input_path: &str,
    archive_path: &str,
    years: f64,
    do_commit: bool,
    all_shares: bool,
) -> Vec<(String, String)> {
    let mut issues = Vec::new();
    if !all_shares {
        if input_path.trim().is_empty() {
            issues.push(("InputPath".into(), "InputPath is empty.".into()));
        } else if !Path::new(input_path).exists() {
            issues.push((
                "InputPath".into(),
                format!("InputPath does not exist: {}", input_path),
            ));
        }
    }
    if archive_path.trim().is_empty() {
        issues.push(("ArchivePath".into(), "ArchivePath is empty.".into()));
    } else if !Path::new(archive_path).exists() && !do_commit {
        issues.push((
            "ArchivePath".into(),
            "ArchivePath does not exist. Create it for preview, or use --commit to create it."
                .into(),
        ));
    }
    if years <= 0.0 || years.is_nan() {
        issues.push((
            "Years".into(),
            "Years must be a number greater than zero.".into(),
        ));
    }
    issues
}

pub fn run(cli: &Cli) -> anyhow::Result<()> {
    let config_file = cli
        .config
        .clone()
        .unwrap_or_else(default_config_path);

    if cli.test_email {
        return run_test_email(cli, &config_file);
    }

    let mut years: f64 = cli
        .years
        .as_ref()
        .map(|s| parse_years(s))
        .transpose()?
        .unwrap_or(0.0);
    let years_from_cli = cli.years.is_some();

    let mut input = cli.input.clone().unwrap_or_default();
    let mut archive = cli.archive.clone().unwrap_or_default();
    let mut do_commit = cli.commit;
    let mut out_pref = cli.output.clone().unwrap_or_default();
    let mut remove_pref = cli.remove_empty_folders.clone().unwrap_or_default();
    let mut age_basis = AgeBasis::LastWriteTime;
    if let Some(ref ab) = cli.age_basis {
        if !ab.trim().is_empty() {
            age_basis = normalize_age_basis_param(ab.trim()).ok_or_else(|| {
                anyhow::anyhow!(
                    "-AgeBasis not recognized. Use LastWriteTime, LastAccessTime, LatestWriteOrAccess, CreationTime, Earliest."
                )
            })?;
        }
    }

    let mut effective_archive_share = cli.archive_share_name.trim().to_string();
    if effective_archive_share.is_empty() {
        effective_archive_share = "Archive".to_string();
    }

    let saved = read_config(&config_file)?;

    let core_on_cli = !archive.trim().is_empty()
        && years > 0.0
        && (cli.all || !input.trim().is_empty());

    let interactive = !core_on_cli;

    if cli.all && !input.trim().is_empty() {
        println!("Note: --all ignores InputPath; each published disk share is scanned in preview-only mode.");
    }

    if core_on_cli {
        println!("Core parameters supplied on command line; skipping saved-config load for InputPath, ArchivePath, and Years.");
        if let Some(ref s) = saved {
            if cli.archive_share_name == "Archive"
                && s.archive_share_name.as_ref().map(|x| x.trim()) != Some("")
            {
                if let Some(ref n) = s.archive_share_name {
                    effective_archive_share = n.trim().to_string();
                }
            }
            if !cli.commit && !cli.all {
                if let Some(c) = s.commit {
                    do_commit = c;
                }
            }
            if cli.output.is_none() {
                if let Some(ref o) = s.output {
                    if let Some(n) = normalize_output_format(Some(o)) {
                        out_pref = n;
                    }
                }
            }
            if cli.remove_empty_folders.is_none() {
                if let Some(ref r) = s.remove_empty_folders {
                    if r.eq_ignore_ascii_case("yes") || r.eq_ignore_ascii_case("no") {
                        remove_pref = r.clone();
                    }
                }
            }
            if cli.age_basis.is_none() {
                if let Some(ab) = parse_age_basis_saved(s.age_basis.as_deref()) {
                    age_basis = ab;
                }
            }
        }
    } else if let Some(ref s) = saved {
        if let Some(ref n) = s.archive_share_name {
            if !n.trim().is_empty() {
                effective_archive_share = n.trim().to_string();
            }
        }

        let use_saved = Confirm::new()
            .with_prompt(format!(
                "Use saved config? Input={} Archive={} Years={:?}",
                s.input_path, s.archive_path, s.years
            ))
            .default(true)
            .interact()?;

        if use_saved {
            if cli.input.is_none() {
                input = s.input_path.clone();
            }
            if cli.archive.is_none() {
                archive = s.archive_path.clone();
            }
            if !years_from_cli {
                years = s.years.unwrap_or(0.0);
            }
            if !cli.commit && !cli.all {
                if let Some(c) = s.commit {
                    do_commit = c;
                }
            }
            if cli.output.is_none() {
                if let Some(n) = normalize_output_format(s.output.as_deref()) {
                    out_pref = n;
                }
            }
            if cli.remove_empty_folders.is_none() {
                if let Some(ref r) = s.remove_empty_folders {
                    if r.eq_ignore_ascii_case("yes") || r.eq_ignore_ascii_case("no") {
                        remove_pref = r.clone();
                    }
                }
            }
            if cli.age_basis.is_none() {
                if let Some(ab) = parse_age_basis_saved(s.age_basis.as_deref()) {
                    age_basis = ab;
                }
            }
        } else {
            prompt_missing_paths_and_options(
                &mut input,
                &mut archive,
                &mut years,
                &mut do_commit,
                &mut out_pref,
                &mut remove_pref,
                &mut age_basis,
                &effective_archive_share,
                s,
                cli,
            )?;
        }
    } else {
        prompt_missing_paths_and_options(
            &mut input,
            &mut archive,
            &mut years,
            &mut do_commit,
            &mut out_pref,
            &mut remove_pref,
            &mut age_basis,
            &effective_archive_share,
            &None,
            cli,
        )?;
    }

    input = normalize_user_path(&input);
    archive = normalize_user_path(&archive);

    // Validation loop
    let mut attempt = 0u32;
    loop {
        attempt += 1;
        if attempt > 25 {
            anyhow::bail!("Configuration could not be validated after 25 attempts.");
        }
        let validate_commit = do_commit && !cli.all;
        let issues = validate_issues(&input, &archive, years, validate_commit, cli.all);
        if issues.is_empty() {
            break;
        }

        let only_archive_missing = !validate_commit
            && issues.len() == 1
            && issues[0].0 == "ArchivePath"
            && issues[0].1.contains("does not exist");

        if only_archive_missing {
            if cli.all {
                println!("The archive folder does not exist yet. Create it manually first; --all only writes HTML reports.");
            } else if Confirm::new()
                .with_prompt("Use --commit for this run so the archive folder can be created?")
                .default(false)
                .interact()?
            {
                do_commit = true;
                continue;
            }
        }

        println!();
        println!("Configuration check — correct the following:");
        for (f, m) in &issues {
            println!("  - [{}] {}", f, m);
        }
        println!();

        for (field, _) in &issues {
            match field.as_str() {
                "InputPath" => {
                    if !cli.skip_share_menu {
                        if let Ok(p) = resolve_input_via_share_menu() {
                            input = normalize_user_path(&p);
                            continue;
                        }
                    }
                    let p: String = Input::new()
                        .with_prompt("InputPath")
                        .default(input.clone())
                        .interact_text()?;
                    input = normalize_user_path(&p);
                }
                "ArchivePath" => {
                    if let Some(p) = local_path_for_share_name(&effective_archive_share) {
                        if Confirm::new()
                            .with_prompt(format!(
                                "Use archive share '{}' -> {}?",
                                effective_archive_share, p
                            ))
                            .default(true)
                            .interact()?
                        {
                            archive = p;
                            continue;
                        }
                    }
                    let p: String = Input::new()
                        .with_prompt("ArchivePath")
                        .default(archive.clone())
                        .interact_text()?;
                    archive = normalize_user_path(&p);
                }
                "Years" => {
                    let p: String = Input::new()
                        .with_prompt("Years")
                        .default(years.to_string())
                        .interact_text()?;
                    if let Ok(y) = parse_years(&p) {
                        if y > 0.0 {
                            years = y;
                        }
                    }
                }
                _ => {}
            }
        }
    }

    if cli.all {
        do_commit = false;
        out_pref = "HTML".to_string();
    }

    if let Some(n) = normalize_output_format(Some(&out_pref)) {
        out_pref = n;
    }

    let config_snapshot = read_config(&config_file)?;
    let email_rt = merge_email_from_cli(
        email_from_json(
            config_snapshot
                .as_ref()
                .and_then(|c| c.email.as_ref()),
        ),
        cli,
        interactive,
    )?;

    let script_version = crate::cli::VERSION;

    if cli.all {
        return run_all_shares(
            &archive,
            years,
            age_basis,
            script_version,
            interactive,
            cli,
            &config_file,
            &input,
            &effective_archive_share,
            do_commit,
            &out_pref,
            &remove_pref,
            &email_rt,
        );
    }

    let input_resolved = fs::canonicalize(Path::new(&input))?;
    if !input_resolved.is_dir() {
        anyhow::bail!("InputPath must be a folder: {}", input_resolved.display());
    }

    let archive_path_pb = Path::new(&archive);
    let archive_resolved = if archive_path_pb.exists() {
        fs::canonicalize(archive_path_pb)?
    } else if do_commit {
        fs::create_dir_all(archive_path_pb)?;
        fs::canonicalize(archive_path_pb)?
    } else {
        anyhow::bail!("ArchivePath does not exist.");
    };

    let in_s = input_resolved.to_string_lossy();
    let ar_s = archive_resolved.to_string_lossy();
    if archive_under_input(&in_s, &ar_s) {
        anyhow::bail!(
            "ArchivePath must not be the same as or inside InputPath. Input: {}  Archive: {}",
            in_s,
            ar_s
        );
    }

    let bound_rm = cli.remove_empty_folders.is_some();
    let rm_param = cli
        .remove_empty_folders
        .as_deref()
        .unwrap_or("")
        .to_string();

    let job = archive_job::invoke_single_archive_job(
        &input_resolved,
        &archive_resolved,
        years,
        age_basis,
        do_commit,
        &out_pref,
        None,
        bound_rm,
        &rm_param,
        &remove_pref,
        script_version,
        interactive,
    )?;

    if !cli.no_save_config {
        match save_config(
            &config_file,
            script_version,
            &input,
            &archive,
            &effective_archive_share,
            years,
            age_basis,
            do_commit,
            &job.out_format,
            job.remove_empty_choice.as_deref(),
            false,
            &email_rt,
        ) {
            Ok(()) => println!("Saved settings for next run: {}", config_file.display()),
            Err(e) => eprintln!("Warning: Could not save config file: {}", e),
        }
    }

    if !cli.no_run_log {
        let run_log_path = config_file
            .parent()
            .unwrap_or_else(|| Path::new("."))
            .join("Archive-OldFiles.run.log");
        if let Err(e) = run_log::append_run_log_single(
            &run_log_path,
            script_version,
            &input,
            &archive,
            &job,
            do_commit,
        ) {
            eprintln!("Warning: Could not write run log: {}", e);
        } else {
            println!(
                "Run log (includes per-file list): {}",
                run_log_path.display()
            );
        }
    }

    if !cli.no_send_email && email_rt.enabled {
        send_summary_email(
            &email_rt,
            script_version,
            &input_resolved.to_string_lossy(),
            &archive_resolved.to_string_lossy(),
            years,
            age_basis,
            do_commit,
            &job,
            None,
            &[],
        )?;
    }

    Ok(())
}

fn run_all_shares(
    archive: &str,
    years: f64,
    age_basis: AgeBasis,
    script_version: &str,
    interactive: bool,
    cli: &Cli,
    config_file: &Path,
    _input_display: &str,
    effective_archive_share: &str,
    do_commit: bool,
    out_pref: &str,
    remove_pref: &str,
    email_rt: &EmailRuntime,
) -> anyhow::Result<()> {
    println!();
    println!("========== ALL SHARES (preview only, one HTML report per share) ==========");
    println!("  Report folder (ArchivePath): {}", archive);
    println!(
        "  Years: {}  AgeBasis: {}",
        years,
        age_basis.as_str()
    );
    println!("  --commit and saved Commit are ignored for moves; no files are moved.");
    println!("==========================================================================");
    println!();

    let archive_resolved = fs::canonicalize(Path::new(archive))
        .map_err(|_| anyhow::anyhow!("ArchivePath does not exist. Create the folder first."))?;

    let shares = published_disk_shares();
    if shares.is_empty() {
        anyhow::bail!("No published disk shares found.");
    }

    let mut share_jobs: Vec<JobResult> = Vec::new();
    let mut html_paths: Vec<PathBuf> = Vec::new();

    for sh in shares {
        let share_root = PathBuf::from(&sh.local_path);
        let sr = share_root.to_string_lossy();
        let ar = archive_resolved.to_string_lossy();
        if archive_under_input(&sr, &ar) {
            println!(
                "Skipping share '{}' ({}): archive folder is this path or inside it.",
                sh.name, sr
            );
            continue;
        }
        let canon = match fs::canonicalize(&share_root) {
            Ok(p) => p,
            Err(e) => {
                eprintln!("Warning: share '{}': {}", sh.name, e);
                continue;
            }
        };

        match archive_job::invoke_single_archive_job(
            &canon,
            &archive_resolved,
            years,
            age_basis,
            false,
            "HTML",
            Some(&sh.name),
            false,
            "",
            remove_pref,
            script_version,
            interactive,
        ) {
            Ok(j) => {
                if let Some(ref hp) = j.html_report_path {
                    html_paths.push(hp.clone());
                }
                share_jobs.push(j);
            }
            Err(e) => eprintln!("Warning: share '{}': {}", sh.name, e),
        }
    }

    println!();
    println!("========== ALL-SHARES SUMMARY ==========");
    println!("  Share jobs completed: {}", share_jobs.len());
    println!("  HTML reports written: {}", html_paths.len());
    for hp in &html_paths {
        println!("    {}", hp.display());
    }
    println!("========================================");
    println!();

    let file_scan_count: usize = share_jobs.iter().map(|j| j.file_scan_count).sum();
    let skipped_too_new: usize = share_jobs.iter().map(|j| j.skipped_too_new).sum();
    let planned_moved: usize = share_jobs.iter().map(|j| j.planned_moved_count).sum();
    let failed: usize = share_jobs.iter().map(|j| j.failed_count).sum();
    let reclaim: u64 = share_jobs.iter().map(|j| j.reclaim_bytes).sum();
    let listed: usize = share_jobs.iter().map(|j| j.results.len()).sum();
    let cutoff = share_jobs
        .first()
        .map(|j| j.cutoff)
        .unwrap_or_else(chrono::Local::now);

    if !cli.no_save_config {
        if let Err(e) = save_config(
            config_file,
            script_version,
            "(All published disk shares)",
            archive,
            effective_archive_share,
            years,
            age_basis,
            false,
            out_pref,
            None,
            true,
            email_rt,
        ) {
            eprintln!("Warning: Could not save config file: {}", e);
        } else {
            println!("Saved settings for next run: {}", config_file.display());
        }
    }

    if !cli.no_run_log {
        let run_log_path = config_file
            .parent()
            .unwrap_or_else(|| Path::new("."))
            .join("Archive-OldFiles.run.log");
        let log_pairs: Vec<(String, Vec<ArchiveRow>)> = share_jobs
            .iter()
            .map(|j| (j.input_resolved.clone(), j.results.clone()))
            .collect();
        let gci_total: usize = share_jobs.iter().map(|j| j.gci_error_count).sum();
        let first_gci = share_jobs
            .iter()
            .map(|j| j.first_gci_error_log.as_str())
            .find(|s| !s.is_empty())
            .unwrap_or("");

        if let Err(e) = run_log::append_run_log(
            &run_log_path,
            script_version,
            "(All published disk shares)",
            "(All published disk shares)",
            archive,
            &archive_resolved.to_string_lossy(),
            years,
            age_basis.as_str(),
            &cutoff.format("%Y-%m-%d %H:%M").to_string(),
            -1,
            -1,
            -1,
            gci_total,
            first_gci,
            true,
            file_scan_count,
            skipped_too_new,
            listed,
            "HTML",
            false,
            planned_moved,
            failed,
            reclaim,
            &format_data_size(reclaim),
            &[],
            Some(&log_pairs),
        ) {
            eprintln!("Warning: Could not write run log: {}", e);
        } else {
            println!(
                "Run log (includes per-file list): {}",
                run_log_path.display()
            );
        }
    }

    if !cli.no_send_email && email_rt.enabled {
        let empty_job = JobResult {
            input_resolved: "(All published disk shares)".into(),
            archive_resolved: archive_resolved.to_string_lossy().into(),
            results: vec![],
            file_scan_count,
            skipped_too_new,
            planned_moved_count: planned_moved,
            failed_count: failed,
            reclaim_bytes: reclaim,
            reclaim_display: format_data_size(reclaim),
            html_report_path: None,
            out_format: "HTML".into(),
            top_level_entry_count: -1,
            depth1_file_count: -1,
            top_level_folder_count: -1,
            gci_error_count: 0,
            first_gci_error_log: String::new(),
            remove_empty_choice: None,
            cutoff,
            age_basis_effective: age_basis,
            years_num: years,
        };
        send_summary_email(
            email_rt,
            script_version,
            "(All published disk shares)",
            &archive_resolved.to_string_lossy(),
            years,
            age_basis,
            false,
            &empty_job,
            Some(listed),
            &html_paths,
        )?;
    }

    Ok(())
}

fn send_summary_email(
    email_rt: &EmailRuntime,
    version: &str,
    input_resolved: &str,
    archive_resolved: &str,
    years: f64,
    age_basis: AgeBasis,
    do_commit: bool,
    job: &JobResult,
    listed_override: Option<usize>,
    html_paths: &[PathBuf],
) -> anyhow::Result<()> {
    let domain = html_report::report_domain_label();
    let scan_lbl = if input_resolved == "(All published disk shares)" {
        "AllShares"
    } else {
        Path::new(input_resolved)
            .file_name()
            .and_then(|n| n.to_str())
            .unwrap_or("Source")
    };
    let mode = if do_commit { "Commit" } else { "Preview" };
    let stamp = chrono::Local::now().format("%Y-%m-%d %H:%M");
    let subject = format!(
        "[Archive-OldFiles] {} - {} - {} - {}",
        domain, scan_lbl, mode, stamp
    );

    let listed = listed_override.unwrap_or(job.results.len());
    let body = format!(
        "Archive Old Files (v{})\nComputer: {}\nDomain label: {}\nInput / scan: {}\nArchive: {}\nYears: {}  AgeBasis: {}  Output: {}  Commit: {}\nFiles scanned (all ages): {}\nMet age rule (listed): {}\nPlanned+Moved: {}  Failed: {}\nReclaim (Planned+Moved): {}",
        version,
        std::env::var("COMPUTERNAME").unwrap_or_default(),
        domain,
        input_resolved,
        archive_resolved,
        years,
        age_basis.as_str(),
        job.out_format,
        do_commit,
        job.file_scan_count,
        listed,
        job.planned_moved_count,
        job.failed_count,
        job.reclaim_display
    );

    let attach: Vec<PathBuf> = html_paths.to_vec();
    if attach.is_empty() {
        if let Some(ref p) = job.html_report_path {
            if p.exists() {
                send_archive_email(email_rt, &subject, &(body + "\n\n(HTML attached.)"), &[p])?;
                println!("Email sent to {}", email_rt.to);
                return Ok(());
            }
        }
        send_archive_email(
            email_rt,
            &subject,
            &(body + "\n\n(No HTML file attached for this run.)"),
            &[] as &[PathBuf],
        )?;
    } else {
        send_archive_email(email_rt, &subject, &body, &attach)?;
    }
    println!("Email sent to {}", email_rt.to);
    Ok(())
}

fn run_test_email(cli: &Cli, config_file: &Path) -> anyhow::Result<()> {
    let saved = read_config(config_file)?;
    let mut e = email_from_json(saved.as_ref().and_then(|c| c.email.as_ref()));
    merge_email_cli_only(&mut e, cli)?;

    let mut fails = Vec::new();
    if e.to.trim().is_empty() {
        fails.push("Recipient (To) is missing.");
    }
    if e.from.trim().is_empty() {
        fails.push("From is missing.");
    }
    if e.smtp_host.trim().is_empty() {
        fails.push("SMTP host is missing.");
    }
    if e.password_protected_base64.is_none() {
        fails.push("SMTP password is missing in config.");
    }
    if e.profile == "Smtp2Go" && e.user_name.trim().is_empty() {
        fails.push("SMTP2GO requires UserName.");
    }
    if !fails.is_empty() {
        anyhow::bail!("Cannot send test email:\n- {}", fails.join("\n- "));
    }

    let domain = html_report::report_domain_label();
    let subject = format!(
        "[Archive Old Files] SMTP test - {} (v{})",
        domain,
        crate::cli::VERSION
    );
    let body = format!(
        "Test from archive-old-files --test-email.\nComputer: {}\nSMTP: {}:{} TLS: {} User: {}\nConfig: {}",
        std::env::var("COMPUTERNAME").unwrap_or_default(),
        e.smtp_host,
        e.smtp_port,
        e.use_ssl,
        if e.user_name.is_empty() { "(none)" } else { &e.user_name },
        config_file.display()
    );
    send_archive_email(&e, &subject, &body, &[] as &[PathBuf])?;
    println!("Test email sent to {}.", e.to);
    Ok(())
}

fn merge_email_cli_only(e: &mut EmailRuntime, cli: &Cli) -> anyhow::Result<()> {
    if let Some(ref t) = cli.email_to {
        if !t.is_empty() {
            e.to = t.clone();
            e.enabled = true;
        }
    }
    if let Some(ref f) = cli.email_from {
        if !f.is_empty() {
            e.from = f.clone();
        }
    }
    if let Some(ref p) = cli.smtp_profile {
        if !p.is_empty() {
            e.profile = normalize_smtp_profile(p);
            apply_smtp_preset(&e.profile, &mut e.smtp_host, &mut e.smtp_port, &mut e.use_ssl);
        }
    }
    if let Some(ref h) = cli.smtp_host {
        if !h.is_empty() {
            e.smtp_host = h.clone();
        }
    }
    if let Some(p) = cli.smtp_port {
        e.smtp_port = p;
    }
    if let Some(ref s) = cli.smtp_use_ssl {
        e.use_ssl = bound_bool(Some(s.as_str()), e.use_ssl);
    }
    if let Some(ref u) = cli.smtp_user {
        if !u.is_empty() {
            e.user_name = u.clone();
        }
    }
    Ok(())
}

fn prompt_missing_paths_and_options(
    input: &mut String,
    archive: &mut String,
    years: &mut f64,
    do_commit: &mut bool,
    out_pref: &mut String,
    remove_pref: &mut String,
    age_basis: &mut AgeBasis,
    effective_archive_share: &str,
    saved: &Option<crate::config::AppConfig>,
    cli: &Cli,
) -> anyhow::Result<()> {
    if !cli.all && input.trim().is_empty() {
        if !cli.skip_share_menu {
            *input = resolve_input_via_share_menu()?;
        } else {
            let def = saved.as_ref().map(|s| s.input_path.as_str()).unwrap_or("");
            let p: String = Input::new()
                .with_prompt("InputPath")
                .default(def.to_string())
                .interact_text()?;
            *input = normalize_user_path(&p);
        }
    }

    if archive.trim().is_empty() {
        if let Some(p) = local_path_for_share_name(effective_archive_share) {
            *archive = p;
            println!(
                "Using archive share '{}' -> {}",
                effective_archive_share, archive
            );
        } else {
            println!(
                "No share named '{}' was found. Specify the archive folder manually.",
                effective_archive_share
            );
            let def = saved.as_ref().map(|s| s.archive_path.as_str()).unwrap_or("");
            let p: String = Input::new()
                .with_prompt("ArchivePath")
                .default(def.to_string())
                .interact_text()?;
            *archive = normalize_user_path(&p);
        }
    }

    if *years <= 0.0 {
        let def_y = saved.as_ref().and_then(|s| s.years).unwrap_or(0.0);
        let p: String = Input::new()
            .with_prompt("Years")
            .default(if def_y > 0.0 {
                def_y.to_string()
            } else {
                String::new()
            })
            .interact_text()?;
        if let Ok(y) = parse_years(&p) {
            if y > 0.0 {
                *years = y;
            }
        } else if def_y > 0.0 {
            *years = def_y;
        }
    }

    if !cli.commit && !cli.all {
        let def_c = saved.as_ref().and_then(|s| s.commit).unwrap_or(false);
        *do_commit = Confirm::new()
            .with_prompt("Run with --commit (actually move files)?")
            .default(def_c)
            .interact()?;
    }

    if cli.output.is_none() && !cli.all {
        let def_o = saved
            .as_ref()
            .and_then(|s| normalize_output_format(s.output.as_deref()))
            .unwrap_or_default();
        let p: String = Input::new()
            .with_prompt("Default report format Text or HTML (blank = prompt at end of run)")
            .default(def_o)
            .interact_text()?;
        let t = p.trim();
        if t.is_empty() {
            *out_pref = def_o;
        } else if let Some(n) = normalize_output_format(Some(t)) {
            *out_pref = n;
        }
    }

    if cli.remove_empty_folders.is_none() && *do_commit && !cli.all {
        let def_rm = saved
            .as_ref()
            .and_then(|s| s.remove_empty_folders.clone())
            .filter(|r| r.eq_ignore_ascii_case("yes") || r.eq_ignore_ascii_case("no"))
            .unwrap_or_default();
        let p: String = Input::new()
            .with_prompt("Remove empty folders after commit Yes/No (blank = prompt later)")
            .default(def_rm)
            .interact_text()?;
        let t = p.trim();
        if t.is_empty() {
            *remove_pref = def_rm;
        } else if t.eq_ignore_ascii_case("y") || t.eq_ignore_ascii_case("yes") {
            *remove_pref = "Yes".into();
        } else {
            *remove_pref = "No".into();
        }
    }

    if cli.age_basis.is_none() {
        let def_ab = saved
            .as_ref()
            .and_then(|s| parse_age_basis_saved(s.age_basis.as_deref()))
            .unwrap_or(AgeBasis::LastWriteTime);
        println!();
        println!("Age basis: A=Modified B=Opened C=Max(mod,access) CreationTime Earliest");
        let r: String = Input::new()
            .with_prompt("AgeBasis")
            .default(def_ab.as_str().to_string())
            .interact_text()?;
        if !r.trim().is_empty() {
            if let Some(n) = normalize_age_basis_param(r.trim()) {
                *age_basis = n;
            }
        } else {
            *age_basis = def_ab;
        }
    }

    Ok(())
}
