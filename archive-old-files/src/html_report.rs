//! HTML report generation (layout/CSS aligned with PowerShell script).

use crate::age_basis::AgeBasis;
use chrono::{DateTime, Local};
use std::collections::BTreeMap;
use std::fs;
use std::path::Path;

#[derive(Clone, Debug)]
pub struct ReportRow {
    pub source_path: String,
    pub compared_for_age: DateTime<Local>,
    pub length: u64,
    pub owner: String,
    pub status: String,
    pub message: String,
}

fn html_encode(text: &str) -> String {
    text.chars()
        .map(|c| match c {
            '&' => "&amp;".to_string(),
            '<' => "&lt;".to_string(),
            '>' => "&gt;".to_string(),
            '"' => "&quot;".to_string(),
            '\'' => "&#39;".to_string(),
            _ => c.to_string(),
        })
        .collect()
}

fn sanitize_filename_part(s: &str, max: usize) -> String {
    let invalid: &[char] = &['/', '\\', ':', '*', '?', '"', '<', '>', '|'];
    let mut d: String = s
        .chars()
        .map(|c| {
            if invalid.contains(&c) {
                '_'
            } else {
                c
            }
        })
        .collect();
    d = d.split_whitespace().collect::<Vec<_>>().join("_");
    d = d.trim_matches(|c| c == '.' || c == '_').to_string();
    if d.is_empty() {
        return "PART".to_string();
    }
    if d.len() > max {
        d.truncate(max);
    }
    d
}

pub fn report_domain_label() -> String {
    #[cfg(windows)]
    {
        if let Some(d) = crate::win_domain::computer_system_domain() {
            if !d.is_empty() && !d.eq_ignore_ascii_case("WORKGROUP") {
                return d;
            }
        }
        if let Ok(u) = std::env::var("USERDNSDOMAIN") {
            let t = u.trim();
            if !t.is_empty() {
                return t.to_string();
            }
        }
    }
    std::env::var("COMPUTERNAME")
        .unwrap_or_else(|_| "HOST".to_string())
        .trim()
        .to_string()
}

pub fn archive_html_report_filename(
    domain_part: &str,
    input_resolved: &str,
    share_name_label: Option<&str>,
) -> String {
    let mut d = sanitize_filename_part(domain_part, 48);
    if d.is_empty() {
        d = "DOMAIN".to_string();
    }

    let mut share_part: String = input_resolved
        .trim_end_matches(['\\', '/'])
        .chars()
        .map(|c| {
            if ['\\', ':'].contains(&c) {
                if c == '\\' {
                    '-'
                } else {
                    '_'
                }
            } else if ['/', '*', '?', '"', '<', '>', '|'].contains(&c) {
                '_'
            } else {
                c
            }
        })
        .collect();
    while share_part.contains("--") {
        share_part = share_part.replace("--", "-");
    }
    share_part = share_part.trim_matches('-').to_string();
    if share_part.is_empty() {
        share_part = "Source".to_string();
    }
    if share_part.len() > 100 {
        share_part.truncate(100);
        share_part = share_part.trim_end_matches('-').to_string();
    }

    if let Some(sn) = share_name_label {
        let mut sn2 = sanitize_filename_part(sn.trim(), 48);
        if sn2.is_empty() {
            sn2 = "Share".to_string();
        }
        share_part = format!("{}-{}", sn2, share_part);
    }

    let stamp = Local::now().format("%Y%m%d_%H%M%S");
    format!("ArchiveReport_{}_{}_{}.html", d, share_part, stamp)
}

const REPORT_CSS: &str = r#"body { font-family: Segoe UI, Roboto, Helvetica, Arial, sans-serif; margin: 0; background: #eef1f5; color: #1a1a1a; font-size: 13px; }
header { background: linear-gradient(135deg, #0d3b66 0%, #1b6ca8 100%); color: #fff; padding: 0.65rem 1rem 0.85rem; }
header h1 { margin: 0 0 0.4rem 0; font-size: 1.1rem; font-weight: 600; }
.head-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 0.5rem 1.25rem; align-items: start; font-size: 0.8rem; }
.head-grid .col h3 { margin: 0 0 0.35rem 0; font-size: 0.72rem; text-transform: uppercase; letter-spacing: 0.04em; opacity: 0.85; font-weight: 600; }
.head-grid dl { display: grid; grid-template-columns: 7.2rem 1fr; gap: 0.15rem 0.5rem; margin: 0; }
.head-grid dt { opacity: 0.88; font-weight: 500; }
.head-grid dd { margin: 0; word-break: break-word; line-height: 1.25; }
.wrap { max-width: 1200px; margin: 0 auto; padding: 0.75rem 1rem 1.25rem; }
.tip-top { font-size: 0.78rem; color: #334; background: #e3ecf7; border: 1px solid #c5d4e8; border-radius: 6px; padding: 0.5rem 0.65rem; margin: 0 0 0.65rem 0; line-height: 1.4; }
.card { background: #fff; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,.1); margin-bottom: 0.75rem; overflow: hidden; }
.card h2 { margin: 0; padding: 0.45rem 0.75rem; font-size: 0.88rem; background: #e8eef4; border-bottom: 1px solid #d0dae6; }
.card .body { padding: 0.5rem 0.65rem; }
table { width: 100%; border-collapse: collapse; font-size: 0.78rem; table-layout: fixed; }
th { text-align: left; background: #0d3b66; color: #fff; padding: 0.35rem 0.45rem; font-weight: 600; }
th:nth-child(1) { width: 36%; }
th:nth-child(2) { width: 22%; }
th:nth-child(3) { width: 18%; }
th:nth-child(4) { width: 24%; }
td { padding: 0.28rem 0.45rem; vertical-align: middle; border-bottom: 1px solid #dde3ea; line-height: 1.2; }
tbody.folder-group { border-bottom: 2px solid #b8c5d4; }
tbody.folder-group.collapsible { cursor: default; background: #fdf6ec; border-left: 4px solid #c9943a; }
tbody.folder-group.collapsible tr.folder-hdr td { background: #e8d4b0; color: #3d3010; border-bottom: 1px solid #d4b78a; }
tbody.folder-group.collapsible tr.frow.stripe0 td { background: #faf3e6; }
tbody.folder-group.collapsible tr.frow.stripe1 td { background: #f3e8d2; }
tr.folder-hdr td { background: #d5dee8; font-weight: 600; color: #0d3b66; padding: 0.4rem 0.45rem; border-bottom: 1px solid #b8c5d4; font-size: 0.76rem; word-break: break-word; }
tr.frow.stripe0 td { background: #f4f5f7; }
tr.frow.stripe1 td { background: #e8f1fb; }
tr.frow.failed td { background: #f5d4d4 !important; color: #6b1212; font-weight: 500; }
td.fn { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
td.own { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: #333; font-size: 0.74rem; }
td.dt { white-space: nowrap; color: #333; }
td.num { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
.empty { text-align: center; color: #666; padding: 1rem !important; }"#;

const REPORT_JS: &str = r#"<script>
(function () {
  document.querySelectorAll('tbody.folder-group.collapsible').forEach(function (tb) {
    function showAll() {
      tb.querySelectorAll('tr.frow').forEach(function (tr) { tr.style.display = 'table-row'; });
    }
    function hideFiles() {
      tb.querySelectorAll('tr.frow').forEach(function (tr) { tr.style.display = 'none'; });
    }
    tb.addEventListener('mouseenter', showAll);
    tb.addEventListener('mouseleave', hideFiles);
  });
})();
</script>"#;

#[allow(clippy::too_many_arguments)]
pub fn write_archive_html_report(
    path: &Path,
    rows: &[ReportRow],
    input_resolved: &str,
    archive_resolved: &str,
    years: f64,
    age_basis: AgeBasis,
    cutoff: DateTime<Local>,
    commit: bool,
    script_version: &str,
    file_scan_count: usize,
    skipped_too_new: usize,
    planned_moved_count: usize,
    failed_count: usize,
    reclaim_bytes: u64,
    reclaim_display: &str,
    domain_label: &str,
) -> anyhow::Result<()> {
    let gen_at = Local::now().format("%x %X").to_string();
    let mode_label = if commit {
        "Commit (moves performed where successful)"
    } else {
        "Preview (no moves)"
    };
    let server_name = std::env::var("COMPUTERNAME").unwrap_or_else(|_| "(unknown)".to_string());

    let input_root = input_resolved.trim_end_matches(['\\', '/']);

    let mut body = String::new();
    if rows.is_empty() {
        body.push_str(
            r#"<tbody><tr><td colspan="4" class="empty">No files met the age rule for this run.</td></tr></tbody>"#,
        );
    } else {
        // Group by parent folder
        let mut groups: BTreeMap<String, Vec<&ReportRow>> = BTreeMap::new();
        for r in rows {
            let parent = std::path::Path::new(&r.source_path)
                .parent()
                .map(|p| p.to_string_lossy().to_string())
                .unwrap_or_else(|| input_root.to_string());
            groups.entry(parent).or_default().push(r);
        }

        for (folder_full, mut group_rows) in groups {
            let folder_display = if folder_full.is_empty() {
                input_root.to_string()
            } else {
                folder_full.trim_end_matches(['\\', '/']).to_string()
            };
            group_rows.sort_by(|a, b| a.source_path.cmp(&b.source_path));
            let fc = group_rows.len();
            let is_collapsible = fc > 20;
            let hdr_enc = html_encode(&folder_display);
            let hdr_extra = if is_collapsible {
                format!(" &mdash; {} files (collapsed; hover anywhere in this folder block to show all, move pointer away to hide)", fc)
            } else {
                format!(
                    " &mdash; {} file{}",
                    fc,
                    if fc != 1 { "s" } else { "" }
                )
            };
            let tbody_class = if is_collapsible {
                "folder-group collapsible"
            } else {
                "folder-group"
            };
            body.push_str(&format!(
                r#"<tbody class="{}">"#,
                html_encode(tbody_class)
            ));
            body.push_str(&format!(
                r#"<tr class="folder-hdr"><td colspan="4">{}{}</td></tr>"#,
                hdr_enc, hdr_extra
            ));

            let mut stripe = 0u32;
            for r in group_rows {
                let leaf = std::path::Path::new(&r.source_path)
                    .file_name()
                    .and_then(|n| n.to_str())
                    .unwrap_or(&r.source_path)
                    .to_string();
                let lwt = r.compared_for_age.format("%Y-%m-%d %H:%M");
                let len = format!("{}", r.length);
                let own = if r.owner.trim().is_empty() {
                    "-".to_string()
                } else {
                    r.owner.clone()
                };
                let is_fail = r.status == "Failed";
                let mut tip = r.source_path.clone();
                if is_fail && !r.message.is_empty() {
                    tip.push('\n');
                    tip.push_str("Error: ");
                    tip.push_str(&r.message);
                }
                let row_class = if is_fail {
                    "frow failed".to_string()
                } else {
                    let c = format!("frow stripe{}", stripe % 2);
                    stripe += 1;
                    c
                };
                let hide_style = if is_collapsible {
                    " style=\"display:none\""
                } else {
                    ""
                };
                body.push_str(&format!(
                    r#"<tr class="{}" title="{}"{}><td class="fn">{}</td><td class="own">{}</td><td class="dt">{}</td><td class="num">{}</td></tr>"#,
                    row_class,
                    html_encode(&tip),
                    hide_style,
                    html_encode(&leaf),
                    html_encode(&own),
                    lwt,
                    len
                ));
            }
            body.push_str("</tbody>");
        }
    }

    let date_hdr = html_encode(age_basis.report_column_header());
    let html = format!(
        r#"<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Archive Old Files Report</title>
<style>{}</style>
</head>
<body>
<header>
  <h1>Archive Old Files &mdash; Report</h1>
  <div class="head-grid">
    <div class="col job">
      <h3>Job</h3>
      <dl>
        <dt>Server</dt><dd>{}</dd>
        <dt>Domain</dt><dd>{}</dd>
        <dt>Source</dt><dd>{}</dd>
        <dt>Archive</dt><dd>{}</dd>
        <dt>Older than</dt><dd>{}</dd>
        <dt>Cutoff</dt><dd>{}</dd>
        <dt>Mode</dt><dd>{}</dd>
      </dl>
    </div>
    <div class="col results">
      <h3>Results</h3>
      <dl>
        <dt>Generated</dt><dd>{}</dd>
        <dt>Script</dt><dd>v{}</dd>
        <dt>Scanned</dt><dd>{} files</dd>
        <dt>Too new</dt><dd>{}</dd>
        <dt>Met age rule</dt><dd>{} total</dd>
        <dt>Planned / moved</dt><dd>{}</dd>
        <dt>Failed</dt><dd>{}</dd>
        <dt>Size off source</dt><dd><strong>{}</strong> ({} B)</dd>
      </dl>
    </div>
  </div>
</header>
<div class="wrap">
  <p class="tip-top">Hover a <strong>file name</strong> for the full source path. <strong>Red</strong> rows failed to move or plan; hover the row for the error. Folders with <strong>more than 20</strong> files start with file lines hidden; move the pointer into that folder&rsquo;s block (header or rows) to show them, and move away to collapse again.</p>
  <div class="card">
    <h2>Files (by folder)</h2>
    <div class="body" style="overflow-x:auto;">
      <table>
        <thead>
          <tr>
            <th>File</th>
            <th>Owner</th>
            <th>{}</th>
            <th>Size (bytes)</th>
          </tr>
        </thead>
{}
      </table>
    </div>
  </div>
</div>
{}
</body>
</html>"#,
        REPORT_CSS,
        html_encode(&server_name),
        html_encode(domain_label),
        html_encode(input_resolved),
        html_encode(archive_resolved),
        html_encode(&format!("{} yr, basis {}", years, age_basis.as_str())),
        html_encode(&cutoff.format("%Y-%m-%d %H:%M").to_string()),
        html_encode(mode_label),
        html_encode(&gen_at),
        script_version,
        file_scan_count,
        skipped_too_new,
        rows.len(),
        planned_moved_count,
        failed_count,
        html_encode(reclaim_display),
        reclaim_bytes,
        date_hdr,
        body,
        REPORT_JS
    );

    fs::write(path, html.as_bytes())?;
    Ok(())
}
