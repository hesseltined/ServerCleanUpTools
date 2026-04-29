//! SMTP send (lettre), password via Windows DPAPI.

use crate::config::EmailRuntime;
use crate::win_dpapi::unprotect_current_user;
use lettre::message::header::ContentType;
use lettre::message::{Attachment, Body, Mailbox, Message, MultiPart, SinglePart};
use lettre::transport::smtp::authentication::Credentials;
use lettre::{SmtpTransport, Transport};
use std::path::Path;

pub fn send_archive_email(
    settings: &EmailRuntime,
    subject: &str,
    body: &str,
    attachment_paths: &[impl AsRef<Path>],
) -> anyhow::Result<()> {
    let plain_pass = match &settings.password_protected_base64 {
        Some(b64) => unprotect_current_user(b64)?,
        None => String::new(),
    };
    if plain_pass.is_empty() {
        anyhow::bail!("SMTP password missing or could not be decrypted.");
    }

    let from: Mailbox = settings.from.trim().parse()?;

    let mut multi = MultiPart::mixed().singlepart(
        SinglePart::builder()
            .header(ContentType::TEXT_PLAIN)
            .body(Body::new(body.to_string())),
    );

    for ap in attachment_paths {
        let p = ap.as_ref();
        if p.exists() {
            let data = std::fs::read(p)?;
            let ct = ContentType::parse("application/octet-stream")?;
            let filename = p
                .file_name()
                .and_then(|n| n.to_str())
                .unwrap_or("attachment")
                .to_string();
            multi = multi.singlepart(Attachment::new(filename).body(data, ct));
        }
    }

    let mut builder = Message::builder().from(from).subject(subject);
    for addr in settings.to.split(',') {
        let a = addr.trim();
        if !a.is_empty() {
            let m: Mailbox = a.parse()?;
            builder = builder.to(m);
        }
    }
    let email = builder.multipart(multi)?;

    let creds = if settings.user_name.trim().is_empty() {
        None
    } else {
        Some(Credentials::new(
            settings.user_name.trim().to_string(),
            plain_pass,
        ))
    };

    let host = settings.smtp_host.trim();

    // Port 587 / STARTTLS: starttls_relay. Plain SMTP: builder_dangerous (UseSsl false in JSON).
    let transport = if settings.use_ssl {
        let mut b = SmtpTransport::starttls_relay(host)?.port(settings.smtp_port);
        if let Some(c) = creds {
            b = b.credentials(c);
        }
        b.build()
    } else {
        let mut b = SmtpTransport::builder_dangerous(host).port(settings.smtp_port);
        if let Some(c) = creds {
            b = b.credentials(c);
        }
        b.build()
    };

    transport.send(&email)?;
    Ok(())
}
