//! Convenience wrappers for command-line invocation of Outlook.
//!
//! # Example
//!
//! Basic usage:
//!
//! ```rust,no_run
//! use outlook_exe;
//!
//! outlook_exe::MessageBuilder::new()
//!     .with_recipient("noreply@example.org")
//!     .with_subject("Hello, World!")
//!     .with_body("Line with spaces\nAnother line")
//!     .with_attachment("C:/tmp/file.txt")
//!     .spawn()
//!     .unwrap();
//! ```

use std::{io, process};

#[macro_use]
extern crate lazy_static;

lazy_static! {
    static ref OUTLOOK_EXE: Option<&'static str> = {
        use winreg::{enums::HKEY_LOCAL_MACHINE, RegKey};

        const OUTLOOK_SUBKEY: &str =
            "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\OUTLOOK.EXE";

        let subkey = match RegKey::predef(HKEY_LOCAL_MACHINE).open_subkey(OUTLOOK_SUBKEY) {
            Ok(subkey) => subkey,
            Err(_) => return None,
        };
        let value: String = match subkey.get_value("") {
            Ok(value) => value,
            Err(_) => return None,
        };
        Some(Box::leak(value.into_boxed_str()))
    };
}

fn percent_escape(s: &str) -> String {
    s.replace('%', "%25") // has to be first to avoid double-encoding '%'
        .replace('"', "%22")
        .replace('&', "%26")
        .replace('?', "%3F")
}

/// The `MessageBuilder` type, for drafting Outlook email messages.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct MessageBuilder {
    subj: String,
    to: Vec<String>,
    cc: Vec<String>,
    bcc: Vec<String>,
    body: String,
    file: String,
}

impl MessageBuilder {
    /// Creates a new `MessageBuilder`.
    #[inline]
    #[must_use]
    pub const fn new() -> Self {
        Self {
            subj: String::new(),
            to: Vec::new(),
            cc: Vec::new(),
            bcc: Vec::new(),
            body: String::new(),
            file: String::new(),
        }
    }

    /// Adds a subject to the email.
    ///
    /// This should only be called once per `MessageBuilder` instance.
    #[inline]
    #[must_use]
    pub fn with_subject<S>(self, subj: S) -> Self
    where
        S: Into<String>,
    {
        debug_assert!(self.subj.is_empty(), "Outlook subject already provided");
        Self {
            subj: subj.into(),
            to: self.to,
            cc: self.cc,
            bcc: self.bcc,
            body: self.body,
            file: self.file,
        }
    }

    /// Adds a recipient to the email.
    #[inline]
    #[must_use]
    pub fn with_recipient<S>(mut self, to: S) -> Self
    where
        S: Into<String>,
    {
        self.to.push(to.into());
        Self {
            subj: self.subj,
            to: self.to,
            cc: self.cc,
            bcc: self.bcc,
            body: self.body,
            file: self.file,
        }
    }

    /// Adds a CC recipient to the email.
    #[inline]
    #[must_use]
    pub fn with_recipient_cc<S>(mut self, cc: S) -> Self
    where
        S: Into<String>,
    {
        self.cc.push(cc.into());
        Self {
            subj: self.subj,
            to: self.to,
            cc: self.cc,
            bcc: self.bcc,
            body: self.body,
            file: self.file,
        }
    }

    /// Adds a BCC recipient to the email.
    #[inline]
    #[must_use]
    pub fn with_recipient_bcc<S>(mut self, bcc: S) -> Self
    where
        S: Into<String>,
    {
        self.bcc.push(bcc.into());
        Self {
            subj: self.subj,
            to: self.to,
            cc: self.cc,
            bcc: self.bcc,
            body: self.body,
            file: self.file,
        }
    }

    /// Adds a body to the email.
    ///
    /// This should only be called once per `MessageBuilder` instance.
    #[inline]
    #[must_use]
    pub fn with_body<S>(self, body: S) -> Self
    where
        S: Into<String>,
    {
        debug_assert!(self.body.is_empty(), "Outlook body already provided");
        Self {
            subj: self.subj,
            to: self.to,
            cc: self.cc,
            bcc: self.bcc,
            body: body.into(),
            file: self.file,
        }
    }

    /// Adds an attachment to the email.
    ///
    /// This should only be called once per `MessageBuilder` instance,
    /// because Outlook's command-line switches only supports attaching
    /// a single file per invocation.
    #[inline]
    #[must_use]
    pub fn with_attachment<S>(self, file: S) -> Self
    where
        S: Into<String>,
    {
        debug_assert!(
            self.file.is_empty(),
            "Outlook's invocation switches do not support attaching multiple files"
        );
        Self {
            subj: self.subj,
            to: self.to,
            cc: self.cc,
            bcc: self.bcc,
            body: self.body,
            file: file.into(),
        }
    }

    /// Spawns an Outlook process, and prompts the user to press "Send".
    ///
    /// # Errors
    ///
    /// Will return `Err(io::Error)` if OUTLOOK.EXE cannot
    /// be located, or if a child process cannot be spawned.
    pub fn spawn(mut self) -> io::Result<process::Child> {
        let mut s = String::new();
        s.push_str(&percent_escape(&self.to.join(";")));
        if !self.cc.is_empty() {
            if !s.is_empty() {
                s.push('&')
            }
            s.push_str("cc=");
            s.push_str(&percent_escape(&self.cc.join(";")));
        }
        if !self.bcc.is_empty() {
            if !s.is_empty() {
                s.push('&')
            }
            s.push_str("bcc=");
            s.push_str(&percent_escape(&self.bcc.join(";")));
        }
        if !self.subj.is_empty() {
            if !s.is_empty() {
                s.push('&')
            }
            s.push_str("subject=");
            s.push_str(&percent_escape(&self.subj));
        }
        if !self.body.is_empty() {
            if !s.is_empty() {
                s.push('&')
            }
            s.push_str("body=");
            s.push_str(&percent_escape(&self.body));
        }
        let mut a = Vec::new();
        if !self.file.is_empty() {
            a.push("/a");
            self.file = percent_escape(&self.file);
            a.push(&self.file);
        }
        let outlook_exe =
            OUTLOOK_EXE.ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "OUTLOOK.EXE"))?;
        process::Command::new(outlook_exe)
            .arg("/c")
            .arg("ipm.note")
            .arg("/m")
            .arg(s)
            .args(a)
            .spawn()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn message_builder() {
        // A dumb test
        let mb = MessageBuilder::new();
        assert_eq!(mb.to.len(), 0);
        assert_eq!(mb.cc.len(), 0);
        assert_eq!(mb.bcc.len(), 0);
        assert_eq!(mb.subj, "");
        assert_eq!(mb.body, "");
        assert_eq!(mb.file, "");
        let mb = mb
            .with_recipient("noreply@example.org")
            .with_subject("Hello, World!")
            .with_body("Line with spaces\nAnother line")
            .with_attachment("C:/tmp/file.txt");
        assert_eq!(mb.to.len(), 1);
        assert_eq!(mb.cc.len(), 0);
        assert_eq!(mb.bcc.len(), 0);
        assert_eq!(mb.to[0], "noreply@example.org");
        assert_eq!(mb.subj, "Hello, World!");
        assert_eq!(mb.body, "Line with spaces\nAnother line");
        assert_eq!(mb.file, "C:/tmp/file.txt");
    }
}
