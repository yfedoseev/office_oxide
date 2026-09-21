use std::path::Path;

/// Supported document formats.
#[derive(
    Debug, Clone, Copy, PartialEq, Eq, Hash, Default, serde::Serialize, serde::Deserialize,
)]
#[serde(rename_all = "lowercase")]
pub enum DocumentFormat {
    #[default]
    /// Office Open XML Word document (`.docx`).
    Docx,
    /// Office Open XML Excel spreadsheet (`.xlsx`).
    Xlsx,
    /// Office Open XML PowerPoint presentation (`.pptx`).
    Pptx,
    /// Legacy Word Binary document (`.doc`).
    Doc,
    /// Legacy Excel Binary workbook (`.xls`).
    Xls,
    /// Legacy PowerPoint Binary presentation (`.ppt`).
    Ppt,
}

impl DocumentFormat {
    /// Returns the canonical file extension (without the leading dot).
    #[must_use]
    pub fn extension(&self) -> &'static str {
        match self {
            Self::Docx => "docx",
            Self::Xlsx => "xlsx",
            Self::Pptx => "pptx",
            Self::Doc => "doc",
            Self::Xls => "xls",
            Self::Ppt => "ppt",
        }
    }

    /// Returns the canonical MIME type.
    #[must_use]
    pub fn mime_type(&self) -> &'static str {
        match self {
            Self::Docx => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            Self::Xlsx => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            Self::Pptx => {
                "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            },
            Self::Doc => "application/msword",
            Self::Xls => "application/vnd.ms-excel",
            Self::Ppt => "application/vnd.ms-powerpoint",
        }
    }
}

impl DocumentFormat {
    /// Detect format from a file extension string (case-insensitive, without dot).
    ///
    /// Macro-enabled and template extensions map to their base OOXML
    /// format: they are the same package layout with a different content
    /// type, which the readers already accept. Rejecting them here made
    /// `Document::open("report.docm")` fail with `UnsupportedFormat`
    /// before any reader ran.
    #[must_use]
    pub fn from_extension(ext: &str) -> Option<Self> {
        match ext.to_ascii_lowercase().as_str() {
            // WordprocessingML: document, macro-enabled document,
            // template, macro-enabled template.
            "docx" | "docm" | "dotx" | "dotm" => Some(Self::Docx),
            // SpreadsheetML: workbook, macro-enabled workbook, templates,
            // macro-enabled add-in.
            "xlsx" | "xlsm" | "xltx" | "xltm" | "xlam" | "xlsb" => Some(Self::Xlsx),
            // PresentationML: presentation, macro-enabled presentation,
            // templates, slideshows.
            "pptx" | "pptm" | "potx" | "potm" | "ppsx" | "ppsm" => Some(Self::Pptx),
            // The legacy compound files: document and template; workbook,
            // template and add-in; presentation, template and show. Each
            // extension is the same container as its sibling — Word 97
            // writes a `.dot` exactly as it writes a `.doc`.
            "doc" | "dot" => Some(Self::Doc),
            "xls" | "xlt" | "xla" => Some(Self::Xls),
            "ppt" | "pot" | "pps" => Some(Self::Ppt),
            _ => None,
        }
    }

    /// Detect format from a file path's extension.
    #[must_use]
    pub fn from_path(path: &Path) -> Option<Self> {
        let ext = path.extension()?.to_str()?;
        Self::from_extension(ext)
    }

    /// If this is a legacy format, return the corresponding OOXML format.
    /// Used when magic bytes reveal the file is actually OOXML despite the extension.
    #[must_use]
    pub fn ooxml_upgrade(&self) -> Option<Self> {
        match self {
            Self::Doc => Some(Self::Docx),
            Self::Xls => Some(Self::Xlsx),
            Self::Ppt => Some(Self::Pptx),
            _ => None,
        }
    }

    /// Returns true if this is a legacy binary format (doc/xls/ppt).
    #[must_use]
    pub fn is_legacy(&self) -> bool {
        matches!(self, Self::Doc | Self::Xls | Self::Ppt)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_from_extension() {
        assert_eq!(DocumentFormat::from_extension("docx"), Some(DocumentFormat::Docx));
        assert_eq!(DocumentFormat::from_extension("XLSX"), Some(DocumentFormat::Xlsx));
        assert_eq!(DocumentFormat::from_extension("pptx"), Some(DocumentFormat::Pptx));
        assert_eq!(DocumentFormat::from_extension("doc"), Some(DocumentFormat::Doc));
        assert_eq!(DocumentFormat::from_extension("XLS"), Some(DocumentFormat::Xls));
        assert_eq!(DocumentFormat::from_extension("ppt"), Some(DocumentFormat::Ppt));
        assert_eq!(DocumentFormat::from_extension("txt"), None);
        assert_eq!(DocumentFormat::from_extension("pdf"), None);
    }

    /// The legacy template/show/add-in extensions were missed when the
    /// OOXML ones were added: a Word 97 template (`.dot`), a PowerPoint
    /// show (`.pps`) and an Excel add-in (`.xla`) failed with
    /// `unsupported format` although each is byte-for-byte the same
    /// container as its `.doc`/`.ppt`/`.xls` sibling.
    #[test]
    fn test_from_extension_legacy_templates_shows_and_add_ins() {
        for (ext, want) in [
            ("dot", DocumentFormat::Doc),
            ("DOT", DocumentFormat::Doc),
            ("xlt", DocumentFormat::Xls),
            ("xla", DocumentFormat::Xls),
            ("pot", DocumentFormat::Ppt),
            ("pps", DocumentFormat::Ppt),
        ] {
            assert_eq!(DocumentFormat::from_extension(ext), Some(want), "{ext}");
        }
    }

    /// Macro-enabled and template extensions used to return `None`, so
    /// `Document::open("file.docm")` failed with `UnsupportedFormat`
    /// before any reader saw the bytes.
    #[test]
    fn test_from_extension_macro_enabled_and_templates() {
        for ext in ["docm", "dotx", "dotm", "DOCM"] {
            assert_eq!(
                DocumentFormat::from_extension(ext),
                Some(DocumentFormat::Docx),
                "{ext} should map to Docx"
            );
        }
        for ext in ["xlsm", "xltx", "xltm", "xlam", "XLSM"] {
            assert_eq!(
                DocumentFormat::from_extension(ext),
                Some(DocumentFormat::Xlsx),
                "{ext} should map to Xlsx"
            );
        }
        for ext in ["pptm", "potx", "potm", "ppsx", "ppsm", "PPTM"] {
            assert_eq!(
                DocumentFormat::from_extension(ext),
                Some(DocumentFormat::Pptx),
                "{ext} should map to Pptx"
            );
        }
        // Still not a document format.
        assert_eq!(DocumentFormat::from_extension("docz"), None);
        assert_eq!(
            DocumentFormat::from_path(Path::new("macros/book.xlsm")),
            Some(DocumentFormat::Xlsx)
        );
    }

    #[test]
    fn test_from_path() {
        assert_eq!(DocumentFormat::from_path(Path::new("report.docx")), Some(DocumentFormat::Docx));
        assert_eq!(
            DocumentFormat::from_path(Path::new("/tmp/data.xlsx")),
            Some(DocumentFormat::Xlsx)
        );
        assert_eq!(DocumentFormat::from_path(Path::new("slides.PPTX")), Some(DocumentFormat::Pptx));
        assert_eq!(DocumentFormat::from_path(Path::new("notes.txt")), None);
        assert_eq!(DocumentFormat::from_path(Path::new("noext")), None);
    }
}
