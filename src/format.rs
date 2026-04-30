use std::path::Path;

/// Supported document formats.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default, serde::Serialize, serde::Deserialize)]
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
    #[must_use]
    pub fn from_extension(ext: &str) -> Option<Self> {
        match ext.to_ascii_lowercase().as_str() {
            "docx" => Some(Self::Docx),
            "xlsx" => Some(Self::Xlsx),
            "pptx" => Some(Self::Pptx),
            "doc" => Some(Self::Doc),
            "xls" => Some(Self::Xls),
            "ppt" => Some(Self::Ppt),
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
    fn from_extension() {
        assert_eq!(DocumentFormat::from_extension("docx"), Some(DocumentFormat::Docx));
        assert_eq!(DocumentFormat::from_extension("XLSX"), Some(DocumentFormat::Xlsx));
        assert_eq!(DocumentFormat::from_extension("pptx"), Some(DocumentFormat::Pptx));
        assert_eq!(DocumentFormat::from_extension("doc"), Some(DocumentFormat::Doc));
        assert_eq!(DocumentFormat::from_extension("XLS"), Some(DocumentFormat::Xls));
        assert_eq!(DocumentFormat::from_extension("ppt"), Some(DocumentFormat::Ppt));
        assert_eq!(DocumentFormat::from_extension("txt"), None);
        assert_eq!(DocumentFormat::from_extension("pdf"), None);
    }

    #[test]
    fn from_path() {
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
