use std::path::Path;

/// Supported document formats.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocumentFormat {
    Docx,
    Xlsx,
    Pptx,
}

impl DocumentFormat {
    /// Detect format from a file extension string (case-insensitive, without dot).
    pub fn from_extension(ext: &str) -> Option<Self> {
        match ext.to_ascii_lowercase().as_str() {
            "docx" => Some(Self::Docx),
            "xlsx" => Some(Self::Xlsx),
            "pptx" => Some(Self::Pptx),
            _ => None,
        }
    }

    /// Detect format from a file path's extension.
    pub fn from_path(path: &Path) -> Option<Self> {
        let ext = path.extension()?.to_str()?;
        Self::from_extension(ext)
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
        assert_eq!(DocumentFormat::from_extension("txt"), None);
        assert_eq!(DocumentFormat::from_extension("pdf"), None);
    }

    #[test]
    fn from_path() {
        assert_eq!(
            DocumentFormat::from_path(Path::new("report.docx")),
            Some(DocumentFormat::Docx)
        );
        assert_eq!(
            DocumentFormat::from_path(Path::new("/tmp/data.xlsx")),
            Some(DocumentFormat::Xlsx)
        );
        assert_eq!(
            DocumentFormat::from_path(Path::new("slides.PPTX")),
            Some(DocumentFormat::Pptx)
        );
        assert_eq!(DocumentFormat::from_path(Path::new("notes.txt")), None);
        assert_eq!(DocumentFormat::from_path(Path::new("noext")), None);
    }
}
