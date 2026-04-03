//! High-level DOC document API.

use std::io::{Read, Seek};

use cfb_oxide::CfbReader;

use crate::error::{DocError, Result};
use crate::fib::Fib;
use crate::images::{extract_images, DocImage};
use crate::piece_table::{extract_text, parse_clx, sanitize_text};

/// A parsed legacy Word document.
#[derive(Debug)]
pub struct DocDocument {
    /// The raw extracted text (after sanitization).
    text: String,
    /// Extracted images from the Data stream.
    images: Vec<DocImage>,
}

impl DocDocument {
    /// Open a DOC file from a reader.
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let mut cfb = CfbReader::new(reader)?;

        let word_doc = cfb.open_stream("WordDocument").map_err(|_| {
            DocError::MissingStream("WordDocument stream not found".into())
        })?;

        let fib = match Fib::parse(&word_doc) {
            Ok(f) => f,
            Err(_) => return Ok(Self { text: String::new(), images: Vec::new() }), // Unsupported Word version
        };

        // Open the appropriate table stream; try preferred first, then fallback.
        let table_stream = if fib.use_table1 {
            cfb.open_stream("1Table")
                .or_else(|_| cfb.open_stream("0Table"))
        } else {
            cfb.open_stream("0Table")
                .or_else(|_| cfb.open_stream("1Table"))
        };
        let table_stream = match table_stream {
            Ok(s) => s,
            Err(_) => return Ok(Self { text: String::new(), images: Vec::new() }), // Word 6/95 or corrupted
        };

        // Extract CLX from the table stream.
        let clx_start = fib.clx_offset as usize;
        let clx_end = clx_start + fib.clx_size as usize;

        if clx_start >= table_stream.len() || clx_size_zero_or_oob(fib.clx_size, clx_start, table_stream.len()) {
            // CLX not available — return empty document.
            return Ok(Self { text: String::new(), images: Vec::new() });
        }

        let clx_end = clx_end.min(table_stream.len());
        let clx_data = &table_stream[clx_start..clx_end];
        let pieces = match parse_clx(clx_data) {
            Ok(p) => p,
            Err(_) => return Ok(Self { text: String::new(), images: Vec::new() }),
        };

        // Extract main document text only (not footnotes, headers, etc.).
        let raw_text = extract_text(&word_doc, &pieces, fib.text_len);
        let text = sanitize_text(&raw_text);

        // Extract images from the Data stream (if present).
        let images = match cfb.open_stream("Data") {
            Ok(data_stream) => extract_images(&data_stream),
            Err(_) => Vec::new(),
        };

        Ok(Self { text, images })
    }

    /// Open a DOC file from a path.
    pub fn open<P: AsRef<std::path::Path>>(path: P) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Get all extracted images.
    pub fn images(&self) -> &[DocImage] {
        &self.images
    }

    /// Get the extracted plain text.
    pub fn plain_text(&self) -> String {
        self.text.clone()
    }

    /// Get a reference to the extracted plain text.
    pub fn plain_text_ref(&self) -> &str {
        &self.text
    }

    /// Convert to markdown (basic: paragraphs separated by blank lines).
    pub fn to_markdown(&self) -> String {
        let mut result = String::new();
        let mut prev_empty = false;

        for line in self.text.lines() {
            let trimmed = line.trim();
            if trimmed.is_empty() {
                if !prev_empty {
                    result.push('\n');
                }
                prev_empty = true;
            } else {
                result.push_str(trimmed);
                result.push_str("\n\n");
                prev_empty = false;
            }
        }

        result
    }
}

fn clx_size_zero_or_oob(clx_size: u32, clx_start: usize, stream_len: usize) -> bool {
    clx_size == 0 || clx_start + clx_size as usize > stream_len + 1024 // allow some slack
}

impl office_core::OfficeDocument for DocDocument {
    fn plain_text(&self) -> String {
        self.plain_text()
    }

    fn to_markdown(&self) -> String {
        self.to_markdown()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn markdown_double_spacing() {
        let doc = DocDocument {
            images: Vec::new(),
            text: "First paragraph\nSecond paragraph\n\nAfter gap".into(),
        };
        let md = doc.to_markdown();
        assert!(md.contains("First paragraph\n\n"));
        assert!(md.contains("Second paragraph\n\n"));
        assert!(md.contains("After gap\n\n"));
    }

    #[test]
    fn plain_text_access() {
        let doc = DocDocument {
            images: Vec::new(),
            text: "Hello World".into(),
        };
        assert_eq!(doc.plain_text(), "Hello World");
    }
}
