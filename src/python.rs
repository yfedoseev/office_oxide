use std::io::Cursor;

use pyo3::prelude::*;

use crate::Document;
use crate::error::OfficeError;
use crate::format::DocumentFormat;

pyo3::create_exception!(office_oxide, OfficeOxideError, pyo3::exceptions::PyException);

impl From<OfficeError> for PyErr {
    fn from(e: OfficeError) -> PyErr {
        OfficeOxideError::new_err(e.to_string())
    }
}

/// A parsed Office document (DOCX, XLSX, or PPTX).
#[pyclass(name = "Document")]
struct PyDocument {
    inner: Document,
}

#[pymethods]
impl PyDocument {
    /// Open a document from a file path. Format is detected from the extension.
    #[staticmethod]
    fn open(path: &str) -> PyResult<Self> {
        let inner = Document::open(path)?;
        Ok(PyDocument { inner })
    }

    /// Open a document from raw bytes with an explicit format string ("docx", "xlsx", or "pptx").
    #[staticmethod]
    fn from_bytes(data: &[u8], format: &str) -> PyResult<Self> {
        let fmt = DocumentFormat::from_extension(format)
            .ok_or_else(|| OfficeOxideError::new_err(format!("unsupported format: {format}")))?;
        let cursor = Cursor::new(data.to_vec());
        let inner = Document::from_reader(cursor, fmt)?;
        Ok(PyDocument { inner })
    }

    /// Return the format name.
    fn format_name(&self) -> &str {
        match self.inner.format() {
            DocumentFormat::Docx => "docx",
            DocumentFormat::Xlsx => "xlsx",
            DocumentFormat::Pptx => "pptx",
            DocumentFormat::Doc => "doc",
            DocumentFormat::Xls => "xls",
            DocumentFormat::Ppt => "ppt",
        }
    }

    /// Extract plain text from the document.
    fn plain_text(&self) -> String {
        self.inner.plain_text()
    }

    /// Convert the document to markdown.
    fn to_markdown(&self) -> String {
        self.inner.to_markdown()
    }

    /// Convert the document to an HTML fragment.
    fn to_html(&self) -> String {
        self.inner.to_html()
    }

    /// Convert the document to a format-agnostic intermediate representation (nested dicts/lists).
    fn to_ir<'py>(&self, py: Python<'py>) -> PyResult<Py<PyAny>> {
        let doc_ir = self.inner.to_ir();
        let json_str = serde_json::to_string(&doc_ir)
            .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(e.to_string()))?;
        let json_module = py.import("json")?;
        let result = json_module.call_method1("loads", (json_str,))?;
        Ok(result.unbind())
    }

    /// Save/convert the document to a file. Legacy formats are converted to OOXML.
    /// Example: doc.save_as("output.docx") converts DOC → DOCX.
    fn save_as(&self, path: &str) -> PyResult<()> {
        self.inner
            .save_as(path)
            .map_err(|e| pyo3::exceptions::PyIOError::new_err(e.to_string()))
    }
}

// ---------------------------------------------------------------------------
// Module-level convenience functions
// ---------------------------------------------------------------------------

/// Extract plain text from a file path.
#[pyfunction]
fn extract_text(path: &str) -> PyResult<String> {
    Ok(crate::extract_text(path)?)
}

/// Convert a file to markdown.
#[pyfunction]
fn to_markdown(path: &str) -> PyResult<String> {
    Ok(crate::to_markdown(path)?)
}

#[pyfunction]
fn to_html(path: &str) -> PyResult<String> {
    Ok(crate::to_html(path)?)
}

/// Python module entry point.
#[pymodule]
fn _native(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<PyDocument>()?;
    m.add("OfficeOxideError", m.py().get_type::<OfficeOxideError>())?;
    m.add_function(wrap_pyfunction!(extract_text, m)?)?;
    m.add_function(wrap_pyfunction!(to_markdown, m)?)?;
    m.add_function(wrap_pyfunction!(to_html, m)?)?;
    Ok(())
}
