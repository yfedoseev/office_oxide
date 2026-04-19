use std::io::Cursor;
use std::path::PathBuf;

use pyo3::prelude::*;

use crate::Document;
use crate::edit::EditableDocument;
use crate::error::OfficeError;
use crate::format::DocumentFormat;

pyo3::create_exception!(office_oxide, OfficeOxideError, pyo3::exceptions::PyException);

impl From<OfficeError> for PyErr {
    fn from(e: OfficeError) -> PyErr {
        OfficeOxideError::new_err(e.to_string())
    }
}

/// A parsed Office document (DOCX, XLSX, PPTX, DOC, XLS, or PPT).
///
/// Supports use as a context manager:
///
///     with Document.open("report.docx") as doc:
///         print(doc.plain_text())
#[pyclass(name = "Document", module = "office_oxide")]
struct PyDocument {
    inner: Option<Document>,
    source: Option<String>,
}

impl PyDocument {
    fn get(&self) -> PyResult<&Document> {
        self.inner
            .as_ref()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("Document is closed"))
    }
}

#[pymethods]
impl PyDocument {
    /// Open a document from a file path (accepts str or os.PathLike).
    ///
    /// Format is detected from the extension; magic-byte sniffing corrects
    /// mismatched extensions. Raises OfficeOxideError on parse failure.
    #[staticmethod]
    #[pyo3(signature = (path, /))]
    fn open(path: PathBuf) -> PyResult<Self> {
        let source = path.display().to_string();
        let inner = Document::open(&path)?;
        Ok(PyDocument {
            inner: Some(inner),
            source: Some(source),
        })
    }

    /// Open a document from raw bytes with an explicit format.
    ///
    /// `format` is one of: "docx", "xlsx", "pptx", "doc", "xls", "ppt".
    #[staticmethod]
    #[pyo3(signature = (data, format, /))]
    fn from_bytes(data: &[u8], format: &str) -> PyResult<Self> {
        let fmt = DocumentFormat::from_extension(format)
            .ok_or_else(|| OfficeOxideError::new_err(format!("unsupported format: {format}")))?;
        let cursor = Cursor::new(data.to_vec());
        let inner = Document::from_reader(cursor, fmt)?;
        Ok(PyDocument {
            inner: Some(inner),
            source: None,
        })
    }

    /// The format as a short string ("docx", "xlsx", …).
    #[getter]
    fn format(&self) -> PyResult<&'static str> {
        Ok(match self.get()?.format() {
            DocumentFormat::Docx => "docx",
            DocumentFormat::Xlsx => "xlsx",
            DocumentFormat::Pptx => "pptx",
            DocumentFormat::Doc => "doc",
            DocumentFormat::Xls => "xls",
            DocumentFormat::Ppt => "ppt",
        })
    }

    /// Back-compat alias for `format`.
    fn format_name(&self) -> PyResult<&'static str> {
        self.format()
    }

    /// Extract plain text from the document.
    fn plain_text(&self) -> PyResult<String> {
        Ok(self.get()?.plain_text())
    }

    /// Convert the document to Markdown.
    fn to_markdown(&self) -> PyResult<String> {
        Ok(self.get()?.to_markdown())
    }

    /// Convert the document to an HTML fragment.
    fn to_html(&self) -> PyResult<String> {
        Ok(self.get()?.to_html())
    }

    /// Convert the document to a format-agnostic intermediate representation
    /// (nested dicts/lists).
    fn to_ir<'py>(&self, py: Python<'py>) -> PyResult<Py<PyAny>> {
        let doc_ir = self.get()?.to_ir();
        let json_str = serde_json::to_string(&doc_ir)
            .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(e.to_string()))?;
        let json_module = py.import("json")?;
        let result = json_module.call_method1("loads", (json_str,))?;
        Ok(result.unbind())
    }

    /// Save/convert the document to a file. Legacy formats are converted to OOXML.
    ///
    /// Example: doc.save_as("output.docx") converts DOC → DOCX.
    #[pyo3(signature = (path, /))]
    fn save_as(&self, path: PathBuf) -> PyResult<()> {
        self.get()?
            .save_as(&path)
            .map_err(|e| pyo3::exceptions::PyIOError::new_err(e.to_string()))
    }

    /// Release resources. The document becomes unusable after this.
    fn close(&mut self) {
        self.inner = None;
    }

    fn __enter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __exit__(
        &mut self,
        _exc_type: Option<Py<PyAny>>,
        _exc_val: Option<Py<PyAny>>,
        _exc_tb: Option<Py<PyAny>>,
    ) -> bool {
        self.close();
        false
    }

    fn __repr__(&self) -> String {
        match (&self.inner, &self.source) {
            (Some(d), Some(s)) => {
                format!("<Document format={:?} source={:?}>", d.format(), s)
            },
            (Some(d), None) => format!("<Document format={:?} from bytes>", d.format()),
            (None, _) => "<Document closed>".into(),
        }
    }
}

/// An editable document that supports text replacement and saving.
#[pyclass(name = "EditableDocument", module = "office_oxide")]
struct PyEditable {
    inner: Option<EditableDocument>,
}

impl PyEditable {
    fn get(&self) -> PyResult<&EditableDocument> {
        self.inner
            .as_ref()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("EditableDocument is closed"))
    }

    fn get_mut(&mut self) -> PyResult<&mut EditableDocument> {
        self.inner
            .as_mut()
            .ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("EditableDocument is closed"))
    }
}

#[pymethods]
impl PyEditable {
    /// Open a document for editing. Supports DOCX, XLSX, PPTX.
    #[staticmethod]
    #[pyo3(signature = (path, /))]
    fn open(path: PathBuf) -> PyResult<Self> {
        let inner = EditableDocument::open(&path)?;
        Ok(PyEditable { inner: Some(inner) })
    }

    /// Replace every occurrence of `find` with `replace` in text content.
    /// Returns the number of replacements.
    #[pyo3(signature = (find, replace, /))]
    fn replace_text(&mut self, find: &str, replace: &str) -> PyResult<usize> {
        Ok(self.get_mut()?.replace_text(find, replace))
    }

    /// Set a cell value in an XLSX document.
    ///
    /// `value` may be None (empty), str, bool, int, or float.
    #[pyo3(signature = (sheet_index, cell_ref, value, /))]
    fn set_cell(
        &mut self,
        sheet_index: usize,
        cell_ref: &str,
        value: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        use crate::xlsx::edit::CellValue;
        let cv = if value.is_none() {
            CellValue::Empty
        } else if let Ok(b) = value.extract::<bool>() {
            CellValue::Boolean(b)
        } else if let Ok(s) = value.extract::<String>() {
            CellValue::String(s)
        } else if let Ok(f) = value.extract::<f64>() {
            CellValue::Number(f)
        } else {
            return Err(pyo3::exceptions::PyTypeError::new_err(
                "value must be None, str, bool, int, or float",
            ));
        };
        self.get_mut()?.set_cell(sheet_index, cell_ref, cv)?;
        Ok(())
    }

    /// Save the edited document to a file.
    #[pyo3(signature = (path, /))]
    fn save(&self, path: PathBuf) -> PyResult<()> {
        self.get()?.save(&path)?;
        Ok(())
    }

    fn close(&mut self) {
        self.inner = None;
    }

    fn __enter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __exit__(
        &mut self,
        _exc_type: Option<Py<PyAny>>,
        _exc_val: Option<Py<PyAny>>,
        _exc_tb: Option<Py<PyAny>>,
    ) -> bool {
        self.close();
        false
    }
}

// ---------------------------------------------------------------------------
// Module-level convenience functions
// ---------------------------------------------------------------------------

/// Extract plain text from a file path.
#[pyfunction]
#[pyo3(signature = (path, /))]
fn extract_text(path: PathBuf) -> PyResult<String> {
    Ok(crate::extract_text(&path)?)
}

/// Convert a file to markdown.
#[pyfunction]
#[pyo3(signature = (path, /))]
fn to_markdown(path: PathBuf) -> PyResult<String> {
    Ok(crate::to_markdown(&path)?)
}

/// Convert a file to HTML.
#[pyfunction]
#[pyo3(signature = (path, /))]
fn to_html(path: PathBuf) -> PyResult<String> {
    Ok(crate::to_html(&path)?)
}

/// Library version (matches the Rust crate version).
#[pyfunction]
fn version() -> &'static str {
    crate::VERSION
}

/// Python module entry point.
#[pymodule]
fn _native(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<PyDocument>()?;
    m.add_class::<PyEditable>()?;
    m.add("OfficeOxideError", m.py().get_type::<OfficeOxideError>())?;
    m.add("__version__", crate::VERSION)?;
    m.add_function(wrap_pyfunction!(extract_text, m)?)?;
    m.add_function(wrap_pyfunction!(to_markdown, m)?)?;
    m.add_function(wrap_pyfunction!(to_html, m)?)?;
    m.add_function(wrap_pyfunction!(version, m)?)?;
    Ok(())
}
