use std::io::Cursor;
use std::path::PathBuf;

use pyo3::prelude::*;

use crate::Document;
use crate::edit::EditableDocument;
use crate::error::OfficeError;
use crate::format::DocumentFormat;

/// Turn a writer's "did the value actually land" flag into a Python error.
///
/// The Rust API returns `bool` from these calls specifically so a silent
/// no-op is detectable — an off-by-one sheet or row index used to discard
/// every value it wrote while reporting success. Every binding dropped that
/// flag, reintroducing the regression the Rust side was fixed for.
fn require_written(ok: bool, what: &str) -> PyResult<()> {
    if ok {
        Ok(())
    } else {
        Err(pyo3::exceptions::PyIndexError::new_err(format!(
            "{what}: the target is out of range, so nothing was written"
        )))
    }
}

pyo3::create_exception!(office_oxide, OfficeOxideError, pyo3::exceptions::PyException);

/// The largest magnitude up to which every integer is exactly an `f64`.
const MAX_EXACT_INT: i64 = 1 << 53;

/// Convert a Python `int` to the `f64` an Excel number is stored as,
/// refusing one that would silently change value.
///
/// Excel stores every number as an IEEE-754 double. An `int` beyond 2**53
/// (or beyond `i64`) cannot be held exactly: `2**53 + 1` was written as
/// `2**53`, so an account or order number came back different. Raising
/// tells the caller to store it as text instead.
fn exact_int(value: &Bound<'_, PyAny>) -> PyResult<f64> {
    let too_big = || {
        pyo3::exceptions::PyValueError::new_err(format!(
            "integer {} cannot be stored exactly as an Excel number (a double is exact only \
             up to 2**53); write it as a string instead",
            value.str().map(|s| s.to_string()).unwrap_or_default()
        ))
    };
    let i: i64 = value.extract().map_err(|_| too_big())?;
    if i.unsigned_abs() > MAX_EXACT_INT as u64 {
        return Err(too_big());
    }
    Ok(i as f64)
}

/// The value kinds every `set_cell` accepts, decoded once.
enum PyCell {
    Empty,
    Bool(bool),
    Str(String),
    Number(f64),
}

/// Decode a Python cell value. `bool` is tested first: it is a subclass of
/// `int`, and pyo3's float extraction accepts it through `__index__`, so
/// `True` checked after the number arms became the number 1.
fn py_cell(value: &Bound<'_, PyAny>) -> PyResult<Option<PyCell>> {
    use pyo3::types::{PyBool, PyFloat, PyInt, PyString};
    Ok(Some(if value.is_none() {
        PyCell::Empty
    } else if value.is_instance_of::<PyBool>() {
        PyCell::Bool(value.extract()?)
    } else if value.is_instance_of::<PyString>() {
        PyCell::Str(value.extract()?)
    } else if value.is_instance_of::<PyInt>() {
        PyCell::Number(exact_int(value)?)
    } else if value.is_instance_of::<PyFloat>() {
        PyCell::Number(value.extract()?)
    } else {
        return Ok(None);
    }))
}

/// Decode a writer cell value; anything that is not None/bool/str/int/float
/// is written as its `str()`.
fn writer_cell_data(value: &Bound<'_, PyAny>) -> PyResult<crate::xlsx::write::CellData> {
    use crate::xlsx::write::CellData;
    Ok(match py_cell(value)? {
        Some(PyCell::Empty) => CellData::Empty,
        Some(PyCell::Bool(b)) => CellData::Boolean(b),
        Some(PyCell::Str(s)) => CellData::String(s),
        Some(PyCell::Number(n)) => CellData::Number(n),
        None => CellData::String(value.str()?.to_string()),
    })
}

impl From<OfficeError> for PyErr {
    fn from(e: OfficeError) -> PyErr {
        OfficeOxideError::new_err(e.to_string())
    }
}

/// A parsed Office document (DOCX, XLSX, PPTX, DOC, XLS, or PPT).
///
/// Supports use as a context manager:
///
/// ```text
/// with Document.open("report.docx") as doc:
///     print(doc.plain_text())
/// ```
///
/// (The block is marked `text` because it is Python shown to Python users;
/// an indented block here is collected as a Rust doctest and fails to
/// compile under `cargo test --all-features --doc`.)
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

    /// Convert to markdown, embedding each image inline as
    /// `[image-base64:<data>]` at its position in the document flow.
    ///
    /// Images are otherwise dropped from markdown entirely, which loses
    /// both their content and their position.
    fn to_markdown_with_images(&self) -> PyResult<String> {
        use crate::ir_render::{ImageEmbed, MarkdownOptions};
        Ok(self.get()?.to_markdown_with(MarkdownOptions {
            image_embed: ImageEmbed::Base64,
        }))
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

    /// Convert the document to a format-agnostic intermediate representation as a JSON string.
    fn to_ir_json(&self) -> PyResult<String> {
        let doc_ir = self.get()?.to_ir();
        serde_json::to_string(&doc_ir)
            .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(e.to_string()))
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

    /// Open a document for editing from raw bytes with an explicit format
    /// ("docx", "xlsx" or "pptx"), without a temporary file.
    #[staticmethod]
    #[pyo3(signature = (data, format, /))]
    fn from_bytes(data: &[u8], format: &str) -> PyResult<Self> {
        let fmt = DocumentFormat::from_extension(format)
            .ok_or_else(|| OfficeOxideError::new_err(format!("unsupported format: {format}")))?;
        let inner = EditableDocument::from_reader(Cursor::new(data.to_vec()), fmt)?;
        Ok(PyEditable { inner: Some(inner) })
    }

    /// Replace every occurrence of `find` with `replace` in text content.
    /// Returns the number of replacements.
    #[pyo3(signature = (find, replace, /))]
    fn replace_text(&mut self, find: &str, replace: &str) -> PyResult<usize> {
        self.get_mut()?
            .replace_text(find, replace)
            .map_err(|e| pyo3::exceptions::PyValueError::new_err(e.to_string()))
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
        let cv = match py_cell(value)? {
            Some(PyCell::Empty) => CellValue::Empty,
            Some(PyCell::Bool(b)) => CellValue::Boolean(b),
            Some(PyCell::Str(s)) => CellValue::String(s),
            Some(PyCell::Number(n)) => CellValue::Number(n),
            None => {
                return Err(pyo3::exceptions::PyTypeError::new_err(
                    "value must be None, str, bool, int, or float",
                ));
            },
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

    /// Serialize the edited document to bytes.
    fn to_bytes<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, pyo3::types::PyBytes>> {
        let mut cursor = Cursor::new(Vec::new());
        self.get()?.write_to(&mut cursor)?;
        Ok(pyo3::types::PyBytes::new(py, &cursor.into_inner()))
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

/// Create an Office document from Markdown text.
///
/// `format` must be one of `"docx"`, `"xlsx"`, or `"pptx"`.
#[pyfunction]
#[pyo3(signature = (markdown, format, path, /))]
fn create_from_markdown(markdown: &str, format: &str, path: PathBuf) -> PyResult<()> {
    let fmt = crate::format::DocumentFormat::from_extension(format)
        .ok_or_else(|| OfficeOxideError::new_err(format!("unsupported format: {format}")))?;
    crate::create::create_from_markdown(markdown, fmt, &path)?;
    Ok(())
}

/// Library version (matches the Rust crate version).
#[pyfunction]
fn version() -> &'static str {
    crate::VERSION
}

// ─── XlsxWriter ─────────────────────────────────────────────────────────────

#[pyclass(name = "XlsxWriter", module = "office_oxide")]
struct PyXlsxWriter {
    writer: crate::xlsx::write::XlsxWriter,
}

#[pymethods]
impl PyXlsxWriter {
    #[new]
    fn new() -> Self {
        Self {
            writer: crate::xlsx::write::XlsxWriter::new(),
        }
    }

    /// Add a worksheet; returns its 0-based index.
    fn add_sheet(&mut self, name: &str) -> usize {
        self.writer.add_sheet_get_index(name)
    }

    /// Set a cell value (str, float, int, bool, or None).
    ///
    /// A `str` is always text, even one starting with `=`; use
    /// `set_formula` for a formula. An `int` beyond 2**53 raises
    /// `ValueError`: Excel stores numbers as doubles and would change it.
    fn set_cell(
        &mut self,
        sheet: usize,
        row: usize,
        col: usize,
        value: &pyo3::Bound<'_, pyo3::PyAny>,
    ) -> PyResult<()> {
        let data = writer_cell_data(value)?;
        require_written(self.writer.sheet_set_cell(sheet, row, col, data), "set_cell")
    }

    /// Set a formula cell, e.g. `set_formula(0, 3, 1, "SUM(B1:B3)")`. A
    /// leading `=` is accepted and dropped.
    fn set_formula(&mut self, sheet: usize, row: usize, col: usize, formula: &str) -> PyResult<()> {
        let data = crate::xlsx::write::CellData::Formula(formula.to_string());
        require_written(self.writer.sheet_set_cell(sheet, row, col, data), "set_formula")
    }

    /// Set a cell with styling. bg_color is a 6-char hex string or None
    /// (the default, as the type stub always said).
    #[pyo3(signature = (sheet, row, col, value, bold, bg_color=None))]
    fn set_cell_styled(
        &mut self,
        sheet: usize,
        row: usize,
        col: usize,
        value: &pyo3::Bound<'_, pyo3::PyAny>,
        bold: bool,
        bg_color: Option<&str>,
    ) -> PyResult<()> {
        use crate::xlsx::write::CellStyle;
        // Booleans had no arm here at all: `True` became the number 1.
        let data = writer_cell_data(value)?;
        let mut style = CellStyle::new();
        if bold {
            style = style.bold();
        }
        if let Some(bg) = bg_color {
            style = style.background(bg.to_string());
        }
        require_written(
            self.writer
                .sheet_set_cell_styled(sheet, row, col, data, style),
            "set_cell_styled",
        )
    }

    /// Merge a rectangular range. row_span and col_span must be >= 1.
    fn merge_cells(
        &mut self,
        sheet: usize,
        row: usize,
        col: usize,
        row_span: usize,
        col_span: usize,
    ) {
        self.writer
            .sheet_merge_cells(sheet, row, col, row_span, col_span);
    }

    /// Set column width in Excel character units (e.g. 20.0).
    fn set_column_width(&mut self, sheet: usize, col: usize, width: f64) {
        self.writer.sheet_set_column_width(sheet, col, width);
    }

    /// Save to file.
    fn save(&self, path: PathBuf) -> PyResult<()> {
        self.writer
            .save(&path)
            .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(e.to_string()))
    }

    /// Serialize to bytes.
    fn to_bytes<'py>(
        &self,
        py: pyo3::Python<'py>,
    ) -> PyResult<pyo3::Bound<'py, pyo3::types::PyBytes>> {
        let mut cursor = std::io::Cursor::new(Vec::new());
        self.writer
            .write_to(&mut cursor)
            .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(e.to_string()))?;
        Ok(pyo3::types::PyBytes::new(py, &cursor.into_inner()))
    }
}

// ─── PptxWriter ──────────────────────────────────────────────────────────────

#[pyclass(name = "PptxWriter", module = "office_oxide")]
struct PyPptxWriter {
    writer: crate::pptx::write::PptxWriter,
}

#[pymethods]
impl PyPptxWriter {
    #[new]
    fn new() -> Self {
        Self {
            writer: crate::pptx::write::PptxWriter::new(),
        }
    }

    /// Override canvas size. 914400 EMU = 1 inch.
    fn set_presentation_size(&mut self, cx: u64, cy: u64) {
        self.writer.set_presentation_size(cx, cy);
    }

    /// Add a slide; returns its 0-based index.
    fn add_slide(&mut self) -> usize {
        self.writer.add_slide_get_index()
    }

    /// Set the slide title.
    fn set_slide_title(&mut self, slide: usize, title: &str) -> PyResult<()> {
        require_written(self.writer.slide_set_title(slide, title), "set_slide_title")
    }

    /// Add a plain text paragraph to the slide body.
    fn add_slide_text(&mut self, slide: usize, text: &str) -> PyResult<()> {
        require_written(self.writer.slide_add_text(slide, text), "add_slide_text")
    }

    /// Embed an image. format: "png" | "jpeg" | "gif". x,y,cx,cy in EMU.
    fn add_slide_image(
        &mut self,
        slide: usize,
        data: &[u8],
        format: &str,
        x: i64,
        y: i64,
        cx: u64,
        cy: u64,
    ) -> PyResult<()> {
        let fmt = match format.to_ascii_lowercase().as_str() {
            "png" => crate::ir::ImageFormat::Png,
            "jpeg" | "jpg" => crate::ir::ImageFormat::Jpeg,
            "gif" => crate::ir::ImageFormat::Gif,
            other => {
                return Err(pyo3::exceptions::PyValueError::new_err(format!(
                    "unsupported image format: {other}"
                )));
            },
        };
        require_written(
            self.writer
                .slide_add_image(slide, data.to_vec(), fmt, x, y, cx, cy),
            "add_slide_image",
        )
    }

    /// Save to file.
    fn save(&self, path: PathBuf) -> PyResult<()> {
        self.writer
            .save(&path)
            .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(e.to_string()))
    }

    /// Serialize to bytes.
    fn to_bytes<'py>(
        &self,
        py: pyo3::Python<'py>,
    ) -> PyResult<pyo3::Bound<'py, pyo3::types::PyBytes>> {
        let mut cursor = std::io::Cursor::new(Vec::new());
        self.writer
            .write_to(&mut cursor)
            .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(e.to_string()))?;
        Ok(pyo3::types::PyBytes::new(py, &cursor.into_inner()))
    }
}

/// Python module entry point.
#[pymodule]
fn _native(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<PyDocument>()?;
    m.add_class::<PyEditable>()?;
    m.add_class::<PyXlsxWriter>()?;
    m.add_class::<PyPptxWriter>()?;
    m.add("OfficeOxideError", m.py().get_type::<OfficeOxideError>())?;
    m.add("__version__", crate::VERSION)?;
    m.add_function(wrap_pyfunction!(extract_text, m)?)?;
    m.add_function(wrap_pyfunction!(to_markdown, m)?)?;
    m.add_function(wrap_pyfunction!(to_html, m)?)?;
    m.add_function(wrap_pyfunction!(create_from_markdown, m)?)?;
    m.add_function(wrap_pyfunction!(version, m)?)?;
    Ok(())
}
