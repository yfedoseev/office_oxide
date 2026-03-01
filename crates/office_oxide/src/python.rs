use std::io::Cursor;

use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use crate::error::OfficeError;
use crate::format::DocumentFormat;
use crate::ir;
use crate::Document;

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
        let fmt = DocumentFormat::from_extension(format).ok_or_else(|| {
            OfficeOxideError::new_err(format!("unsupported format: {format}"))
        })?;
        let cursor = Cursor::new(data.to_vec());
        let inner = Document::from_reader(cursor, fmt)?;
        Ok(PyDocument { inner })
    }

    /// Return the format name: "docx", "xlsx", or "pptx".
    fn format_name(&self) -> &str {
        match self.inner.format() {
            DocumentFormat::Docx => "docx",
            DocumentFormat::Xlsx => "xlsx",
            DocumentFormat::Pptx => "pptx",
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

    /// Convert the document to a format-agnostic intermediate representation (nested dicts/lists).
    fn to_ir<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyDict>> {
        let doc_ir = self.inner.to_ir();
        ir_to_pydict(py, &doc_ir)
    }
}

// ---------------------------------------------------------------------------
// IR → Python dict conversion
// ---------------------------------------------------------------------------

fn ir_to_pydict<'py>(py: Python<'py>, doc: &ir::DocumentIR) -> PyResult<Bound<'py, PyDict>> {
    let dict = PyDict::new(py);

    // metadata
    let meta = PyDict::new(py);
    meta.set_item("format", format!("{:?}", doc.metadata.format))?;
    meta.set_item("title", doc.metadata.title.as_deref())?;
    dict.set_item("metadata", meta)?;

    // sections
    let sections = PyList::empty(py);
    for section in &doc.sections {
        sections.append(section_to_pydict(py, section)?)?;
    }
    dict.set_item("sections", sections)?;

    Ok(dict)
}

fn section_to_pydict<'py>(py: Python<'py>, section: &ir::Section) -> PyResult<Bound<'py, PyDict>> {
    let dict = PyDict::new(py);
    dict.set_item("title", section.title.as_deref())?;

    let elements = PyList::empty(py);
    for elem in &section.elements {
        elements.append(element_to_pyobj(py, elem)?)?;
    }
    dict.set_item("elements", elements)?;

    Ok(dict)
}

fn element_to_pyobj<'py>(py: Python<'py>, elem: &ir::Element) -> PyResult<Bound<'py, PyDict>> {
    let dict = PyDict::new(py);
    match elem {
        ir::Element::Heading(h) => {
            dict.set_item("type", "heading")?;
            dict.set_item("level", h.level)?;
            dict.set_item("content", inline_content_to_pylist(py, &h.content)?)?;
        }
        ir::Element::Paragraph(p) => {
            dict.set_item("type", "paragraph")?;
            dict.set_item("content", inline_content_to_pylist(py, &p.content)?)?;
        }
        ir::Element::Table(t) => {
            dict.set_item("type", "table")?;
            let rows = PyList::empty(py);
            for row in &t.rows {
                let row_dict = PyDict::new(py);
                row_dict.set_item("is_header", row.is_header)?;
                let cells = PyList::empty(py);
                for cell in &row.cells {
                    let cell_dict = PyDict::new(py);
                    cell_dict.set_item("col_span", cell.col_span)?;
                    cell_dict.set_item("row_span", cell.row_span)?;
                    let cell_elements = PyList::empty(py);
                    for e in &cell.content {
                        cell_elements.append(element_to_pyobj(py, e)?)?;
                    }
                    cell_dict.set_item("content", cell_elements)?;
                    cells.append(cell_dict)?;
                }
                row_dict.set_item("cells", cells)?;
                rows.append(row_dict)?;
            }
            dict.set_item("rows", rows)?;
        }
        ir::Element::List(l) => {
            dict.set_item("type", "list")?;
            dict.set_item("ordered", l.ordered)?;
            dict.set_item("items", list_items_to_pylist(py, &l.items)?)?;
        }
        ir::Element::Image(img) => {
            dict.set_item("type", "image")?;
            dict.set_item("alt_text", img.alt_text.as_deref())?;
        }
        ir::Element::ThematicBreak => {
            dict.set_item("type", "thematic_break")?;
        }
    }
    Ok(dict)
}

fn inline_content_to_pylist<'py>(
    py: Python<'py>,
    content: &[ir::InlineContent],
) -> PyResult<Bound<'py, PyList>> {
    let list = PyList::empty(py);
    for item in content {
        let dict = PyDict::new(py);
        match item {
            ir::InlineContent::Text(span) => {
                dict.set_item("type", "text")?;
                dict.set_item("text", &span.text)?;
                dict.set_item("bold", span.bold)?;
                dict.set_item("italic", span.italic)?;
                dict.set_item("strikethrough", span.strikethrough)?;
                dict.set_item("hyperlink", span.hyperlink.as_deref())?;
            }
            ir::InlineContent::LineBreak => {
                dict.set_item("type", "line_break")?;
            }
        }
        list.append(dict)?;
    }
    Ok(list)
}

fn list_items_to_pylist<'py>(
    py: Python<'py>,
    items: &[ir::ListItem],
) -> PyResult<Bound<'py, PyList>> {
    let list = PyList::empty(py);
    for item in items {
        let dict = PyDict::new(py);
        dict.set_item("content", inline_content_to_pylist(py, &item.content)?)?;
        if let Some(ref nested) = item.nested {
            let nested_dict = PyDict::new(py);
            nested_dict.set_item("ordered", nested.ordered)?;
            nested_dict.set_item("items", list_items_to_pylist(py, &nested.items)?)?;
            dict.set_item("nested", nested_dict)?;
        } else {
            dict.set_item("nested", py.None())?;
        }
        list.append(dict)?;
    }
    Ok(list)
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

/// Python module entry point.
#[pymodule]
fn _native(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<PyDocument>()?;
    m.add("OfficeOxideError", m.py().get_type::<OfficeOxideError>())?;
    m.add_function(wrap_pyfunction!(extract_text, m)?)?;
    m.add_function(wrap_pyfunction!(to_markdown, m)?)?;
    Ok(())
}
