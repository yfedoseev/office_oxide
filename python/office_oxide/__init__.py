"""office_oxide — fast Office document processing (DOCX, XLSX, PPTX, DOC, XLS, PPT).

Quick start::

    from office_oxide import Document

    with Document.open("report.docx") as doc:
        print(doc.plain_text())
        print(doc.to_markdown())

Editing::

    from office_oxide import EditableDocument

    with EditableDocument.open("report.docx") as ed:
        ed.replace_text("{{NAME}}", "Alice")
        ed.save("out.docx")
"""

from office_oxide._native import (
    Document,
    EditableDocument,
    OfficeOxideError,
    extract_text,
    to_html,
    to_markdown,
    version,
)

__version__ = version()

__all__ = [
    "Document",
    "EditableDocument",
    "OfficeOxideError",
    "__version__",
    "extract_text",
    "to_html",
    "to_markdown",
    "version",
]
