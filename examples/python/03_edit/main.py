"""03_edit — Edit a document by replacing placeholder text.

Creates a DOCX from Markdown that contains {{NAME}} and {{DATE}},
opens it with EditableDocument, replaces both placeholders, saves to
a second temp file, then reads back and verifies. Exit 0 on success.
"""
import sys
import tempfile
import os

from office_oxide import Document, EditableDocument, create_from_markdown


MARKDOWN = """\
# Invoice

Dear {{NAME}},

Please find attached your invoice for services rendered on {{DATE}}.

## Summary

- Service: Office Oxide Pro License
- Amount: $499.00
- Due date: 30 days from {{DATE}}

Thank you for your business!
"""


def main() -> None:
    template_tmp = None
    output_tmp = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            template_tmp = f.name
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            output_tmp = f.name

        # Create template
        create_from_markdown(MARKDOWN, "docx", template_tmp)

        # Edit
        with EditableDocument.open(template_tmp) as ed:
            n1 = ed.replace_text("{{NAME}}", "Alice Smith")
            n2 = ed.replace_text("{{DATE}}", "2026-04-26")
            ed.save(output_tmp)

        print(f"Replacements: {{{{NAME}}}} x{n1}, {{{{DATE}}}} x{n2}")

        # Verify
        with Document.open(output_tmp) as doc:
            text = doc.plain_text()

        assert "Alice Smith" in text, "name replacement failed"
        assert "2026-04-26" in text, "date replacement failed"
        assert "{{NAME}}" not in text, "placeholder still present"
        assert "{{DATE}}" not in text, "placeholder still present"

        print("Edit verified successfully.")
        print("--- final text ---")
        print(text)
    finally:
        for p in [template_tmp, output_tmp]:
            if p and os.path.exists(p):
                os.unlink(p)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        sys.exit(1)
