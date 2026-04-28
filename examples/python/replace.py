"""Fill a DOCX/PPTX template by replacing placeholders."""
import sys
from office_oxide import EditableDocument


def main(src: str, dst: str) -> None:
    with EditableDocument.open(src) as ed:
        n = ed.replace_text("{{NAME}}", "Alice")
        n += ed.replace_text("{{DATE}}", "2026-04-18")
        print(f"replacements: {n}")
        ed.save(dst)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit("usage: replace.py <template> <output>")
    main(sys.argv[1], sys.argv[2])
