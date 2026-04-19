"""Extract plain text and Markdown from an Office document."""
import sys
from office_oxide import Document


def main(path: str) -> None:
    with Document.open(path) as doc:
        print(f"format: {doc.format}")
        print("--- plain text ---")
        print(doc.plain_text())
        print("--- markdown ---")
        print(doc.to_markdown())


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("usage: extract.py <file>")
    main(sys.argv[1])
