"""Load an XLSX and print its IR sections as simple tables."""
import sys
from office_oxide import Document


def main(path: str) -> None:
    with Document.open(path) as doc:
        ir = doc.to_ir()
        for i, section in enumerate(ir["sections"]):
            print(f"# sheet {i}: {section.get('title')}")
            for el in section["elements"]:
                if el["type"] == "table":
                    for row in el["rows"]:
                        cells = [c.get("text", "") for c in row["cells"]]
                        print("\t".join(cells))


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("usage: read_xlsx.py <file.xlsx>")
    main(sys.argv[1])
