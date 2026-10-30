"""08_limits_strict_damage: strict options reject a damaged file; a lenient open reports damage or fails cleanly.

Run: python examples/python/08_limits_strict_damage/main.py
"""
import os
import sys
import tempfile

from office_oxide import Document, OfficeOxideError, ParseOptions, create_from_markdown


MARKDOWN = """\
# Quarterly Report

Revenue grew in every region.

| Region | Q1  | Q2  |
|--------|-----|-----|
| North  | 120 | 135 |
| South  | 95  | 110 |
"""


def build_docx_bytes(title: str = "Quarterly Report") -> bytes:
    # The only DOCX writer is create_from_markdown, which writes a path; the
    # file lives in the temp dir just long enough to read its bytes back.
    fd, path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        create_from_markdown(MARKDOWN.replace("Quarterly Report", title), "docx", path)
        with open(path, "rb") as f:
            return f.read()
    finally:
        os.unlink(path)

def check(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def main() -> None:
    print("step 1: build a DOCX in memory and keep only its first half")
    data = build_docx_bytes()
    damaged = data[: len(data) // 2]
    print(f"  chars: {len(damaged)}")

    print("step 2: open it in strict mode")
    try:
        Document.from_bytes(damaged, "docx", options=ParseOptions(strict=True))
    except OfficeOxideError as exc:
        print(f"  error: {exc}")
    else:
        check(False, "strict mode accepted a truncated archive")

    print("step 3: open it leniently")
    try:
        with Document.from_bytes(damaged, "docx") as doc:
            kinds = [d["kind"] for d in doc.damage()]
            print(f"  damage: {kinds}")
            check(len(kinds) > 0, "a lenient open succeeded but reported no damage")
    except OfficeOxideError as exc:
        print(f"  error: {exc}")

    print("step 4: a size limit applies to one document only")
    try:
        Document.from_bytes(data, "docx", options=ParseOptions(part_bytes=100))
    except OfficeOxideError as exc:
        print(f"  error: {exc}")
    else:
        check(False, "the part_bytes limit was not applied")
    with Document.from_bytes(data, "docx") as doc:  # the default limits still accept it
        check("Quarterly Report" in doc.plain_text(), "the default open lost the text")

    print("OK 08_limits_strict_damage")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"error: {exc!r}", file=sys.stderr)
        sys.exit(1)
