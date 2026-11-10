"""08_limits_strict_damage: strict mode refuses a damaged archive; a lenient open reports the damage.

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


def markdown_to_bytes(markdown: str, fmt: str) -> bytes:
    # create_from_markdown writes to a path, so the file lives in the temp dir
    # just long enough to read its bytes back.
    fd, path = tempfile.mkstemp(suffix=f".{fmt}")
    os.close(fd)
    try:
        create_from_markdown(markdown, fmt, path)
        with open(path, "rb") as f:
            return f.read()
    finally:
        os.unlink(path)


def build_docx_bytes(title: str = "Quarterly Report") -> bytes:
    return markdown_to_bytes(MARKDOWN.replace("Quarterly Report", title), "docx")


PART = "word/_rels/document.xml.rels"


def damage_part(data: bytes) -> bytes:
    # Overwrite the start of one part's compressed data. A first deflate byte of
    # 0xFF declares a reserved block type, so every inflater rejects the part,
    # and the central directory stays intact. The part is found by name, not by
    # offset, so the recipe does not depend on how big the other parts are.
    name = PART.encode()
    at = data.find(name)  # the first hit is the part's local header
    check(at > 0, "the part to damage is missing from the archive")
    # The extra-field length sits in the two bytes just before the name.
    start = at + len(name) + int.from_bytes(data[at - 2 : at], "little")
    return data[:start] + b"\xff" * 16 + data[start + 16 :]


def check(condition: bool, message: str) -> None:
    if not condition:
        print(f"FAIL: {message}", file=sys.stderr)
        sys.exit(1)


def has(text: str, expected: str) -> None:
    check(expected in text, f'"{expected}" is missing from the text')


def main() -> None:
    print("step 1: build a DOCX in memory")
    good = build_docx_bytes()

    print(f"step 2: damage one part: overwrite 16 bytes of {PART}")
    damaged = damage_part(good)

    print("step 3: open it in strict mode")
    try:
        Document.from_bytes(damaged, "docx", options=ParseOptions(strict=True))
    except OfficeOxideError as exc:
        print(f"  error: {exc}")
    else:
        check(False, "strict mode accepted a damaged part")

    print("step 4: open it leniently")
    with Document.from_bytes(damaged, "docx") as doc:
        kinds = [d["kind"] for d in doc.damage()]
        text = doc.plain_text()
    print(f"  damage: {', '.join(kinds)}")
    check("unreadable_part" in kinds, "the damage report does not name an unreadable part")
    has(text, "Quarterly Report")

    print("step 5: truncate the archive to its first half and open it both ways")
    half = good[: len(good) // 2]
    for message, options in (
        ("strict mode accepted a truncated archive", ParseOptions(strict=True)),
        ("a lenient open accepted a truncated archive", None),
    ):
        try:
            Document.from_bytes(half, "docx", options=options)
        except OfficeOxideError as exc:
            print(f"  error: {exc}")
        else:
            check(False, message)

    print("step 6: a limit applies to one document only")
    try:
        Document.from_bytes(good, "docx", options=ParseOptions(part_bytes=100))
    except OfficeOxideError as exc:
        print(f"  error: {exc}")
    else:
        check(False, "the part_bytes limit was not applied")
    with Document.from_bytes(good, "docx") as doc:  # the default limits still accept it
        has(doc.plain_text(), "Quarterly Report")

    print("OK 08_limits_strict_damage")


if __name__ == "__main__":
    main()
