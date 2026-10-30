"""09_concurrency: open and read 32 documents from 8 workers at once, with logging on.

Run: python examples/python/09_concurrency/main.py
"""
import logging
import os
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor, TimeoutError

from office_oxide import Document, create_from_markdown


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


WORKERS, DOCUMENTS, TIMEOUT_S = 8, 32, 60
log = logging.getLogger("example.concurrency")


def read_one(i: int, data: bytes) -> str:
    log.debug("document %d: parse start", i)
    with Document.from_bytes(data, "docx") as doc:
        text = doc.plain_text()
    log.debug("document %d: parse done", i)
    return text


def check(condition: bool, message: str) -> None:
    if not condition:
        print(f"FAIL: {message}", file=sys.stderr)
        sys.exit(1)


def has(text: str, expected: str) -> None:
    check(expected in text, f'"{expected}" is missing from the text')


def main() -> None:
    print("step 1: build 32 DOCX files in memory")
    inputs = [build_docx_bytes(f"Quarterly Report {i}") for i in range(DOCUMENTS)]
    print(f"  documents: {DOCUMENTS}")

    print("step 2: switch logging on")
    logging.basicConfig(level=logging.DEBUG, stream=sys.stderr, format="%(threadName)s %(message)s")

    print("step 3: parse them on 8 workers")
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futures = [pool.submit(read_one, i, data) for i, data in enumerate(inputs)]
        try:
            texts = [f.result(timeout=TIMEOUT_S) for f in futures]
        except TimeoutError:
            print("FAIL: a parse did not finish: possible deadlock", file=sys.stderr)
            os._exit(1)  # a stuck worker would block a normal exit
    print(f"  workers: {WORKERS}")

    print("step 4: check every result")
    check(len(texts) == DOCUMENTS, "not every document was read")
    for i, text in enumerate(texts):
        check(f"Quarterly Report {i}" in text, f"document {i} came back with the wrong text")

    print("OK 09_concurrency")


if __name__ == "__main__":
    main()
