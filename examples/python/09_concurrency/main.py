"""09_concurrency: parse 32 documents on 8 threads with Python logging switched on.

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

WORKERS, DOCUMENTS, TIMEOUT_S = 8, 32, 60
log = logging.getLogger("example.concurrency")


def read_one(i: int, data: bytes) -> str:
    log.debug("document %d: parse start", i)
    with Document.from_bytes(data, "docx") as doc:
        text = doc.plain_text()
    log.debug("document %d: parse done", i)
    return text


def main() -> None:
    logging.basicConfig(level=logging.DEBUG, stream=sys.stdout, format="  log %(threadName)s %(message)s")
    print("step 1: build 32 DOCX files in memory")
    inputs = [build_docx_bytes(f"Quarterly Report {i}") for i in range(DOCUMENTS)]
    print(f"  documents: {DOCUMENTS}")

    print("step 2: parse them on 8 threads with logging on")
    print(f"  workers: {WORKERS}")
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futures = [pool.submit(read_one, i, d) for i, d in enumerate(inputs)]
        try:
            texts = [f.result(timeout=TIMEOUT_S) for f in futures]
        except TimeoutError:
            print("a parse did not finish: possible deadlock", file=sys.stderr)
            os._exit(2)  # a stuck worker would block a normal exit

    print("step 3: check every result")
    for i, text in enumerate(texts):
        check(f"Quarterly Report {i}" in text, f"document {i} came back with the wrong text")

    print("OK 09_concurrency")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"error: {exc!r}", file=sys.stderr)
        sys.exit(1)
