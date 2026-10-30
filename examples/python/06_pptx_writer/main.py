"""06_pptx_writer: build a two-slide presentation with titles, text and an image, save to bytes, read it back.

Run: python examples/python/06_pptx_writer/main.py
"""
import base64
import sys

from office_oxide import Document, PptxWriter


PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="
)


def build_pptx_bytes() -> bytes:
    w = PptxWriter()
    w.set_presentation_size(12192000, 6858000)  # 16:9, in EMU
    one = w.add_slide()
    w.set_slide_title(one, "Welcome")
    w.add_slide_text(one, "Quarterly Report")
    two = w.add_slide()
    w.set_slide_title(two, "Results")
    w.add_slide_text(two, "Revenue grew in every region.")
    w.add_slide_image(two, PNG, "png", 914400, 2743200, 1828800, 1828800)
    return w.to_bytes()

def check(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def main() -> None:
    print("step 1: build a PPTX in memory")
    data = build_pptx_bytes()
    print(f"  chars: {len(data)}")

    print("step 2: open it from bytes")
    with Document.from_bytes(data, "pptx") as doc:
        text = doc.plain_text()
        ir = doc.to_ir()
        with_images = doc.to_markdown_with_images()
        print(f"  format: {doc.format}")
        print(f"  chars: {len(text)}")

    print("step 3: check the slides")
    slides = len(ir["sections"])
    print(f"  slides: {slides}")
    check(slides == 2, "expected two slides")
    for expected in ("Welcome", "Quarterly Report", "Results", "Revenue grew in every region."):
        check(expected in text, f"{expected!r} is missing from the text")
    check(text.index("Welcome") < text.index("Results"), "the slides are out of order")
    # Plain Markdown drops images; the _with_images variant keeps them in place.
    check("[image-base64:" in with_images, "the slide image is missing")

    print("OK 06_pptx_writer")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"error: {exc!r}", file=sys.stderr)
        sys.exit(1)
