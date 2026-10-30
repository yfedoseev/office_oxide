/* 06_pptx_writer: write a presentation with a size, slides, titles, text and an image.
 * Run: make -C examples/c run   (or build/06_pptx_writer after make -C examples/c) */
#include "common.h"

int main(void) {
    printf("step 1: build a PPTX in memory\n");
    size_t len = 0;
    uint8_t *bytes = build_pptx_bytes(&len); /* see common.h for the slide-by-slide calls */

    printf("step 2: check that a write to a missing slide is refused\n");
    OfficePptxWriterHandle *w = office_pptx_writer_new();
    check(w != NULL, "pptx writer");
    check(office_pptx_slide_set_title(w, 9, "ghost") == OFFICE_ERR_UNSUPPORTED,
          "a missing slide is an error, not a silent drop");
    office_pptx_writer_free(w);

    printf("step 3: open it from bytes\n");
    int err = 0;
    OfficeDocumentHandle *doc = open_bytes(bytes, len, "pptx");
    char *text = take_str(office_document_plain_text(doc, &err), err, "plain text");
    /* Plain Markdown drops images; the _with_images variant keeps them inline. */
    char *md = take_str(office_document_to_markdown_with_images(doc, &err), err, "markdown with images");
    printf("  format: %s\n  chars: %zu\n  slides: 2\n", office_document_format(doc), strlen(text));

    printf("step 4: check the slides\n");
    check_has(text, "Welcome", "slide 1 title is present");
    check_has(text, "Quarterly Report", "slide 1 text is present");
    check_has(text, "Results", "slide 2 title is present");
    check_has(text, "Revenue grew in every region.", "slide 2 text is present");
    check_has(md, "[image-base64:", "the image is present");

    office_oxide_free_string(text);
    office_oxide_free_string(md);
    office_document_free(doc);
    office_oxide_free_bytes(bytes, len);
    printf("OK 06_pptx_writer\n");
    return 0;
}
