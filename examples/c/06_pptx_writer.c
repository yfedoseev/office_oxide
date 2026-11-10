/* 06_pptx_writer: write a presentation with a size, slides, titles, text and an image, then read it back.
 *
 * Run: make -C examples/c build/06_pptx_writer && examples/c/build/06_pptx_writer */
#include "common.h"

int main(void) {
    printf("step 1: build a PPTX in memory\n");
    size_t len = 0;
    uint8_t *data = build_pptx_bytes(&len); /* see common.h for the slide-by-slide calls */

    printf("step 2: open it from bytes\n");
    int err = 0;
    OfficeDocumentHandle *doc = open_bytes(data, len, "pptx");
    char *text = take_str(office_document_plain_text(doc, &err), err, "plain text");
    char *ir_json = take_str(office_document_to_ir_json(doc, &err), err, "ir json");
    int slides = count_sections(ir_json);
    /* Plain Markdown drops images; the _with_images variant keeps them inline. */
    char *with_images = take_str(office_document_to_markdown_with_images(doc, &err), err, "markdown with images");
    printf("  format: %s\n  chars: %zu\n  slides: %d\n", office_document_format(doc), strlen(text), slides);

    printf("step 3: check the slides\n");
    check(slides == 2, "expected two slides");
    const char *expected[] = {"Welcome", "Quarterly Report", "Results", "Revenue grew in every region."};
    for (size_t i = 0; i < sizeof expected / sizeof *expected; i++) has(text, expected[i]);
    check(strstr(text, "Welcome") < strstr(text, "Results"), "the slides are out of order");
    check(strstr(with_images, "[image-base64:") != NULL, "the slide image is missing");

    office_oxide_free_string(text);
    office_oxide_free_string(ir_json);
    office_oxide_free_string(with_images);
    office_document_free(doc);
    office_oxide_free_bytes(data, len);
    printf("OK 06_pptx_writer\n");
    return 0;
}
